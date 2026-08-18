/*
 * Bytes BCC Guard — OnMessageSend handler (NPVG)
 *
 * Blocks sending when more than 5 recipients are in To/Cc combined
 * AND at least one of those To/Cc recipients is external to this client's
 * internal domain(s). Bcc recipients are never inspected or counted.
 *
 * Supports multiple internal domains (some clients operate under more than
 * one brand/entity but share one staff directory) — an address only counts
 * as external if its domain matches none of the entries in INTERNAL_DOMAINS.
 */

const MAX_TO_CC_RECIPIENTS = 5;
const INTERNAL_DOMAINS = ["northparkvets.co.uk", "equusvets.co.uk"];

const BLOCK_MESSAGE =
  "This message was not delivered because it has more than " + MAX_TO_CC_RECIPIENTS +
  " recipients in the To/Cc fields, including at least one external address. " +
  "To protect recipient privacy, please resend using Bcc for the additional recipients instead of To or Cc. " +
  "If you believe this message should not have been blocked, please contact Bytes Computers, your IT support.";

function isExternal(emailAddress) {
  if (!emailAddress) return false;
  const domain = emailAddress.split("@")[1];
  if (!domain) return false;
  const domainLower = domain.toLowerCase();
  return !INTERNAL_DOMAINS.some(function (d) { return d === domainLower; });
}

function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  // Get To and Cc recipients. Bcc is intentionally never requested here.
  Promise.all([getRecipients(item.to), getRecipients(item.cc)])
    .then(function (results) {
      const toRecipients = results[0] || [];
      const ccRecipients = results[1] || [];

      // De-duplicate by email address in case the same person appears in both
      // To and Cc — we want a count of unique people, not raw field entries.
      const combined = toRecipients.concat(ccRecipients);
      const uniqueAddresses = {};
      combined.forEach(function (r) {
        if (r && r.emailAddress) {
          uniqueAddresses[r.emailAddress.toLowerCase()] = true;
        }
      });
      const uniqueList = Object.keys(uniqueAddresses);

      const hasExternal = uniqueList.some(function (addr) {
        return isExternal(addr);
      });

      const recipientCount = uniqueList.length;

      if (recipientCount > MAX_TO_CC_RECIPIENTS && hasExternal) {
        event.completed({
          allowEvent: false,
          errorMessage: BLOCK_MESSAGE
        });
        return;
      }

      // Passes the check — allow the send.
      event.completed({ allowEvent: true });
    })
    .catch(function (error) {
      // If something goes wrong reading recipients, fail open (allow the send)
      // rather than blocking legitimate mail on an add-in error.
      console.error("Bytes BCC Guard: error reading recipients, allowing send.", error);
      event.completed({ allowEvent: true });
    });
}

// Wraps the callback-based Office.js getAsync API in a Promise so it can be
// used with Promise.all above.
function getRecipients(recipientField) {
  return new Promise(function (resolve, reject) {
    recipientField.getAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(asyncResult.value);
      } else {
        reject(asyncResult.error);
      }
    });
  });
}

// Register the handler immediately and unconditionally. This is required for
// classic Outlook on Windows: Microsoft's own documentation confirms that on
// this platform, code inside Office.onReady() and Office.initialize does NOT
// run when an event handler is invoked via the JS-only runtime — so wrapping
// this call in Office.onReady alone left classic desktop Outlook with no
// handler ever registered, causing the send dialog to hang indefinitely on
// "taking longer than expected" rather than firing or erroring cleanly.
// Confirmed as the real cause of a live production hang on 11/08/2026.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

Office.onReady(function () {
  // Also register here for the browser-runtime clients (OWA, Mac, new
  // Outlook on Windows), where this wrapper was originally needed to fix a
  // similar hang during EPIC's initial testing. Registering the same name
  // twice is safe — Outlook simply keeps the latest registration.
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});
