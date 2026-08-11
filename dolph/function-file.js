/*
 * Bytes BCC Guard — OnMessageSend handler (DOLPH)
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
const INTERNAL_DOMAINS = ["dolphindevon.co.uk"];

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

Office.onReady(function () {
  // Required wrapper — without this, Outlook can invoke the handler before
  // it's actually associated, causing the send dialog to hang indefinitely
  // on "taking longer than expected" instead of firing cleanly.
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});
