/*
 * Bytes BCC Guard — OnMessageSend handler
 *
 * Blocks sending when more than 5 recipients are in To/Cc combined AND at least
 * one of those To/Cc recipients is external to epicsolutions.org.uk.
 * Bcc recipients are never inspected or counted — Office.js exposes To, Cc,
 * and Bcc as separate arrays, so there is no envelope-flattening problem here
 * the way there was with the Purview DLP approach.
 *
 * MAX_TO_CC_RECIPIENTS: number of To+Cc recipients allowed before blocking.
 * "In excess of 5" means this blocks at 6 or more — set to 5.
 */

const MAX_TO_CC_RECIPIENTS = 5;
const INTERNAL_DOMAIN = "epicsolutions.org.uk";

const BLOCK_MESSAGE =
  "This message was not delivered because it has more than " + MAX_TO_CC_RECIPIENTS +
  " recipients in the To/Cc fields, including at least one external address. " +
  "To protect recipient privacy, please resend using Bcc for the additional recipients instead of To or Cc. " +
  "If you believe this message should not have been blocked, please contact Bytes Computers, your IT support.";

function isExternal(emailAddress) {
  if (!emailAddress) return false;
  const domain = emailAddress.split("@")[1];
  if (!domain) return false;
  return domain.toLowerCase() !== INTERNAL_DOMAIN.toLowerCase();
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
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});
