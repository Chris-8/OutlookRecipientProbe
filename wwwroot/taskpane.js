/* global Office */

Office.onReady(() => {
  const button = document.getElementById("logRecipients");
  button?.addEventListener("click", logAllRecipients);
});

function getRecipientsAsync(collectionName) {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    const collection = item?.[collectionName];

    if (!collection || typeof collection.getAsync !== "function") {
      reject(new Error(`Recipient collection '${collectionName}' is unavailable.`));
      return;
    }

    collection.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve((result.value || []).map(normalizeRecipient));
      } else {
        reject(new Error(`${collectionName}.getAsync failed: ${result.error?.message || "unknown error"}`));
      }
    });
  });
}

function normalizeRecipient(recipient) {
  return {
    displayName: recipient.displayName ?? null,
    emailAddress: recipient.emailAddress ?? null,
    recipientType: recipient.recipientType ?? null
  };
}

async function logAllRecipients() {
  const output = document.getElementById("output");

  try {
    const [to, cc, bcc] = await Promise.all([
      getRecipientsAsync("to"),
      getRecipientsAsync("cc"),
      getRecipientsAsync("bcc")
    ]);

    const payload = {
      itemType: Office.context.mailbox.item?.itemType ?? null,
      itemClass: Office.context.mailbox.item?.itemClass ?? null,
      to,
      cc,
      bcc
    };

    console.log("Recipient probe payload:", payload);
    output.textContent = JSON.stringify(payload, null, 2);
  } catch (error) {
    console.error("Recipient probe failed:", error);
    output.textContent = String(error);
  }
}
