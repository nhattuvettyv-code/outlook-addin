/* Outlook Smart Alert Add-in — Customer Name Validator
 *
 * Triggered on OnMessageSend. Extracts [Customer Name] from the subject,
 * looks it up in mapping.json, then validates every recipient's email domain
 * against the allowed domains/addresses for that customer.
 * Hard-blocks the send if any recipient does not match.
 */

Office.onReady(() => {
  // Register the event handler name used in manifest.xml
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});

// ── Helpers ────────────────────────────────────────────────────────────────

function getSubjectAsync(item) {
  return new Promise((resolve, reject) => {
    item.subject.getAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        reject(result.error);
      }
    });
  });
}

function getRecipientsAsync(field) {
  return new Promise((resolve, reject) => {
    field.getAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
      } else {
        reject(result.error);
      }
    });
  });
}

async function getAllRecipients(item) {
  const [to, cc, bcc] = await Promise.all([
    getRecipientsAsync(item.to),
    getRecipientsAsync(item.cc),
    getRecipientsAsync(item.bcc),
  ]);
  return [...to, ...cc, ...bcc];
}

function domainOf(email) {
  const parts = (email || "").toLowerCase().split("@");
  return parts.length === 2 ? parts[1] : "";
}

function isAllowed(email, rule) {
  const lowerEmail = (email || "").toLowerCase();
  const domain = domainOf(lowerEmail);

  const allowedDomains = (rule.domains || []).map(d => d.toLowerCase());
  const allowedAddresses = (rule.addresses || []).map(a => a.toLowerCase());

  return allowedDomains.includes(domain) || allowedAddresses.includes(lowerEmail);
}

function caseInsensitiveLookup(mapping, key) {
  const lowerKey = key.toLowerCase();
  for (const [k, v] of Object.entries(mapping)) {
    if (k.toLowerCase() === lowerKey) return v;
  }
  return null;
}

// ── Smart Alert handler ────────────────────────────────────────────────────

async function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  try {
    // 1. Read subject and extract [Customer Name]
    const subject = await getSubjectAsync(item);
    const bracketMatch = subject.match(/\[([^\]]+)\]/);

    if (!bracketMatch) {
      // No bracketed customer name — nothing to validate
      event.completed({ allowEvent: true });
      return;
    }

    const customerName = bracketMatch[1].trim();

    // 2. Load mapping (relative to commands.html location)
    let mapping;
    try {
      const response = await fetch("mapping.json");
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      mapping = await response.json();
    } catch (fetchErr) {
      // If mapping cannot be loaded, fail safe: block the send
      event.completed({
        allowEvent: false,
        errorMessage:
          "Send blocked: the customer validation mapping could not be loaded.\n" +
          "Please contact your IT administrator.\n\n" +
          `Detail: ${fetchErr.message}`,
      });
      return;
    }

    // 3. Look up the customer rule (case-insensitive)
    const rule = caseInsensitiveLookup(mapping, customerName);

    if (!rule) {
      event.completed({
        allowEvent: false,
        errorMessage:
          `Send blocked: the customer name "${customerName}" in the subject is not recognised.\n\n` +
          "Please check the subject line or update the customer mapping.",
      });
      return;
    }

    // 4. Get all recipients and validate
    const recipients = await getAllRecipients(item);

    if (recipients.length === 0) {
      // No recipients yet — let Outlook's own validation handle this
      event.completed({ allowEvent: true });
      return;
    }

    const mismatches = recipients.filter(
      r => !isAllowed(r.emailAddress, rule)
    );

    // 5. Block or allow
    if (mismatches.length > 0) {
      const recipientList = mismatches
        .map(r => `  • ${r.displayName ? r.displayName + " " : ""}<${r.emailAddress}>`)
        .join("\n");

      event.completed({
        allowEvent: false,
        errorMessage:
          `Send blocked: the following recipient(s) do not match the customer "${customerName}":\n\n` +
          `${recipientList}\n\n` +
          "Please correct the subject line or the recipient list before sending.",
      });
    } else {
      event.completed({ allowEvent: true });
    }
  } catch (err) {
    // Unexpected error — fail safe by blocking
    event.completed({
      allowEvent: false,
      errorMessage:
        "Send blocked: an unexpected error occurred during recipient validation.\n\n" +
        `Detail: ${err.message}`,
    });
  }
}
