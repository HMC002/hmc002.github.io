/*
 * Inbox Desk Reply - Outlook add-in, event handler.
 *
 * Runs when a new message window opens. For a reply or reply-all it asks the
 * HMC reply writer (a Make scenario) for the reply text and places it at the
 * top of the reply. Nothing is sent.
 *
 * Classic Outlook on Windows loads this one file on its own, so it has no
 * imports, does not wait for Office.onReady, keeps to ES2016 syntax and uses
 * absolute URLs only.
 *
 * This file is served from a public address. It holds no reply writer
 * address, no key and no email address. The address and key come from the
 * mailbox's roaming settings (see settings.js). Bcc addresses come back from
 * the reply writer.
 */

var HMC = {
  SETTING_URL: "hmcReplyWriterUrl",
  SETTING_KEY: "hmcReplyWriterKey",
  NOTICE: "hmcReplyWriter",
  MAX_BODY_CHARS: 12000,
  MAX_PARAGRAPHS: 20,
  MAX_REPLY_CHARS: 6000,
  TIMEOUT_MS: 60000,
  PARAGRAPH_STYLE: "margin:0cm;font-size:11.0pt;font-family:Calibri,sans-serif"
};

/* ---------- Pure helpers (covered by test/launchevent.test.js) ---------- */

function hmcEscapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/* One Calibri 11 pt paragraph per string, each followed by a blank line. Same layout as the Make draft builder. */
function hmcParagraphsToHtml(paragraphs) {
  var open = "<p style='" + HMC.PARAGRAPH_STYLE + "'>";
  var parts = [];
  for (var i = 0; i < paragraphs.length; i++) {
    parts.push(open + hmcEscapeHtml(paragraphs[i]) + "</p>");
    parts.push(open + "&nbsp;</p>");
  }
  return parts.join("");
}

/* Reply and reply-all only. If Outlook cannot say what was pressed, go by the subject. */
function hmcIsReply(composeType, subject) {
  if (composeType === "reply" || composeType === "replyAll") {
    return true;
  }
  if (composeType === "newMail" || composeType === "forward") {
    return false;
  }
  return /^\s*re:/i.test(String(subject || ""));
}

/* Guard: email text may only be sent to a Make webhook address. */
function hmcIsAllowedUrl(url) {
  return /^https:\/\/hook\.[a-z0-9-]+\.make\.com\/[A-Za-z0-9]+$/.test(String(url || ""));
}

function hmcEncodeForm(fields) {
  var pairs = [];
  for (var name in fields) {
    if (Object.prototype.hasOwnProperty.call(fields, name)) {
      var value = fields[name] === null || fields[name] === undefined ? "" : String(fields[name]);
      pairs.push(encodeURIComponent(name) + "=" + encodeURIComponent(value));
    }
  }
  return pairs.join("&");
}

/* Validates the reply writer's answer. Throws if it is not in the agreed shape. */
function hmcReadResponse(text) {
  var data = JSON.parse(text);
  var decision = data && data.decision;
  var reply = decision && decision.reply;
  if (!reply || typeof reply.create !== "boolean") {
    throw new Error("answer not in the agreed shape");
  }
  var paragraphs = [];
  if (reply.create) {
    if (!Array.isArray(reply.paragraphs)) {
      throw new Error("paragraphs missing");
    }
    var total = 0;
    for (var i = 0; i < reply.paragraphs.length; i++) {
      var paragraph = reply.paragraphs[i];
      if (typeof paragraph === "string" && paragraph.trim() !== "") {
        paragraphs.push(paragraph.trim());
        total += paragraph.trim().length;
      }
    }
    if (paragraphs.length === 0 || paragraphs.length > HMC.MAX_PARAGRAPHS || total > HMC.MAX_REPLY_CHARS) {
      throw new Error("reply empty or too long");
    }
  }
  var bcc = [];
  var list = data.preset && data.preset.message && data.preset.message.bccRecipients;
  if (Array.isArray(list)) {
    for (var j = 0; j < list.length; j++) {
      var address = list[j] && list[j].emailAddress && list[j].emailAddress.address;
      if (typeof address === "string") {
        bcc.push(address.trim().toLowerCase());
      }
    }
  }
  return {
    create: reply.create,
    paragraphs: paragraphs,
    reason: decision && typeof decision.reason === "string" ? decision.reason : "",
    bcc: bcc
  };
}

/* Guard: Bcc is only ever added for addresses in the mailbox owner's own domain. */
function hmcSameDomain(addresses, userAddress) {
  var at = String(userAddress || "").lastIndexOf("@");
  if (at < 1) {
    return [];
  }
  var domain = String(userAddress).slice(at).toLowerCase();
  var kept = [];
  for (var i = 0; i < addresses.length; i++) {
    var address = String(addresses[i] || "").toLowerCase();
    var isAddress = /^[^@\s]+@[^@\s]+$/.test(address);
    if (isAddress && address.slice(address.lastIndexOf("@")) === domain && kept.indexOf(address) === -1) {
      kept.push(address);
    }
  }
  return kept;
}

/* Addresses on the email being answered: the From, To and Cc lines of the first quoted header in the reply window. */
function hmcChainAddresses(bodyText) {
  var text = String(bodyText || "");
  var start = text.search(/^[ \t>]*From:/im);
  if (start === -1) {
    return [];
  }
  var rest = text.slice(start);
  var end = rest.search(/^[ \t>]*Subject:/im);
  var block = end === -1 ? rest.slice(0, 2000) : rest.slice(0, end);
  var found = block.match(/[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g) || [];
  var addresses = [];
  for (var i = 0; i < found.length; i++) {
    var address = found[i].toLowerCase();
    if (addresses.indexOf(address) === -1) {
      addresses.push(address);
    }
  }
  return addresses;
}

/* Leaves out of Bcc anyone already on the email: To or Cc of the reply, or From, To or Cc of the email being answered. */
function hmcNotOnEmail(addresses, to, cc, bodyText) {
  var onEmail = hmcChainAddresses(bodyText);
  var listed = (String(to || "") + "," + String(cc || "")).toLowerCase().split(",");
  for (var i = 0; i < listed.length; i++) {
    var entry = listed[i].trim();
    if (entry !== "") {
      onEmail.push(entry);
    }
  }
  var kept = [];
  for (var j = 0; j < addresses.length; j++) {
    if (onEmail.indexOf(String(addresses[j]).toLowerCase()) === -1) {
      kept.push(addresses[j]);
    }
  }
  return kept;
}

/* ---------- Outlook calls ---------- */

function hmcNotify(item, type, message, done) {
  var details = { type: type, message: String(message).slice(0, 150) };
  if (type === Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage) {
    details.icon = "Icon.16x16";
    details.persistent = false;
  }
  item.notificationMessages.replaceAsync(HMC.NOTICE, details, function () {
    if (done) {
      done();
    }
  });
}

function hmcGetComposeType(item, done) {
  if (typeof item.getComposeTypeAsync !== "function") {
    done("");
    return;
  }
  item.getComposeTypeAsync(function (result) {
    var ok = result.status === Office.AsyncResultStatus.Succeeded && result.value;
    done(ok ? String(result.value.composeType || "") : "");
  });
}

function hmcGetRecipients(field, done) {
  field.getAsync(function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
      done("");
      return;
    }
    var addresses = [];
    for (var i = 0; i < result.value.length; i++) {
      addresses.push(result.value[i].emailAddress);
    }
    done(addresses.join(", "));
  });
}

function hmcAddBcc(item, addresses, done) {
  if (addresses.length === 0) {
    done(false);
    return;
  }
  var recipients = [];
  for (var i = 0; i < addresses.length; i++) {
    recipients.push({ displayName: addresses[i], emailAddress: addresses[i] });
  }
  item.bcc.addAsync(recipients, function (result) {
    done(result.status === Office.AsyncResultStatus.Succeeded);
  });
}

function hmcPost(url, form) {
  var request = fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: form
  }).then(function (response) {
    return response.text().then(function (text) {
      return { status: response.status, text: text };
    });
  });
  var timer = new Promise(function (resolve, reject) {
    setTimeout(function () {
      reject(new Error("timeout"));
    }, HMC.TIMEOUT_MS);
  });
  return Promise.race([request, timer]);
}

function onNewMessageComposeHandler(event) {
  var item = Office.context.mailbox.item;
  var types = Office.MailboxEnums.ItemNotificationMessageType;
  var finished = false;

  function finish() {
    if (!finished) {
      finished = true;
      event.completed();
    }
  }

  /* Every failure ends with a notice on the email. Nothing fails silently. */
  function fail(message) {
    hmcNotify(item, types.ErrorMessage, message, finish);
  }

  try {
    hmcGetComposeType(item, function (composeType) {
      item.subject.getAsync(function (subjectResult) {
        var subject = subjectResult.status === Office.AsyncResultStatus.Succeeded ? subjectResult.value : "";
        if (!hmcIsReply(composeType, subject)) {
          finish();
          return;
        }
        var settings = Office.context.roamingSettings;
        var url = settings.get(HMC.SETTING_URL);
        var key = settings.get(HMC.SETTING_KEY);
        if (!hmcIsAllowedUrl(url) || !key) {
          fail("Inbox Desk Reply is not set up. Open Reply settings and enter the address and key.");
          return;
        }
        hmcNotify(item, types.ProgressIndicator, "Inbox Desk is writing your reply...", function () {
          hmcGetRecipients(item.to, function (to) {
            hmcGetRecipients(item.cc, function (cc) {
              item.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
                if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
                  fail("Inbox Desk could not read this email. Nothing was added.");
                  return;
                }
                var form = hmcEncodeForm({
                  key: key,
                  composeType: composeType || "reply",
                  subject: subject,
                  to: to,
                  cc: cc,
                  body: String(bodyResult.value || "").slice(0, HMC.MAX_BODY_CHARS)
                });
                hmcPost(url, form)
                  .then(function (result) {
                    if (result.status === 401) {
                      fail("The reply writer did not accept the key. Check Reply settings.");
                      return;
                    }
                    if (result.status !== 200) {
                      fail("Inbox Desk could not write this reply. Nothing was added.");
                      return;
                    }
                    var answer = hmcReadResponse(result.text);
                    if (!answer.create) {
                      hmcNotify(item, types.InformationalMessage, "No reply written. " + answer.reason, finish);
                      return;
                    }
                    item.body.prependAsync(
                      hmcParagraphsToHtml(answer.paragraphs),
                      { coercionType: Office.CoercionType.Html },
                      function (insertResult) {
                        if (insertResult.status !== Office.AsyncResultStatus.Succeeded) {
                          fail("Inbox Desk wrote a reply but could not place it in this email.");
                          return;
                        }
                        var ownDomain = hmcSameDomain(answer.bcc, Office.context.mailbox.userProfile.emailAddress);
                        var bcc = hmcNotOnEmail(ownDomain, to, cc, bodyResult.value);
                        hmcAddBcc(item, bcc, function (added) {
                          var note = "Draft reply written. Bcc could not be added.";
                          if (bcc.length === 0) {
                            note = "Draft reply written. No Bcc needed.";
                          } else if (added) {
                            note = "Draft reply written and Bcc added.";
                          }
                          hmcNotify(item, types.InformationalMessage, note + " Check it and complete any [TJ: ...] markers before sending.", finish);
                        });
                      }
                    );
                  })
                  .catch(function () {
                    fail("Inbox Desk could not write this reply. Nothing was added.");
                  });
              });
            });
          });
        });
      });
    });
  } catch (error) {
    finish();
  }
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}

/* For the self-check under Node only. Outlook has no "module". */
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    HMC: HMC,
    hmcEscapeHtml: hmcEscapeHtml,
    hmcParagraphsToHtml: hmcParagraphsToHtml,
    hmcIsReply: hmcIsReply,
    hmcIsAllowedUrl: hmcIsAllowedUrl,
    hmcEncodeForm: hmcEncodeForm,
    hmcReadResponse: hmcReadResponse,
    hmcSameDomain: hmcSameDomain,
    hmcChainAddresses: hmcChainAddresses,
    hmcNotOnEmail: hmcNotOnEmail
  };
}
