/*
 * Inbox Desk Reply - settings pane.
 * Saves the reply writer address and key to this mailbox's roaming settings
 * for this add-in. Neither value is held in any file.
 */
Office.onReady(function () {
  var URL_SETTING = "hmcReplyWriterUrl";
  var KEY_SETTING = "hmcReplyWriterKey";
  var settings = Office.context.roamingSettings;
  var urlBox = document.getElementById("url");
  var keyBox = document.getElementById("key");
  var status = document.getElementById("status");

  urlBox.value = settings.get(URL_SETTING) || "";
  status.textContent = settings.get(KEY_SETTING)
    ? "A key is saved. Leave the key box empty to keep it."
    : "No key is saved yet.";

  document.getElementById("save").addEventListener("click", function () {
    var url = urlBox.value.trim();
    var key = keyBox.value;
    if (!/^https:\/\/hook\.[a-z0-9-]+\.make\.com\/[A-Za-z0-9]+$/.test(url)) {
      status.textContent = "Enter the reply writer address from Make. It starts with https://hook.";
      return;
    }
    if (!key && !settings.get(KEY_SETTING)) {
      status.textContent = "Enter the key.";
      return;
    }
    if (key && key.length < 12) {
      status.textContent = "The key must be at least 12 characters.";
      return;
    }
    settings.set(URL_SETTING, url);
    if (key) {
      settings.set(KEY_SETTING, key);
    }
    settings.saveAsync(function (result) {
      keyBox.value = "";
      status.textContent =
        result.status === Office.AsyncResultStatus.Succeeded
          ? "Saved. Close this pane and press Reply on an email to test."
          : "Not saved: " + result.error.message;
    });
  });
});
