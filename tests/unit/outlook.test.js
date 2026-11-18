"use strict";
const { test } = require("node:test");
const assert = require("node:assert/strict");
const { isOutlookUrl } = require("../../app/outlook/urls");
const { isSenderAllowed, allowedChannels } = require("../../app/security/ipcValidator");

test("Outlook accepts official work and personal HTTPS endpoints", () => {
  for (const host of ["outlook.office.com", "outlook.office365.com", "outlook.cloud.microsoft", "outlook.live.com"]) {
    assert.equal(isOutlookUrl(`https://${host}/mail/0/`), true);
  }
});

test("Outlook rejects lookalike hosts, unsafe schemes and embedded credentials", () => {
  for (const url of [
    "http://outlook.office.com/", "https://outlook.office.com.evil.test/",
    "https://evil.outlook.office.com/", "file:///tmp/mail.html",
    "javascript:alert(1)", "https://name:secret@outlook.office.com/",
    "https://outlook.office.com:8443/", "not a URL", null,
  ]) assert.equal(isOutlookUrl(url), false, String(url));
});

test("workspace actions are allowlisted and reject remote senders", () => {
  assert.equal(allowedChannels.has("workspace-tab-action"), true);
  assert.equal(isSenderAllowed("workspace-tab-action", {
    senderFrame: { url: "https://outlook.office.com/mail/" },
  }), false);
});
