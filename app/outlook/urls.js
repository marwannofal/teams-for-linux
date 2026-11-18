const DEFAULT_URL = "https://outlook.office.com/mail/";
const OUTLOOK_HOSTS = new Set([
  "outlook.office.com", "outlook.office365.com", "outlook.cloud.microsoft",
  "outlook.live.com",
]);
const AUTH_HOSTS = new Set([
  "login.microsoftonline.com", "login.microsoft.com", "login.live.com",
]);

function isOutlookUrl(value) {
  try {
    const url = new URL(value);
    return url.protocol === "https:" && !url.username && !url.password &&
      !url.port && OUTLOOK_HOSTS.has(url.hostname);
  } catch {
    return false;
  }
}

module.exports = { DEFAULT_URL, OUTLOOK_HOSTS, AUTH_HOSTS, isOutlookUrl };
