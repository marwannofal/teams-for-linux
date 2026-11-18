# Outlook tab

An optional Outlook web view runs alongside Teams in the same application window.
The Teams / Outlook navbar spans the top of the application window. Both apps
have their own viewport below it, so their search boxes and other controls stay
visible. Switching tabs keeps both pages alive, preserving drafts and the Teams
renderer. The multi-account profile pill remains at the bottom-left.

## Run from this checkout

Use Node.js 24, as the repository's CI does, and npm. A graphical Linux session
and Electron's system libraries are required.

```bash
cd /home/marwan/Documents/teams-for-linux
npm ci
npm run start:outlook
```

Click **Outlook** in the header navbar. The first selection loads
`https://outlook.office.com/mail/`. Complete Microsoft's work/school sign-in and
MFA if requested. Click **Teams** to return. The arrow button reloads Outlook;
after a network failure its accessible label changes to **Retry Outlook**.

`npm start` also works once Outlook is enabled in `config.json`; its `prestart`
hook reinstalls dependencies with `npm ci`. `start:outlook` launches directly,
so install dependencies first.

### Ubuntu sandbox-helper startup error

If startup aborts with **The SUID sandbox helper binary was found, but is not
configured correctly**, use the explicit development launcher:

```bash
npm run start:outlook:dev
```

For an isolated development session:

```bash
npm run start:outlook:dev -- --user-data-dir=/tmp/teams-outlook-dev
```

This uses `--no-sandbox`, matching the repository's existing `start:dev` command.
It disables Chromium's process sandbox for that launch; use it as a local
development workaround. The normal `start:outlook` command retains the default
sandbox behavior. Do not run the application with `sudo` or change ownership of
the project to work around this error. Reinstalling dependencies with `npm ci`
does not configure a privileged sandbox helper.

## Persistent configuration

Merge this into your existing `config.json`, preserving other settings:

```json
{
  "outlook": {
    "enabled": true,
    "url": "https://outlook.office.com/mail/"
  }
}
```

For personal Outlook/Hotmail accounts use `https://outlook.live.com/mail/0/`.
The URL must use HTTPS and one of `outlook.office.com`, `outlook.office365.com`,
`outlook.cloud.microsoft` or `outlook.live.com`. Other configured hosts disable
the feature with a warning; Teams still starts. Changes require a restart.

Configuration locations:

| Installation | File |
| --- | --- |
| Source, deb, rpm, AppImage | `~/.config/teams-for-linux/config.json` |
| Snap | `~/snap/teams-for-linux/current/.config/teams-for-linux/config.json` |
| Flatpak | `~/.var/app/com.github.IsmaelMartinez.teams_for_linux/config/teams-for-linux/config.json` |

For a separate development session without changing your normal login:

```bash
npm run start:outlook -- --user-data-dir=/tmp/teams-outlook-dev
```

Set `outlook.enabled` to `false` and restart to remove the switcher.

## Sessions and behavior

- Outlook shares the primary Teams session partition, including Microsoft
  sign-in cookies. SSO can reuse those cookies, but tenant policy or account
  selection can still require a separate sign-in.
- With multiple Teams profiles, Outlook stays on the primary partition. It
  does not switch mail accounts when you switch Teams profiles.
- Switching tabs never navigates or reloads Teams. Actual call continuity
  still requires verification with a signed-in meeting and working devices.
- Outlook uses normal browser notifications and the existing session download
  manager. Teams' unread badge and presence integrations remain Teams-specific.
- The application menus and global call shortcuts continue to target Teams.
  Use Outlook's own interface for mail actions and browser history.
- Regular external links open in your default browser. Microsoft sign-in
  popups stay in sandboxed child windows; they share the session.

## Implementation

- `index.js`: Outlook view lifecycle, tab selection, popup policy, error/retry,
  resize, focus and shutdown. Outlook loads lazily on first selection.
- `window.js`: native `BaseWindow` containing a persistent Teams view, with the
  primary renderer interface used by existing desktop services.
- `layout.js`: shared header height and viewport bounds for Teams, Outlook and
  additional Teams profiles.
- `urls.js`: exact official-host URL policy, shared with permission handling.
- `tabs.html`, `tabs.css`, `tabs.js`: local tab-switcher UI with keyboard support.
- `preload.js`: narrow bridge on the local switcher only.

The Outlook page has `contextIsolation: true`, `sandbox: true`,
`nodeIntegration: false` and no preload. No Teams DOM tools, Graph tokens or
Electron bridge are injected into mail pages. The tab action IPC channel is
allowlisted, restricted to local app pages and additionally pinned to the exact
switcher webContents and main frame. Mail permissions are limited to official
Outlook pages and notifications, clipboard and fullscreen.

## Verification

See [the recorded verification results](TEST-RESULTS.md) for this checkout's
test counts, baseline failure comparison and live sign-in-page checks.

```bash
npm run lint
npm run test:unit
npx playwright install chromium
npm run test:e2e -- tests/e2e/outlook.spec.js
npm run test:e2e
npm run pack
```

The Outlook E2E tests launch the real application against controlled Teams and
Outlook fixtures. They verify lazy loading, session sharing, renderer isolation,
real mouse tab selection, drafts and renderer state across switches, resizing,
load failure and retry, and Microsoft authentication popup isolation. Fixtures
do not prove that a real mailbox or tenant authentication works.

For a real-account acceptance test:

1. Start with Outlook enabled and sign into Teams.
2. Select Outlook, finish sign-in/MFA and confirm your inbox appears.
3. Open a message; compose a draft and switch between tabs. Confirm the draft
   remains and both pages accept typing.
4. Send a test message to yourself and confirm it arrives. Download an attachment.
5. Join a Teams test call, switch to Outlook and confirm audio/video continue.
6. Disconnect networking, reload Outlook, reconnect and use Retry.
7. Restart the app and check the Microsoft session is retained according to
   your organization's sign-in policy.

On agents that set `ELECTRON_RUN_AS_NODE=1`, unset that variable when launching
Electron or E2E tests. The CI launch path uses `CI=true` to pass `--no-sandbox`
for Chromium's OS process sandbox in restricted test environments; the normal
desktop launch should use its default sandbox settings.

Build an unpacked application and launch it with Outlook enabled:

```bash
npm run pack
./dist/linux-unpacked/teams-for-linux --outlook.enabled
```

Create packages with `npm run dist:linux` or an individual target such as
`npm run dist:linux:deb`. Platform packaging tools are required for those targets.
