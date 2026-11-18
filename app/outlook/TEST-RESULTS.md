# Outlook verification results

Verified on 2026-11-18 in the local Linux x86_64 desktop environment.

## Environment

- Node.js 24.22.3, matching the major version used by repository CI.
- Electron 45.3.2 from the existing lockfile.
- Playwright 1.65.0 and its pinned Chromium browser installed.
- Project dependencies installed with `npm ci`; dependency versions unchanged.
- Tests used isolated temporary user-data directories, without personal accounts.
- Electron launches unset the agent environment's `ELECTRON_RUN_AS_NODE=1`.
  E2E runs used `CI=true` to use the repository's `--no-sandbox` test launch.

## Automated results

| Check | Result |
| --- | --- |
| `npm run lint` | Passed |
| `npm run test:unit` | 1,353 passed, 1 skipped, 0 failed |
| Full E2E suite, retries disabled | 38 passed, 1 failed |
| Outlook E2E checks within that suite | All 8 passed |
| `npm run pack` | Passed; `dist/linux-unpacked/teams-for-linux` created |
| Packaged asset inspection | Outlook module, preload, HTML, CSS and scripts included in `app.asar`; final recovery and keyboard code present |
| Dynamic library check | No missing libraries reported by `ldd` for the built executable |
| Configuration and IPC generators | Regenerated successfully |
| `git diff --check` | Passed |
| `npm run start:outlook:dev` | Started successfully with an isolated profile; Teams reached Microsoft sign-in and the Outlook switcher loaded |

Outlook runtime checks covered lazy loading, session sharing, renderer isolation,
mouse selection, keyboard navigation, draft preservation, Teams renderer
continuity, resizing, load failure/retry, authentication popups, IPC sender
validation and coexistence with the multi-account profile switcher.

### Header navbar follow-up

The footer switcher was replaced with a full-width 52-pixel header navbar.
With Outlook enabled, a native BaseWindow owns separate Teams, Outlook and
navbar views. Both app viewports and additional Teams profiles start below the
header; no Teams page CSS is patched to make room.

All eight focused E2E checks passed after this change, including native Teams
viewport dimensions after resizing and the selected profile's header offset.
Lint, the unit suite (1,353 passed, 1 skipped) and the Linux build also passed.
The source app reached Outlook's real Microsoft sign-in page again, and the
header was captured and visually inspected at
`/tmp/teams-outlook-header-tabs.png`. The full regression rerun for the header
finished with 38 passed and the same baseline profile-rename failure. Logs are
in `/tmp/teams-header-full-e2e.log`, `/tmp/teams-header-e2e.log`,
`/tmp/teams-unit-header.log`, `/tmp/teams-header-lint.log` and
`/tmp/teams-header-pack.log`.
The rebuilt Linux application also reached Outlook's real sign-in page with
the header navbar; its navbar screenshot was visually checked at
`/tmp/teams-outlook-header-packaged-tabs.png`. The launch result is recorded in
`/tmp/teams-header-packaged.log`.

The existing failure is
`tests/e2e/multi-account-enabled.spec.js:434`, **toggling a notification switch
mid-rename keeps the typed name**. It fails with a renderer script execution
error while Outlook is disabled. The same test was run against an untouched
`git archive HEAD` copy in `/tmp/teams-outlook-baseline`, using the same installed
dependencies, and failed with the same error. This confirms the failure also
exists in the baseline under this environment. No profile-rename code was
changed for the Outlook feature.

## Live service checks

The source application reached `login.microsoftonline.com` from the Outlook
tab and displayed **Sign in to your account**, without entering credentials.
The Teams page remained in its own renderer. The built Linux application also
reached the same Microsoft sign-in page successfully. Its mail view and tab
switcher were captured and visually inspected. Screenshots are in
`/tmp/teams-outlook-packaged.png` and `/tmp/teams-outlook-packaged-tabs.png`;
the packaged smoke-test output is in `/tmp/teams-packaged-outlook.log`.

## Scope still requiring a signed-in user

The automated tests use controlled fixtures. Live sign-in-page checks establish
connectivity and page rendering, not access to a mailbox. The following remain
manual acceptance checks:

- Account sign-in, tenant policies and MFA completion.
- Inbox contents, sending/receiving mail and attachment downloads.
- Persistence of a real authenticated session after restart.
- Audio/video continuity during an actual Teams call and device-dependent sharing.
- Other operating systems, distributions and compositor configurations.

See [the run guide](README.md) for the account acceptance checklist and complete
configuration, development, test and build instructions.

## Reproduce

```bash
npm ci
npx playwright install chromium
npm run lint
npm run test:unit
npm run test:e2e -- tests/e2e/outlook.spec.js
npm run test:e2e -- --retries=0
npm run pack
```

On this agent environment, prefix Electron test commands with
`env -u ELECTRON_RUN_AS_NODE CI=true`. A normal desktop terminal does not need
the agent-specific variable adjustment.

Detailed execution logs from this verification are in
`/tmp/teams-unit-tests.log`, `/tmp/teams-full-e2e.log`,
`/tmp/teams-baseline-e2e.log`, `/tmp/teams-lint.log`, `/tmp/teams-pack.log`
and `/tmp/teams-live-outlook.log`.

The normal source launcher failed on this Ubuntu host because the downloaded
`chrome-sandbox` helper is user-owned with mode 755 and unprivileged user
namespaces are restricted by AppArmor. The explicit development launcher uses
`--no-sandbox`, matching `start:dev`; it was smoke-tested directly. This bypasses
Chromium's process sandbox for development and does not repair the host's sandbox
configuration. Its temporary test instance was limited to a 15-second run.

Dependency installation also reported 13 npm audit findings (7 moderate and
6 high) in the existing dependency tree. Dependency upgrades were outside this
feature change; the lockfile was not modified.
