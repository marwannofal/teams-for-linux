# Teams + Outlook Workspace Flatpak

This is Marwan Nofal's independent fork of Teams for Linux by Ismael Martinez
and contributors. The GPL-3.0-or-later license and upstream attribution remain
intact. It is not an official Microsoft application or the upstream Flathub
release.

## Install and run

Download `io.github.marwannofal.teams_outlook-x86_64.flatpak` and its `.sha256`
file from the [personal releases page](https://github.com/marwannofal/teams-for-linux/releases).
Run these commands from the download directory:

```bash
sha256sum -c io.github.marwannofal.teams_outlook-x86_64.flatpak.sha256
flatpak remote-add --if-not-exists --user flathub https://dl.flathub.org/repo/flathub.flatpakrepo
flatpak install --user ./io.github.marwannofal.teams_outlook-x86_64.flatpak
flatpak run io.github.marwannofal.teams_outlook
```

Choose **Teams + Outlook Workspace** from the application menu for later
launches. The header contains Teams and Outlook tabs. Outlook loads on its
first selection and keeps its page open while you use Teams. Sign in through
Microsoft's normal pages. Outlook shares the primary Teams session; accounts
added through the Teams profile switcher do not get separate Outlook tabs.

This x86_64 build has a separate application ID and configuration directory
under `~/.var/app/io.github.marwannofal.teams_outlook/`. It can coexist with
`com.github.IsmaelMartinez.teams_for_linux`. Existing official-app cookies and
settings are not imported, so a fresh sign-in may be needed.

The launcher uses `zypak-wrapper` for Electron inside Flatpak. It enables
Outlook automatically. No host `chrome-sandbox` permission change is needed.
Permissions follow the upstream Flatpak manifest, including home-directory,
audio, device, network, desktop-notification and screen-sharing access.

## Update or remove

Download the next bundle from the personal releases page and install it with
the same `flatpak install --user ./...flatpak` command. A single-file bundle
does not provide automatic application updates. Flathub runtime dependencies
can still be updated with `flatpak update --user`.

```bash
# Remove this fork while retaining its data:
flatpak uninstall --user io.github.marwannofal.teams_outlook

# Only if you also want to erase its accounts and configuration:
flatpak uninstall --user --delete-data io.github.marwannofal.teams_outlook
```

## Build from source on Ubuntu

Use Node.js 24 LTS and an x86_64 Linux machine. Flatpak and Git must be
installed. On Ubuntu, install host tools with:

```bash
sudo apt-get update
sudo apt-get install -y flatpak git dbus-bin
flatpak remote-add --if-not-exists --user flathub https://dl.flathub.org/repo/flathub.flatpakrepo
flatpak install --user flathub org.flatpak.Builder org.freedesktop.Sdk//25.08 org.electronjs.Electron2.BaseApp//25.08

git clone --branch outlook-flatpak https://github.com/marwannofal/teams-for-linux.git
cd teams-for-linux
npm ci
npm run lint
npm run test:unit
npm run dist:outlook:flatpak
```

The script packages the checked-out source with electron-builder, builds the
pinned libcbor/libfido2 modules, exports a local Flatpak repository, and writes:

- `dist/io.github.marwannofal.teams_outlook-x86_64.flatpak`
- `dist/io.github.marwannofal.teams_outlook-x86_64.flatpak.sha256`
- `dist/flatpak-repo/` (local export repository)

The manifest is `io.github.marwannofal.teams_outlook.json` in this directory.
Runtime/base versions follow the upstream Flatpak's 25.08 baseline. The build
requires internet access for npm, Electron, Flathub and the pinned FIDO2 sources.
On a machine without a session bus, use:

```bash
dbus-run-session -- npm run dist:outlook:flatpak
```

For source development, use `npm run start:outlook`. If Ubuntu refuses the
source checkout's SUID sandbox, `npm run start:outlook:dev` is the existing
development workaround. That command disables Chromium's sandbox and should
only be used for local development. Prefer the Flatpak for daily use. If your
shell exports `ELECTRON_RUN_AS_NODE`, unset it before launching Electron.

## Verify changes

```bash
npm run lint
npm run test:unit
npx playwright install chromium
npm run test:e2e -- tests/e2e/outlook.spec.js
```

The Outlook E2E suite checks lazy loading, session sharing, renderer isolation,
draft preservation, header layout and resizing, retry after failed navigation,
authentication popups, IPC sender validation, keyboard navigation, and profile
coexistence. See [the feature test report](../../app/outlook/TEST-RESULTS.md)
for the current suite results and existing baseline failure.

After building, install the bundle and check that both tabs reach Microsoft
sign-in, switch without losing their page, survive resizing, and reopen after
quitting. Inbox, MFA, meetings, microphone/camera, screen sharing, notifications
and hardware security keys additionally need testing with your own account
and devices; reaching sign-in does not verify those features.

## Publish the next personal release

Work on the `outlook-flatpak` branch of your fork and keep `origin` pointing at
the repository you intend to push to. If your original checkout still uses
the upstream remote, push explicitly to your personal remote instead.

After checks pass, commit the source and build from that commit. Create a
unique tag such as `outlook-v2.32.1-2`, then publish the bundle and checksum
with the corresponding source tag:

```bash
git tag outlook-v2.32.1-2
git push marwan outlook-flatpak outlook-v2.32.1-2
gh release create outlook-v2.32.1-2 \
  --repo marwannofal/teams-for-linux \
  --title 'Teams + Outlook Workspace 2.32.1 — build 2' \
  --notes-file release-notes.md \
  dist/io.github.marwannofal.teams_outlook-x86_64.flatpak \
  dist/io.github.marwannofal.teams_outlook-x86_64.flatpak.sha256
```

Prepare `release-notes.md` with the changes, verification and known limitations
before running the command. GitHub supplies source archives for the release
tag, so recipients can obtain the corresponding GPL application source.

The **Personal Outlook Flatpak** GitHub Actions workflow can also build
downloadable artifacts when enabled on your fork, either manually or by
pushing an `outlook-v*` tag. It runs lint/unit checks and builds the bundle;
it does not publish a release or run desktop E2E tests. Other inherited upstream
release workflows are not needed for this personal distribution.
