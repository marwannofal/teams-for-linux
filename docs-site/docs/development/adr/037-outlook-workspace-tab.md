# ADR-037: Optional Outlook workspace tab

## Status

Implemented in the local Outlook-tab change; not released.

## Context

A user requested Outlook alongside Teams in the same application. Navigating
the existing Teams renderer to Outlook would discard Teams state and end calls.
The current main window is the root Teams webContents, with existing profile
views and a compact bottom-left profile switcher layered above it.

## Decision

Add an opt-in `outlook.enabled` configuration option and a dedicated sandboxed
`WebContentsView` for Outlook. Load it lazily on first selection. Switching tabs
changes visibility and z-order without navigating Teams. Keep both views
attached while switching, following the profile-view input handling documented
in ADR-020.

Use a full-width local Teams / Outlook navbar in the header, as requested after
testing the initial footer switcher. Reserve 52 pixels above both apps and
additional Teams profiles. With Outlook enabled, use Electron's native
`BaseWindow` and independent `WebContentsView` instances; retain a stable Teams
webContents and provide the renderer conveniences existing main-process services
use. The ordinary BrowserWindow path remains in place when Outlook is disabled.
This avoids covering Teams' search and call controls or relying on CSS patches
against Microsoft's page layout. The profile pill stays at the bottom-left.

Share the primary Teams session partition to allow Microsoft's existing SSO
cookies to participate in mail sign-in. Outlook remains associated with this
partition when the user switches Teams profiles. Account selection and tenant
policy remain Microsoft's responsibility.

Do not reuse the Teams preload in Outlook. Outlook and its authentication
popups use context isolation, the renderer sandbox, no Node integration and
no Electron bridge. Only the local switcher has a small preload, with fixed
actions pinned to its exact webContents and main frame. Limit configured entry
URLs to official Outlook HTTPS hosts. Extend session permissions only for
notifications, clipboard and fullscreen on official Outlook pages.

## Consequences

- Teams calls and page state remain in their original renderer; the runtime
  tests verify renderer and draft preservation. A signed-in call is still a
  separate acceptance check.
- Outlook uses an additional renderer and consumes resources after first use.
- Both app viewports start below the header, including multi-account profile
  views. Resizing updates native bounds without reloading either app.
- The workspace window owns and closes its primary Teams view explicitly;
  unlike BrowserWindow, BaseWindow doesn't destroy child webContents itself.
- Teams menus, badge counts, presence and shortcuts retain their Teams meaning.
- This implementation provides one Outlook workspace, not a mailbox per Teams
  profile. Sign-in reuse is possible, not guaranteed.
- Automated fixture tests exercise tab behavior and security. Real inbox,
  tenant policy, MFA, mail delivery and device-dependent calls require manual
  acceptance testing.
