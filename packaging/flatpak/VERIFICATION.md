# Personal Flatpak verification

Verified on this Ubuntu x86_64 desktop on 2026-11-18.

| Check | Result |
| --- | --- |
| Application ID | `io.github.marwannofal.teams_outlook` |
| Version | `2.32.1-outlook.1` |
| Runtime / Electron BaseApp | Freedesktop 25.08 / Electron2 25.08 |
| Build from local source | Passed |
| AppStream metadata composition | Passed |
| Bundle installation | Passed, user installation |
| Coexistence | Official Teams 2.32.0 remains installed system-wide under its original ID |
| Launch through zypak | Passed without the host SUID sandbox crash |
| Teams live page | Reached `teams.cloud.microsoft` |
| Outlook live page | Reached `login.microsoftonline.com`, title **Sign in to your account** |
| View security | Outlook has sandbox and context isolation enabled; Node integration disabled |
| Session | Outlook and primary Teams use the same Electron session |
| Switching | Teams/Outlook selection succeeds; Outlook's view remains alive |
| Resize | At 900 × 650 content size, header is 900 × 52; both app views start at y=52 with height 598 |
| Visual inspection | Header and Microsoft sign-in screenshots checked |
| Runtime libraries | `ldd` reports no missing libraries |
| Optional security-key tooling | `/app/bin/fido2-token` present |
| Desktop integration assets | Own desktop entry, icon, metadata and GPL license installed |
| Lint | Passed |
| Full unit suite | 1,353 passed, 1 skipped, 0 failed |
| Focused Outlook E2E suite | All 8 passed |

The final bundle is approximately 104 MiB. SHA-256:

```text
49c30b04cac23fe6fa1f4fcba1d07266851f396c98a97e4b05b86686dcf3c98b
```

The installed-package probe used an isolated temporary profile and a temporary
main-process inspector to inspect native Electron views, then quit the app
and removed the profile. No credentials were entered. The normal desktop
launcher has no debugging flags. Screenshots are included with the release.

Signed-in inbox operations, MFA, authenticated-session persistence, actual
calls, camera/microphone, screen sharing, notifications and physical security
keys still require account/device acceptance testing. The full source E2E run
previously produced 38 passes and one independently reproduced baseline
profile-rename failure; see [the feature test report](../../app/outlook/TEST-RESULTS.md).
This is local-machine verification, not a cross-distribution certification or
a completed GitHub Actions run.
