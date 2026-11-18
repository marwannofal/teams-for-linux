#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")/.."

if [ "$(uname -m)" != x86_64 ]; then
  echo 'This release script currently supports x86_64 Linux only.' >&2
  exit 1
fi

APP_ID=io.github.marwannofal.teams_outlook
MANIFEST="packaging/flatpak/$APP_ID.json"
BUNDLE="dist/$APP_ID-x86_64.flatpak"

# Build from the checked-out source; never download an upstream application.
npm run pack -- --linux --x64 --publish never
flatpak run org.flatpak.Builder --user --force-clean --disable-rofiles-fuse \
  --install-deps-from=flathub --arch=x86_64 --default-branch=stable \
  --repo=dist/flatpak-repo dist/flatpak-build "$MANIFEST"
flatpak build-bundle --arch=x86_64 \
  --runtime-repo=https://dl.flathub.org/repo/flathub.flatpakrepo \
  dist/flatpak-repo "$BUNDLE" "$APP_ID" stable
(cd dist && sha256sum "$(basename "$BUNDLE")" > "$(basename "$BUNDLE").sha256")
echo "Built $BUNDLE"
