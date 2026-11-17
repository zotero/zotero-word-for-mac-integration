#!/bin/bash
#
# Builds install/libZoteroWordIntegration.dylib (universal, arm64 and x86_64) with Xcode's
# command-line tools, and checks that it loads on the oldest macOS that Zotero supports.
#
# The dylib must not require a newer macOS than Zotero, or use chained fixups (macOS 12 and
# later): Zotero loads it at startup, and crashes on older macOS if it can't. Recent Xcode
# versions only accept deployment targets from macOS 12, so the project sets
# MACOSX_DEPLOYMENT_TARGET to 12.0 and LLVM_TARGET_TRIPLE_OS_VERSION, the target OS that Xcode
# passes to the compiler and linker, to the macOS that the dylib actually needs.

set -e

# The oldest macOS that Zotero supports. arm64 Macs start at macOS 11, which clang uses instead.
MIN_MACOS=10.15

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
DYLIB="$PROJECT_ROOT/install/libZoteroWordIntegration.dylib"
BUILD_DIR="$(mktemp -d)"
trap 'rm -rf "$BUILD_DIR"' EXIT

fail() {
	echo "Error: $1"
	exit 1
}

echo "Building libZoteroWordIntegration.dylib"
cd "$PROJECT_ROOT/build/src"
touch "$BUILD_DIR/start"
# The project copies the dylib into install/, without the ad-hoc signature of the build. Zotero
# signs it when it packages the app.
if ! xcodebuild -project libZoteroMacWordIntegration.xcodeproj -target ZoteroWordIntegration \
		-configuration Release ARCHS="arm64 x86_64" ONLY_ACTIVE_ARCH=NO SYMROOT="$BUILD_DIR/xcode" \
		clean build > "$BUILD_DIR/build.log" 2>&1; then
	grep -E "(^| )error:" "$BUILD_DIR/build.log" || tail -20 "$BUILD_DIR/build.log"
	fail "The build failed"
fi
# APIs newer than the deployment target would fail at runtime on older macOS
grep -E "warning: .*(availab|introduced in macOS)" "$BUILD_DIR/build.log" | sort -u || true

[ "$DYLIB" -nt "$BUILD_DIR/start" ] || fail "The build didn't copy the dylib into install/"

archs=$(lipo -archs "$DYLIB")
[[ " $archs " == *" arm64 "* && " $archs " == *" x86_64 "* ]] \
	|| fail "The dylib has architectures \"$archs\" instead of arm64 and x86_64"

for arch in x86_64 arm64; do
	expected=$MIN_MACOS
	[ "$arch" = arm64 ] && expected=11.0
	minimum=$(otool -arch "$arch" -l "$DYLIB" | awk '
		/cmd LC_VERSION_MIN_MACOSX/ { kind = "version" }
		/cmd LC_BUILD_VERSION/ { kind = "minos" }
		kind && $1 == kind { print $2; exit }')
	[ "$minimum" = "$expected" ] \
		|| fail "The $arch dylib requires macOS $minimum instead of $expected. Check LLVM_TARGET_TRIPLE_OS_VERSION in the project."
	! otool -arch "$arch" -l "$DYLIB" | grep -q LC_DYLD_CHAINED_FIXUPS \
		|| fail "The $arch dylib uses chained fixups, which crash Zotero before macOS 12"
	echo "$arch: requires macOS $minimum, no chained fixups"
done

echo "Built $DYLIB"
