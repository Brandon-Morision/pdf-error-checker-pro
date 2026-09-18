"""
test_updater.py — exercises the full auto-update path end to end without
needing a published GitHub release.

A local HTTP server stands in for api.github.com, serving a real GitHub
release JSON payload and a real downloadable asset. Everything from
UpdateChecker.check_for_updates() through download_asset() and
apply_source_update() runs for real against it — so this covers the exact
code that would run against a live release, minus GitHub's own hosting.

Run:  python3 test_updater.py
"""
import json
import os
import shutil
import sys
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from pdf_checker_core import (  # noqa: E402
    UpdateChecker, download_asset, apply_source_update,
)

PASS, FAIL = "PASS", "FAIL"
results = []


def check(label, condition, detail=""):
    results.append((PASS if condition else FAIL, label, detail))
    print(f"[{PASS if condition else FAIL}] {label}" + (f" — {detail}" if detail else ""))


# The content the "new version" asset will contain. apply_source_update
# relaunches whatever it installs, so this must be a harmless, instantly
# exiting script.
NEW_VERSION_SOURCE = (
    "# updated version 9.9.9\n"
    "import sys\n"
    "sys.exit(0)\n"
)

ASSET_NAME = "pdf_checker_qt.py"


class FakeGitHubHandler(BaseHTTPRequestHandler):
    """Serves a release payload and its asset, shaped exactly like the
    GitHub REST API responses the real code consumes."""

    def do_GET(self):
        # Checked BEFORE the generic /releases/latest branch — this path
        # also ends with "/releases/latest", so ordering matters here.
        if "/no-releases/" in self.path:
            body = b'{"message":"Not Found"}'
            self.send_response(404)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif self.path.endswith("/releases/latest"):
            payload = {
                "tag_name": "v9.9.9",
                "html_url": "http://localhost/fake/releases/tag/v9.9.9",
                "body": "Fake release notes for the updater test.",
                "assets": [{
                    "name": ASSET_NAME,
                    "size": len(NEW_VERSION_SOURCE.encode()),
                    "browser_download_url": f"http://{self.server.server_address[0]}:{self.server.server_address[1]}/asset",
                }],
            }
            body = json.dumps(payload).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif self.path == "/asset":
            body = NEW_VERSION_SOURCE.encode()
            self.send_response(200)
            self.send_header("Content-Type", "text/plain")
            # Content-Length present so download_asset reports real
            # fractional progress rather than falling back to indeterminate.
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, *args):
        pass  # keep test output clean


def main():
    server = HTTPServer(("127.0.0.1", 0), FakeGitHubHandler)
    host, port = server.server_address
    threading.Thread(target=server.serve_forever, daemon=True).start()
    base = f"http://{host}:{port}"
    print(f"Fake GitHub server listening on {base}\n")

    try:
        # --- 1. Version comparison logic -------------------------------
        checker = UpdateChecker("fake/repo", "0.2.0")
        check("newer version detected", checker.compare_versions("9.9.9", "0.2.0"))
        check("same version not flagged", not checker.compare_versions("0.2.0", "0.2.0"))
        check("older version not flagged", not checker.compare_versions("0.1.0", "0.2.0"))
        check("longer-but-equal prefix counts as newer",
              checker.compare_versions("0.2.0.1", "0.2.0"))
        check("malformed version string handled",
              not checker.compare_versions("not-a-version", "0.2.0"))

        # --- 2. Live check against the fake API ------------------------
        checker.api_url = f"{base}/repos/fake/repo/releases/latest"
        has_update, latest, data = checker.check_for_updates()
        check("check_for_updates reports an update", has_update)
        check("version parsed with 'v' stripped", latest == "9.9.9", f"got {latest!r}")
        check("release data returned", isinstance(data, dict) and "assets" in data)

        # --- 3. A repo with no releases (the real repo's current state) --
        checker_404 = UpdateChecker("fake/repo", "0.2.0")
        checker_404.api_url = f"{base}/repos/no-releases/releases/latest"
        has_update_404, latest_404, data_404 = checker_404.check_for_updates()
        check("404 (no releases) treated as 'no update', not an error",
              has_update_404 is False and latest_404 is None and data_404 is None)

        # --- 4. Asset selection ----------------------------------------
        # Fixed as part of this pass: source-mode installs no longer
        # auto-select ANY asset. The app is two files now, not one, and a
        # release asset is always exactly one file — auto-installing just
        # pdf_checker_qt.py while a release also changed pdf_checker_core.py
        # would silently leave the engine file out of sync. Source updates
        # always fall back to "Open Release Page" until a full-bundle
        # (e.g. zip) release format exists.
        asset = UpdateChecker.find_installable_asset(data)
        check("source-mode installs no longer auto-select an asset "
              "(disabled — see Bug 5 fix)", asset is None, f"got {asset}")

        frozen_release = {
            "assets": [
                {"name": ASSET_NAME, "browser_download_url": "http://x/py"},
                {"name": "PDF-Error-Checker-Pro-windows.exe", "browser_download_url": "http://x/exe"},
            ]
        }
        try:
            sys.frozen = True
            orig_platform = sys.platform
            sys.platform = "win32"
            frozen_asset = UpdateChecker.find_installable_asset(frozen_release)
            check("a frozen Windows build DOES select the .exe asset",
                  frozen_asset is not None and frozen_asset["name"].endswith(".exe"),
                  f"got {frozen_asset}")
        finally:
            sys.platform = orig_platform
            del sys.frozen
        check("no asset selected when release has none",
              UpdateChecker.find_installable_asset({"assets": []}) is None)

        # download_asset / apply_source_update are still tested directly
        # below with a manually-chosen asset (bypassing find_installable_asset,
        # which — correctly, per the fix above — won't pick one for source
        # mode). The functions themselves remain correct building blocks,
        # just not wired into the automatic flow for a multi-file source
        # install anymore.
        asset = data["assets"][0]

        # --- 5. Real download with progress ----------------------------
        tmpdir = Path(tempfile.mkdtemp())
        dest = tmpdir / ASSET_NAME
        seen_progress = []
        downloaded = download_asset(asset["browser_download_url"], dest,
                                    on_progress=seen_progress.append)
        check("asset downloaded to disk", dest.exists())
        check("downloaded content matches what the server sent",
              dest.read_text() == NEW_VERSION_SOURCE)
        check("progress callback fired", len(seen_progress) > 0,
              f"{len(seen_progress)} callback(s), last={seen_progress[-1] if seen_progress else None}")
        check("progress ends at 100%",
              bool(seen_progress) and abs(seen_progress[-1] - 1.0) < 1e-9)
        check("download_asset returns the destination path",
              Path(downloaded) == dest)

        # --- 6. Real source-update install -----------------------------
        # A stand-in for the running app script, in its own directory so
        # the real files are never touched.
        app_dir = Path(tempfile.mkdtemp())
        running_script = app_dir / "pdf_checker_qt.py"
        original_content = "# original version 0.2.0\nimport sys\nsys.exit(0)\n"
        running_script.write_text(original_content)

        apply_source_update(dest, current_script_path=running_script)

        backup = running_script.with_suffix(running_script.suffix + ".bak")
        check("running script replaced with the new version",
              running_script.read_text() == NEW_VERSION_SOURCE)
        check("a .bak backup of the previous version was kept", backup.exists())
        check("the .bak contains the ORIGINAL content",
              backup.exists() and backup.read_text() == original_content)

        # --- 7. Failure handling ---------------------------------------
        try:
            download_asset(f"{base}/does-not-exist", tmpdir / "nope.bin")
            check("bad download URL raises", False, "no exception raised")
        except Exception as e:
            check("bad download URL raises rather than failing silently",
                  True, type(e).__name__)

        checker_dead = UpdateChecker("fake/repo", "0.2.0")
        checker_dead.api_url = "http://127.0.0.1:1/unreachable"
        dead_result = checker_dead.check_for_updates()
        check("unreachable API degrades to 'no update' instead of crashing",
              dead_result == (False, None, None), str(dead_result))

        shutil.rmtree(tmpdir, ignore_errors=True)
        # apply_source_update relaunches the installed script via
        # subprocess.Popen (non-blocking). Give that process a moment to
        # start before deleting the directory out from under it —
        # otherwise the test prints a spurious "can't open file" to stderr
        # that looks like a real failure but is purely a teardown race.
        import time
        time.sleep(0.5)
        shutil.rmtree(app_dir, ignore_errors=True)

    finally:
        server.shutdown()

    print()
    failed = [r for r in results if r[0] == FAIL]
    print(f"{len(results) - len(failed)}/{len(results)} checks passed.")
    if failed:
        print("\nFailures:")
        for _, label, detail in failed:
            print(f"  - {label} {detail}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
