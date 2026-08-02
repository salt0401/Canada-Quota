# -*- coding: utf-8 -*-
"""
Publish data/canada_trq_tracker.xlsx to the rolling 'latest' GitHub release.

This replaces the `softprops/action-gh-release` step in
.github/workflows/weekly_scrape.yml, which ran this until the weekly job moved
to the MEPS company server. That server has no GitHub CLI installed and
deliberately does not get one (see docs/SERVER_DEPLOYMENT.md), so the same work
is done here against the REST API with `requests` -- already a pinned
dependency, so this adds nothing to install.

THIS IS THE STEP THE END USER ACTUALLY SEES. download_latest.bat fetches
    https://github.com/salt0401/Canada-Quota/releases/latest/download/canada_trq_tracker.xlsx
so the release asset -- not the git commit -- is the delivery surface. Semantics
therefore match the action it replaces, deliberately and exactly:

  * the release is created if missing, and its name/body/make_latest are
    re-asserted every run (the action sent them every run too),
  * make_latest stays TRUE, because `/releases/latest/download/` resolves to
    whichever release is flagged latest -- get this wrong and every colleague's
    downloader 404s or silently serves an older release,
  * an asset of the same name is deleted before re-upload, which is what the
    action's default clobber behaviour did,
  * anything unexpected raises, so the caller exits non-zero and the scheduled
    task reports failure rather than silently leaving a stale workbook up.

The token is read from a file, never from a command-line argument: arguments
are visible in the process list to every account on a shared machine.

Usage:
    python tools/publish_release.py --token-file C:\\path\\to\\token
    python tools/publish_release.py --token-file ... --dry-run
"""

import argparse
import io
import os
import sys
import time

import requests

DEFAULT_REPO = "salt0401/Canada-Quota"
DEFAULT_TAG = "latest"
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_ASSET = os.path.join(_PROJECT_ROOT, "data", "canada_trq_tracker.xlsx")

RELEASE_NAME = "Latest TRQ Data"
RELEASE_BODY = "Weekly TRQ data update. Download the Excel file below."

API_ROOT = "https://api.github.com"
UPLOAD_ROOT = "https://uploads.github.com"
TIMEOUT = 120


def read_token(token_file: str) -> str:
    """Read the PAT from a file. Whitespace-stripped; must be non-empty."""
    with io.open(token_file, "r", encoding="utf-8-sig") as f:
        token = f.read().strip()
    if not token:
        raise RuntimeError(f"Token file {token_file} is empty.")
    return token


def make_session(token: str) -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "MEPS-CanadaQuota-Publisher",
    })
    return s


def get_release(session: requests.Session, repo: str, tag: str):
    r = session.get(f"{API_ROOT}/repos/{repo}/releases/tags/{tag}", timeout=TIMEOUT)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()


def create_release(session: requests.Session, repo: str, tag: str) -> dict:
    r = session.post(
        f"{API_ROOT}/repos/{repo}/releases",
        json={
            "tag_name": tag,
            "name": RELEASE_NAME,
            "body": RELEASE_BODY,
            "make_latest": "true",
        },
        timeout=TIMEOUT)
    r.raise_for_status()
    return r.json()


def assert_release_metadata(session: requests.Session, repo: str, release: dict):
    """Re-assert name/body/make_latest, as the replaced action did every run.

    make_latest is the one that matters: download_latest.bat resolves
    /releases/latest/download/, so a release that stopped being flagged latest
    would silently start serving a different (or no) file to every colleague.
    """
    r = session.patch(
        f"{API_ROOT}/repos/{repo}/releases/{release['id']}",
        json={
            "name": RELEASE_NAME,
            "body": RELEASE_BODY,
            "make_latest": "true",
            "draft": False,
            "prerelease": False,
        },
        timeout=TIMEOUT)
    r.raise_for_status()
    return r.json()


def delete_asset(session: requests.Session, repo: str, asset_id: int):
    r = session.delete(
        f"{API_ROOT}/repos/{repo}/releases/assets/{asset_id}", timeout=TIMEOUT)
    # 404 means someone else already removed it; that is the state we wanted.
    if r.status_code not in (204, 404):
        r.raise_for_status()


def upload_asset(session: requests.Session, repo: str, release_id: int,
                 path: str, retries: int = 3) -> dict:
    """Upload one file, retrying on transient transport failures.

    The retry is not defensive padding. The old asset is deleted before the new
    one goes up, so a single dropped connection here would leave the release
    with NO workbook at all -- and every colleague's download_latest.bat would
    fail until someone noticed. Retry, then raise loudly.
    """
    name = os.path.basename(path)
    with io.open(path, "rb") as f:
        blob = f.read()
    last = None
    for attempt in range(1, retries + 1):
        try:
            r = session.post(
                f"{UPLOAD_ROOT}/repos/{repo}/releases/{release_id}/assets",
                params={"name": name},
                data=blob,
                headers={"Content-Type": "application/octet-stream"},
                timeout=TIMEOUT)
            r.raise_for_status()
            return r.json()
        except requests.RequestException as e:
            last = e
            if attempt == retries:
                break
            time.sleep(2 ** attempt)
    raise RuntimeError(f"Upload of {name} failed after {retries} attempts: {last}")


def publish(repo: str, tag: str, token: str, asset_path: str,
            dry_run: bool = False) -> dict:
    if not os.path.exists(asset_path):
        raise RuntimeError(
            f"{asset_path} does not exist. Did canada_trq_tracker.py actually "
            f"complete? Refusing to touch the release.")
    size_kb = os.path.getsize(asset_path) / 1024.0
    name = os.path.basename(asset_path)

    session = make_session(token)
    release = get_release(session, repo, tag)

    if release is None:
        if dry_run:
            print(f"  [dry-run] release '{tag}' does not exist; would create it "
                  f"and add {name} ({size_kb:.0f} KB)")
            return {"created": True, "uploaded": [], "dry_run": True}
        print(f"  Release '{tag}' not found - creating it")
        release = create_release(session, repo, tag)
    elif not dry_run:
        assert_release_metadata(session, repo, release)

    existing = {a["name"]: a["id"] for a in release.get("assets", [])}
    if dry_run:
        verb = "replace" if name in existing else "add"
        print(f"  [dry-run] would {verb} {name} ({size_kb:.0f} KB) on release '{tag}'")
        return {"created": False, "uploaded": [], "dry_run": True}

    if name in existing:
        delete_asset(session, repo, existing[name])
    upload_asset(session, repo, release["id"], asset_path)
    print(f"  Uploaded {name} ({size_kb:.0f} KB)")
    return {"created": False, "uploaded": [name], "dry_run": False}


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Upload the TRQ workbook to the rolling GitHub release.")
    parser.add_argument("--repo", default=DEFAULT_REPO,
                        help=f"owner/name (default: {DEFAULT_REPO})")
    parser.add_argument("--tag", default=DEFAULT_TAG,
                        help=f"release tag (default: {DEFAULT_TAG})")
    parser.add_argument("--asset", default=DEFAULT_ASSET,
                        help="path to the workbook to publish")
    parser.add_argument("--token-file", required=True,
                        help="file containing the GitHub token (never pass the "
                             "token itself: arguments are visible in the "
                             "process list)")
    parser.add_argument("--dry-run", action="store_true",
                        help="report what would happen, change nothing")
    args = parser.parse_args(argv)

    token = read_token(args.token_file)
    result = publish(args.repo, args.tag, token, args.asset, dry_run=args.dry_run)
    if args.dry_run:
        print("  Dry run complete - nothing was changed on the release.")
    else:
        print(f"  Release '{args.tag}' updated: {len(result['uploaded'])} asset(s).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
