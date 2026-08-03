#!/usr/bin/env python3
"""Cut a Box of Scraps release in one command.

Runs every step of the release runbook in order, so none of them can be
forgotten — in particular the two that are easy to skip and painful to notice:
bumping APP_VERSION (a stale value makes the shipped exe re-offer the same
update forever) and uploading the .sha256 the updater verifies against.

    python release.py 1.7.2 --notes notes.md
    python release.py 1.8.0-beta.1 --notes notes.md --prerelease
    python release.py 1.7.2 --notes notes.md --dry-run     # plan only, no changes

What it does:
    1. Preflight   — repo root, on main, clean tree, version sane and newer,
                     tag unused locally and on the remote, notes present,
                     GitHub credentials work.
    2. Bump+build  — writes APP_VERSION, syntax-checks, builds the one-file exe,
                     then runs the exe with --version-file to confirm the binary
                     really reports the new version.
    3. Publish     — commits, tags, pushes, creates the release, uploads
                     BoxOfScraps.exe and BoxOfScraps.exe.sha256.
    4. Archive     — drafts the previously-visible stable release, so at most one
                     stable release is public at a time. Skipped for prereleases.

Everything before step 3 is reversible; you are asked to confirm before anything
is pushed or published.
"""
import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OWNER, REPO = "lbangrazi43", "Allpowerfulfileprep"
API = f"https://api.github.com/repos/{OWNER}/{REPO}"
SPEC = "BoxOfScraps.spec"
EXE_NAME = "BoxOfScraps.exe"
DIST_EXE = ROOT / "dist" / EXE_NAME
VERSION_RE = re.compile(r"^\d+\.\d+\.\d+(-[0-9A-Za-z][0-9A-Za-z.-]*)?$")

GREEN, RED, YELLOW, DIM, RESET = "\033[32m", "\033[31m", "\033[33m", "\033[2m", "\033[0m"


def step(msg):
    print(f"\n{'=' * 68}\n{msg}\n{'=' * 68}")


def ok(msg):
    print(f"  {GREEN}OK{RESET}    {msg}")


def warn(msg):
    print(f"  {YELLOW}NOTE{RESET}  {msg}")


def fail(msg):
    print(f"\n  {RED}FAILED{RESET}  {msg}\n", file=sys.stderr)
    sys.exit(1)


def run(cmd, capture=True, check=True, **kw):
    """Run a command in the repo root."""
    r = subprocess.run(cmd, cwd=str(ROOT), text=True,
                       capture_output=capture, **kw)
    if check and r.returncode != 0:
        out = (r.stdout or "") + (r.stderr or "")
        fail(f"`{' '.join(cmd)}` exited {r.returncode}\n{out.strip()}")
    return r


def parse_version(tag):
    """Comparable tuple; a final release ranks above its own prereleases.
    Mirrors _parse_version in File.py — deliberately duplicated so this script
    doesn't have to import File.py (and its GUI dependencies)."""
    s = (tag or "").strip().lstrip("vV")
    core, _, pre = s.partition("-")
    parts = []
    for chunk in core.split(".")[:3]:
        m = re.match(r"\d+", chunk)
        parts.append(int(m.group(0)) if m else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts) + (0 if pre else 1,)


def current_version():
    """Read APP_VERSION out of File.py by regex — importing it would drag in
    tkinterdnd2/pywin32, which may not exist in whichever interpreter runs this."""
    src = (ROOT / "File.py").read_text(encoding="utf-8")
    m = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"', src, re.M)
    if not m:
        fail("Couldn't find APP_VERSION in File.py")
    return m.group(1)


def set_version(new):
    path = ROOT / "File.py"
    src = path.read_text(encoding="utf-8")
    out, n = re.subn(r'^(APP_VERSION\s*=\s*)"[^"]+"', rf'\g<1>"{new}"', src,
                     count=1, flags=re.M)
    if n != 1:
        fail("Couldn't rewrite APP_VERSION in File.py")
    path.write_text(out, encoding="utf-8")


def github_token():
    """Ask git's credential helper for a GitHub token (`gh` isn't installed)."""
    r = subprocess.run(["git", "credential", "fill"], cwd=str(ROOT), text=True,
                       input="protocol=https\nhost=github.com\n\n",
                       capture_output=True)
    for line in (r.stdout or "").splitlines():
        if line.startswith("password="):
            return line.split("=", 1)[1].strip()
    fail("No GitHub token from `git credential fill`. Run a `git push` first so "
         "the credential helper has something stored.")


def api(url, token, data=None, method=None, ctype="application/json"):
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("User-Agent", "box-of-scraps-release")
    if data is not None:
        req.add_header("Content-Type", ctype)
    try:
        with urllib.request.urlopen(req, timeout=900) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        fail(f"GitHub {method or 'GET'} {url} -> HTTP {e.code}\n"
             f"{e.read().decode('utf-8', 'replace')[:500]}")


def confirm(prompt, assume_yes):
    if assume_yes:
        print(f"  {DIM}(--yes){RESET} {prompt} -> yes")
        return True
    try:
        return input(f"\n  {prompt} [y/N] ").strip().lower() in ("y", "yes")
    except EOFError:
        return False


# ─────────────────────────────────────────────
def main():
    p = argparse.ArgumentParser(
        description="Build, tag, and publish a Box of Scraps release.")
    p.add_argument("version", help="new version, e.g. 1.7.2 or 1.8.0-beta.1")
    p.add_argument("--notes", required=True, metavar="PATH",
                   help="markdown file with the release notes")
    p.add_argument("--prerelease", action="store_true",
                   help="publish as a prerelease; does not archive anything")
    p.add_argument("--dry-run", action="store_true",
                   help="run preflight and stop, changing nothing")
    p.add_argument("--yes", action="store_true", help="skip confirmations")
    p.add_argument("--skip-build", action="store_true",
                   help="reuse the existing dist exe (must already match)")
    args = p.parse_args()

    new = args.version.lstrip("vV")
    tag = f"v{new}"
    notes_path = Path(args.notes)

    # ── 1. preflight ──────────────────────────────────────────────
    step(f"1/4  Preflight for {tag}")

    if not (ROOT / "File.py").exists() or not (ROOT / SPEC).exists():
        fail(f"{ROOT} doesn't look like the repo root")
    ok(f"repo root: {ROOT}")

    if not VERSION_RE.match(new):
        fail(f"'{new}' isn't a valid version (expected 1.2.3 or 1.2.3-beta.1)")
    ok(f"version format: {new}")

    cur = current_version()
    if parse_version(new) <= parse_version(cur):
        fail(f"{new} is not newer than the current APP_VERSION ({cur})")
    ok(f"newer than current APP_VERSION ({cur} -> {new})")

    branch = run(["git", "branch", "--show-current"]).stdout.strip()
    if branch != "main":
        fail(f"on branch '{branch}' — releases are cut from main")
    ok("on main")

    dirty = run(["git", "status", "--porcelain"]).stdout.strip()
    if dirty:
        fail("working tree isn't clean:\n" + dirty)
    ok("working tree clean")

    if run(["git", "tag", "-l", tag]).stdout.strip():
        fail(f"tag {tag} already exists locally")
    remote_tag = run(["git", "ls-remote", "--tags", "origin", tag]).stdout.strip()
    if remote_tag:
        fail(f"tag {tag} already exists on the remote")
    ok(f"tag {tag} is unused")

    if not notes_path.is_file():
        fail(f"notes file not found: {notes_path}")
    notes = notes_path.read_text(encoding="utf-8").strip()
    if not notes:
        fail(f"notes file is empty: {notes_path}")
    ok(f"release notes: {notes_path} ({len(notes)} chars)")

    token = github_token()
    who = api("https://api.github.com/user", token)
    ok(f"GitHub auth: {who['login']}")

    prev = None
    if not args.prerelease:
        releases = api(f"{API}/releases", token)
        stable = [r for r in releases
                  if not r["draft"] and not r["prerelease"] and r["tag_name"] != tag]
        prev = stable[0] if stable else None
        if prev:
            ok(f"will archive after publishing: {prev['tag_name']}")
        else:
            warn("no published stable release to archive")
    else:
        warn("prerelease — nothing will be archived")

    if args.dry_run:
        print(f"\n{GREEN}Preflight passed.{RESET} Nothing was changed (--dry-run).")
        return

    # ── 2. bump + build ───────────────────────────────────────────
    step(f"2/4  Bump to {new} and build")

    set_version(new)
    ok(f"APP_VERSION = {new}")
    run([sys.executable, "-m", "py_compile", "File.py"])
    ok("File.py compiles")

    if args.skip_build:
        if not DIST_EXE.exists():
            fail(f"--skip-build but {DIST_EXE} doesn't exist")
        warn("reusing the existing dist exe (--skip-build)")
    else:
        # The GLOBAL 3.12 has PyInstaller; the .venv has the runtime deps but not
        # PyInstaller. build/ is removed first because OneDrive can make --clean
        # fail with a PermissionError.
        shutil.rmtree(ROOT / "build", ignore_errors=True)
        print(f"  {DIM}building (a few minutes)…{RESET}")
        r = run(["py", "-3.12", "-m", "PyInstaller", "--noconfirm", SPEC],
                capture=True, check=False)
        if r.returncode != 0:
            tail = ((r.stdout or "") + (r.stderr or "")).strip().splitlines()[-25:]
            fail("PyInstaller failed:\n" + "\n".join(tail))
        ok("PyInstaller exited 0")

    if not DIST_EXE.exists():
        fail(f"{DIST_EXE} was not produced")
    size = DIST_EXE.stat().st_size
    with open(DIST_EXE, "rb") as f:
        if f.read(2) != b"MZ":
            fail("built file isn't a Windows executable")
    ok(f"{EXE_NAME}: {size:,} bytes")

    # Ask the binary what version it thinks it is, rather than assuming the build
    # picked up the bump. The exe is windowed, so it answers via a file.
    with tempfile.TemporaryDirectory() as td:
        vf = Path(td) / "v.txt"
        proc = subprocess.Popen([str(DIST_EXE), "--version-file", str(vf)],
                                cwd=str(DIST_EXE.parent),
                                stdout=subprocess.DEVNULL,
                                stderr=subprocess.DEVNULL)
        try:
            proc.wait(timeout=120)
        except subprocess.TimeoutExpired:
            # A build predating the flag ignores it and opens the GUI instead,
            # which would otherwise hang here with a window on screen.
            proc.kill()
            try:
                proc.wait(timeout=30)
            except subprocess.TimeoutExpired:
                pass
            fail("the built exe ignored --version-file and launched its window "
                 "instead — that build predates the version flag, so it is stale")
        reported = vf.read_text(encoding="utf-8").strip() if vf.exists() else ""
    if reported != new:
        fail(f"the built exe reports version '{reported or '(nothing)'}', "
             f"expected '{new}' — the build did not pick up the bump")
    ok(f"built exe reports version {reported}")

    digest = hashlib.sha256(DIST_EXE.read_bytes()).hexdigest()
    ok(f"sha256 {digest}")

    # ── 3. publish ────────────────────────────────────────────────
    step(f"3/4  Publish {tag}")
    print(f"  commit + tag {tag}, push to origin/main")
    print(f"  create {'PRERELEASE' if args.prerelease else 'release'} {tag}")
    print(f"  upload {EXE_NAME} ({size:,} bytes) + {EXE_NAME}.sha256")
    if prev:
        print(f"  then archive {prev['tag_name']} (set to draft)")
    if not confirm("Proceed? This pushes and publishes.", args.yes):
        run(["git", "checkout", "--", "File.py"])
        fail("aborted — APP_VERSION reverted, nothing pushed")

    run(["git", "add", "File.py"])
    msg = (f"Release {tag}\n\nBump APP_VERSION to {new}.\n\n"
           f"Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>\n")
    run(["git", "commit", "-m", msg])
    ok("committed")
    run(["git", "tag", tag])
    run(["git", "push", "origin", "main"], capture=True)
    run(["git", "push", "origin", tag], capture=True)
    ok(f"pushed main and {tag}")

    body = json.dumps({
        "tag_name": tag,
        "name": f"Box of Scraps {tag}",
        "body": notes,
        "draft": False,
        "prerelease": bool(args.prerelease),
    }).encode()
    rel = api(f"{API}/releases", token, data=body, method="POST")
    ok(f"created release {rel['tag_name']} (id {rel['id']})")

    upload = rel["upload_url"].split("{")[0]
    asset = api(f"{upload}?name={EXE_NAME}", token, data=DIST_EXE.read_bytes(),
                method="POST", ctype="application/octet-stream")
    ok(f"uploaded {asset['name']} ({asset['size']:,} bytes)")

    sums = f"{digest}  {EXE_NAME}\n".encode()
    sasset = api(f"{upload}?name={EXE_NAME}.sha256", token, data=sums,
                 method="POST", ctype="text/plain")
    ok(f"uploaded {sasset['name']}")

    # ── 4. archive the previous stable ────────────────────────────
    step("4/4  Archive the previous stable release")
    if args.prerelease:
        warn("prerelease — previous stable left visible")
    elif prev:
        api(f"{API}/releases/{prev['id']}", token,
            data=json.dumps({"draft": True}).encode(), method="PATCH")
        ok(f"{prev['tag_name']} set to draft (tag and assets preserved)")
    else:
        warn("nothing to archive")

    # ── verify what the world now sees ────────────────────────────
    public = [r for r in api(f"{API}/releases", token) if not r["draft"]]
    latest = api(f"{API}/releases/latest", token)
    print(f"\n{GREEN}Done.{RESET}  {rel['html_url']}")
    print(f"  public releases : {', '.join(r['tag_name'] for r in public) or '(none)'}")
    print(f"  /releases/latest: {latest['tag_name']}  "
          f"(what the in-app updater will offer)")
    stable_public = [r for r in public if not r["prerelease"]]
    if len(stable_public) > 1:
        warn(f"{len(stable_public)} stable releases are public — the runbook "
             f"expects at most one")


if __name__ == "__main__":
    main()
