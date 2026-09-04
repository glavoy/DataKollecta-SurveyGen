"""Byte-compare the packages this tool generates before and after a change.

Generates a package from a git revision and from the working tree, unzips
both, and `diff -r`s them -- so a refactor can be shown to have changed
nothing, and a deliberate change can be shown to have changed only what was
intended and nowhere else.

The README described this as a manual procedure; nothing in tests/ reads a
config*.json, because they need real dictionaries that live outside the repo.

    .venv/bin/python tools/package_diff.py HEAD config_prism_css.json
    .venv/bin/python tools/package_diff.py HEAD config.json

`config_r21.json` points at a dictionary filename that no longer exists, so
only those two configs are runnable.
"""
import json, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

repo = Path(__file__).resolve().parent.parent

def build(tree: Path, config_path: Path, out: Path) -> Path:
    out.mkdir(parents=True, exist_ok=True)
    cfg = json.loads(config_path.read_text())
    cfg["outputPath"] = str(out)
    tmp_cfg = out / "config.json"
    tmp_cfg.write_text(json.dumps(cfg))
    r = subprocess.run(
        [sys.executable, "main.py", "--config", str(tmp_cfg)],
        cwd=tree, capture_output=True, text=True,
    )
    zips = list(out.glob("*.zip"))
    if not zips:
        print(r.stdout[-3000:]); print(r.stderr[-2000:])
        raise SystemExit(f"no zip produced in {out}")
    dest = out / "unzipped"
    with zipfile.ZipFile(zips[0]) as z:
        z.extractall(dest)
    return dest

def main(rev: str, config_name: str):
    work = Path(tempfile.mkdtemp(prefix="pkgdiff-"))
    base_tree = work / "base"
    base_tree.mkdir()
    subprocess.run(f"git archive {rev} | tar -x -C {base_tree}",
                   shell=True, cwd=repo, check=True)
    cfg = repo / config_name
    before = build(base_tree, cfg, work / "before")
    after = build(repo, cfg, work / "after")
    print(f"--- diff -r  ({rev} -> working tree)  config={config_name} ---")
    d = subprocess.run(["diff", "-r", str(before), str(after)],
                       capture_output=True, text=True)
    print(d.stdout or "(no difference)")
    print(f"exit={d.returncode}   workdir={work}")

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
