"""Build a source-only delivery ZIP from an explicit allowlist (standard library only)."""

import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import zipfile


ROOT = Path(__file__).resolve().parents[1]
FILES = [
    "run.py", "requirements.txt", "requirements-win-py313.lock.txt",
    "start_windows.cmd", "DEPLOY_WINDOWS_OFFLINE.md", "DEPLOYMENT_VERIFICATION.md",
    "CLAUDE.md", "DEV_PLAN.md", "WINDOWS_OFFLINE_DEPLOYMENT_PLAN.md",
    "项目相关信息表.xlsx", "规模及工作量估计表.xlsx",
    "scripts/build_delivery.py", "scripts/verify_offline.py",
    "app/__init__.py", "app/main.py", "app/database.py", "app/models.py",
    "app/routers/__init__.py", "app/routers/projects.py", "app/routers/config_items.py",
    "app/routers/wbs_tasks.py", "app/routers/estimates.py", "app/routers/export.py",
    "app/services/__init__.py", "app/services/importer.py", "app/services/estimator.py",
    "app/services/task_classifier.py", "app/services/excel_builder.py",
    "app/static/index.html", "app/static/vendor/alpinejs-3.14.1.min.js",
    "app/static/vendor/alpinejs-LICENSE.md", "app/static/vendor/README.md",
]


def build(output: Path):
    contents = {name: (ROOT / name).read_bytes() for name in FILES}
    try:
        commit = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True, stderr=subprocess.DEVNULL,
        ).strip()
        dirty = bool(subprocess.check_output(
            ["git", "status", "--porcelain"], cwd=ROOT, text=True,
        ).strip())
    except (OSError, subprocess.CalledProcessError):
        commit, dirty = None, None
    manifest = {
        "source_commit": commit,
        "source_has_uncommitted_changes": dirty,
        "windows_acceptance": "pending; see DEPLOYMENT_VERIFICATION.md",
        "files": {name: hashlib.sha256(data).hexdigest() for name, data in contents.items()},
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    # Exclusive creation prevents accidentally overwriting an earlier release.
    with zipfile.ZipFile(output, "x", zipfile.ZIP_DEFLATED) as archive:
        for name, data in contents.items():
            archive.writestr("EstimateM/" + name, data)
        archive.writestr(
            "EstimateM/DELIVERY_MANIFEST.json",
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        )
    digest = hashlib.sha256(output.read_bytes()).hexdigest()
    print(f"Created {output}\nSHA256 {digest}\nSource {commit}; dirty={dirty}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    build(parser.parse_args().output.resolve())
