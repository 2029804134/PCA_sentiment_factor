from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]

BUILD_STEPS = [
    ("Build PCA sentiment factors", PROJECT_ROOT / "code" / "Get_PCAMS.py"),
]

ANALYSIS_STEPS = [
    ("Run CSI 300 regressions", PROJECT_ROOT / "code" / "Baseline.py"),
    ("Run CSI All regressions", PROJECT_ROOT / "code" / "Robust.py"),
    ("Generate descriptive statistics and figures", PROJECT_ROOT / "code" / "Stats.py"),
    ("Run heterogeneity analysis", PROJECT_ROOT / "code" / "heterogeneity.py"),
    ("Run mediation bootstrap", PROJECT_ROOT / "code" / "mediation_bootstrap.py"),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run the paper replication pipeline from the repository root."
    )
    parser.add_argument(
        "--stage",
        choices=["all", "build", "analysis"],
        default="all",
        help="Choose which part of the pipeline to execute.",
    )
    return parser.parse_args()


def get_steps(stage: str) -> list[tuple[str, Path]]:
    if stage == "build":
        return BUILD_STEPS
    if stage == "analysis":
        return ANALYSIS_STEPS
    return BUILD_STEPS + ANALYSIS_STEPS


def run_step(name: str, script_path: Path) -> None:
    if not script_path.exists():
        raise FileNotFoundError(f"Missing script: {script_path}")

    print(f"\n[RUN] {name}")
    print(f"      {script_path.relative_to(PROJECT_ROOT)}")
    subprocess.run(
        [sys.executable, str(script_path)],
        cwd=PROJECT_ROOT,
        check=True,
    )


def main() -> None:
    args = parse_args()
    for name, script_path in get_steps(args.stage):
        run_step(name, script_path)
    print("\nPipeline finished successfully.")


if __name__ == "__main__":
    main()
