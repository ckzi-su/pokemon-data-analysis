from pathlib import Path
import argparse
import shutil

import kagglehub

DATASET = "abcsds/pokemon"
ROOT = Path(__file__).resolve().parents[1]
RAW_DIR = ROOT / "data" / "raw"


def download(force: bool = False) -> Path:
    target = RAW_DIR / "Pokemon.csv"
    if target.exists() and not force:
        print(f"Dataset already present: {target}")
        return target

    cache_path = Path(kagglehub.dataset_download(DATASET))
    source = next(cache_path.glob("*.csv"))
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)
    print(f"Saved dataset to {target}")
    return target


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Download the Pokemon dataset from Kaggle.")
    parser.add_argument("--force", action="store_true", help="Re-download even if the file exists.")
    args = parser.parse_args()
    download(force=args.force)
