#!/usr/bin/env python3
import argparse
from pathlib import Path

DEFAULT_EXTS = {".xlsx", ".pdf"}


def collect_targets(directory: Path, prefix: str, exts: set[str]) -> list[Path]:
    targets = []
    for path in directory.iterdir():
        if not path.is_file():
            continue
        if not path.name.startswith(prefix):
            continue
        if path.suffix not in exts:
            continue
        targets.append(path)
    return sorted(targets)


def main() -> None:
    parser = argparse.ArgumentParser(description="生成物のXLSX/PDFをまとめて削除します")
    parser.add_argument("--dir", default="examples", help="削除対象のディレクトリ")
    parser.add_argument("--prefix", default="captions", help="削除対象のファイル接頭辞")
    parser.add_argument("--dry-run", action="store_true", help="削除せずに対象だけ表示")
    args = parser.parse_args()

    directory = Path(args.dir)
    if not directory.exists():
        raise SystemExit(f"ディレクトリが存在しません: {directory}")

    targets = collect_targets(directory, args.prefix, DEFAULT_EXTS)
    if not targets:
        print("削除対象がありません")
        return

    if args.dry_run:
        for path in targets:
            print(path)
        return

    for path in targets:
        path.unlink()
        print(f"削除: {path}")


if __name__ == "__main__":
    main()
