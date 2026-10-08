"""Create a source-only, installable skill ZIP; omit build/test caches."""
import argparse
from pathlib import Path
import zipfile


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    out = Path(args.output).resolve()
    if out.exists():
        raise SystemExit("输出压缩包已存在，请使用新路径。")
    out.parent.mkdir(parents=True, exist_ok=True)
    excluded = {"bin", "obj", "__pycache__", ".gongwen-cache", ".git"}
    source_dirs = {"agents", "assets", "references", "scripts", "examples"}
    source_files = {"SKILL.md", "README.md", "LICENSE", "requirements-dev.txt"}
    files = sorted(
        p for p in root.rglob("*")
        if p.is_file() and not p.is_symlink() and p.resolve() != out
        and not excluded.intersection(p.relative_to(root).parts)
        and (p.relative_to(root).parts[0] in source_dirs or p.relative_to(root).as_posix() in source_files)
    )
    if not (root / "SKILL.md").is_file():
        raise SystemExit("未找到技能入口。")
    with zipfile.ZipFile(out, "x", zipfile.ZIP_DEFLATED) as archive:
        for path in files:
            archive.write(path, Path(root.name) / path.relative_to(root))
    print(f"{out}\n{len(files)} files")


if __name__ == "__main__":
    main()
