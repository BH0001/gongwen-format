#!/usr/bin/env python3
"""Portable entry point. Python stdlib only; engine: .NET 8 + Open XML SDK."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def execute(args, **kwargs):
    result = subprocess.run([str(x) for x in args], capture_output=True, text=True,
                            encoding="utf-8", errors="replace", **kwargs)
    return result


def run(args, **kwargs):
    result = execute(args, **kwargs)
    if result.returncode:
        raise RuntimeError((result.stderr + "\n" + result.stdout).strip())
    if result.stderr.strip():
        print(result.stderr.strip(), file=sys.stderr)
    return result.stdout


def executable(name):
    path = shutil.which(name)
    if not path:
        raise RuntimeError(f"未找到 {name}。请安装后重试；未生成或覆盖输出文件。")
    return path


def engine():
    project = ROOT / "scripts/engine/Gongwen.csproj"
    sources = sorted(project.parent.glob("*.cs")) + [project]
    fingerprint = hashlib.sha256(b"".join(p.read_bytes() for p in sources)).hexdigest()[:16]
    cache = Path(os.environ.get("GONGWEN_CACHE_DIR", str(Path.cwd() / ".gongwen-cache"))).resolve() / fingerprint
    dll = cache / "bin/Gongwen.dll"
    dotnet = executable("dotnet")
    if not dll.exists():
        intermediate = (cache / "obj").as_posix() + "/"
        cache.mkdir(parents=True, exist_ok=True)
        run([dotnet, "build", project, "-c", "Release", "--nologo", "-v", "quiet", "--output", cache / "bin",
             "-p:BaseIntermediateOutputPath=" + intermediate, "-p:MSBuildProjectExtensionsPath=" + intermediate],
            env={**os.environ, "DOTNET_CLI_TELEMETRY_OPTOUT": "1"})
    return [dotnet, dll]


def fonts_available():
    names = set()
    if sys.platform == "win32":
        import winreg
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as key:
                    for i in range(winreg.QueryInfoKey(key)[1]):
                        name, _, _ = winreg.EnumValue(key, i)
                        name = re.sub(r"\s*\([^)]*\)$", "", name)
                        names.update(x.strip() for x in name.split("&"))
            except OSError:
                pass
        aliases = {"FangSong": "仿宋", "KaiTi": "楷体", "SimHei": "黑体", "SimSun": "宋体"}
        for en, zh in aliases.items():
            if en in names or zh in names:
                names.update((en, zh))
    elif shutil.which("fc-list"):
        for line in run(["fc-list", "--format", "%{family}\n"]).splitlines():
            names.update(x.strip() for x in line.split(","))
    else:
        raise RuntimeError("无法检测字体。Windows 使用系统字体注册表；macOS/Linux 需要 fc-list。")
    return names


def resolve_profile(value, overrides):
    path = ROOT / "assets/profiles" / (value + ".json")
    if not path.is_file():
        path = Path(value).resolve()
    profile = json.loads(path.read_text(encoding="utf-8"))
    names = fonts_available()
    chosen = {}
    for item in overrides:
        role, sep, name = item.partition("=")
        if not sep or role not in profile["fonts"]:
            raise ValueError("--font 必须为已知角色=已安装字体，例如 title=方正小标宋简体")
        profile["fonts"][role] = [name]
    for role, candidates in profile["fonts"].items():
        chosen[role] = next((x for x in candidates if x.casefold() in {n.casefold() for n in names}), None)
        if chosen[role] is None:
            raise RuntimeError(f"缺少 {role} 字体：{' / '.join(candidates)}。请安装所需字体或用 --font 显式指定替代字体。")
    profile["resolvedFonts"] = chosen
    return profile


def prepare(source, target):
    source, target = Path(source).resolve(), Path(target).resolve()
    if source.suffix.lower() not in (".md", ".markdown", ".txt", ".docx"):
        raise ValueError("首版仅支持 MD、TXT、DOCX；不接收 DOC、PDF。")
    if source == target or target.exists():
        raise ValueError("准备文件必须为新的 .docx 路径，不能覆盖源文件或已有文件。")
    if target.suffix.lower() != ".docx":
        raise ValueError("准备文件扩展名必须为 .docx")
    if not source.is_file():
        raise ValueError("输入文件不存在。")
    target.parent.mkdir(parents=True, exist_ok=True)
    # Convert away from the public path. A failed converter must not leave a
    # partial DOCX, and publishing must not overwrite a concurrently made file.
    with tempfile.TemporaryDirectory(prefix=".gongwen-prepare-", dir=target.parent) as tmp:
        staged = Path(tmp) / "prepared.docx"
        prepare_staged(source, staged)
        if sys.platform == "win32":
            os.rename(staged, target)  # Windows rename refuses an existing target.
        else:
            os.link(staged, target)  # POSIX rename would silently replace it.


def prepare_staged(source, target):
    if source.suffix.lower() == ".docx":
        shutil.copy2(source, target)
        return
    pandoc = executable("pandoc")
    if source.suffix.lower() == ".txt":
        # Every physical line is a paragraph, including deliberate blank lines.
        text = source.read_text(encoding="utf-8-sig")
        version = json.loads(run([pandoc, "-f", "markdown", "-t", "json"], input=""))["pandoc-api-version"]
        blocks = [{"t": "Para", "c": [{"t": "Str", "c": line}]} for line in text.splitlines()]
        payload = json.dumps({"pandoc-api-version": version, "meta": {}, "blocks": blocks}, ensure_ascii=False)
        run([pandoc, "-f", "json", "-t", "docx", "-o", target], input=payload)
    else:
        # Resource paths follow the input, not the caller's current directory.
        run([pandoc, source, "-f", "markdown", "-t", "docx", "--standalone", "--fail-if-warnings",
             "--resource-path", source.parent, "-o", target], cwd=source.parent)


def main():
    parser = argparse.ArgumentParser(description="公文排版：准备 → 检查结构 → 应用配置 → 验证")
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser("prepare", help="将 MD/TXT/DOCX 准备成新 DOCX，供结构映射使用")
    p.add_argument("input"); p.add_argument("--output", required=True)
    p = sub.add_parser("inspect", help="输出全文、段落 ID、角色建议和字形片段")
    p.add_argument("input"); p.add_argument("--output")
    p = sub.add_parser("format", help="另存为公文 DOCX，不覆盖已有文件")
    p.add_argument("input"); p.add_argument("--output", required=True)
    p.add_argument("--profile", default="gbt9704")
    p.add_argument("--map"); p.add_argument("--report")
    p.add_argument("--numbering", choices=["preserve", "official"], default="preserve")
    p.add_argument("--font", action="append", default=[])
    p = sub.add_parser("check", help="验证结构、格式及与原始 DOCX 的内容一致性")
    p.add_argument("input"); p.add_argument("--against"); p.add_argument("--output")
    args = parser.parse_args()
    if args.command == "prepare":
        prepare(args.input, args.output)
        print(json.dumps({"prepared": str(Path(args.output).resolve())}, ensure_ascii=False))
        return
    if args.command in ("inspect", "check"):
        dest = Path(args.output).resolve() if args.output else None
        if dest and (dest.exists() or dest == Path(args.input).resolve()):
            raise ValueError("报告路径已存在，请选择新路径。")
        command = engine() + [args.command, "--input", Path(args.input).resolve()]
        if getattr(args, "against", None):
            command += ["--against", Path(args.against).resolve()]
        result = execute(command)
        out = result.stdout
        # check intentionally returns a nonzero code for a valid failure report.
        # Preserve its JSON before propagating that status to callers.
        if not out.strip():
            raise RuntimeError(result.stderr.strip() or "引擎未返回检查报告。")
        json.loads(out)
        if dest:
            dest.parent.mkdir(parents=True, exist_ok=True)
            with dest.open("x", encoding="utf-8") as report_file:
                report_file.write(out)
        else:
            print(out)
        if result.stderr.strip():
            print(result.stderr.strip(), file=sys.stderr)
        if result.returncode:
            raise SystemExit(result.returncode)
        return
    source, target = Path(args.input).resolve(), Path(args.output).resolve()
    if not source.is_file():
        raise ValueError("输入文件不存在。")
    if target.suffix.lower() != ".docx" or target.exists() or source == target:
        raise ValueError("输出必须为新的 .docx 文件，不能覆盖源文件或已有文件。")
    if args.map and source.suffix.lower() != ".docx":
        raise ValueError("使用结构映射时，先 prepare，再对同一个准备好的 DOCX inspect 和 format。")
    profile = resolve_profile(args.profile, args.font)
    report = Path(args.report).resolve() if args.report else target.with_suffix(".report.json")
    if report.exists() or report in (source, target) or report.suffix.lower() != ".json":
        raise ValueError("检查报告必须为新的 .json 路径。")
    target.parent.mkdir(parents=True, exist_ok=True)
    report.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".gongwen-", dir=target.parent) as tmp:
        tmp = Path(tmp)
        prepared = source
        if source.suffix.lower() != ".docx":
            prepared = tmp / "prepared.docx"
            prepare(source, prepared)
        cfg = tmp / "profile.json"
        cfg.write_text(json.dumps(profile, ensure_ascii=False), encoding="utf-8")
        command = engine() + ["format", "--input", prepared, "--output", target,
                              "--profile", cfg, "--report", report, "--numbering", args.numbering]
        if args.map:
            command += ["--map", Path(args.map).resolve()]
        print(run(command))


if __name__ == "__main__":
    try:
        main()
    except (RuntimeError, ValueError, OSError, KeyError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        sys.exit(2)
