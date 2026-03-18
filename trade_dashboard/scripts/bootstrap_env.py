from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path

MIRRORS = [
    "https://pypi.org/simple",
    "https://pypi.tuna.tsinghua.edu.cn/simple",
    "https://mirrors.aliyun.com/pypi/simple",
]


def required_packages(requirements_file: Path) -> list[str]:
    lines = requirements_file.read_text(encoding="utf-8").splitlines()
    pkgs: list[str] = []
    for raw in lines:
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        name = line.split("==")[0].split(">=")[0].split("<=")[0].strip()
        pkgs.append(name)
    return pkgs


def missing_modules(pkgs: list[str]) -> list[str]:
    mapping = {
        "pyyaml": "yaml",
    }
    missing: list[str] = []
    for pkg in pkgs:
        module = mapping.get(pkg.lower(), pkg.replace("-", "_"))
        if importlib.util.find_spec(module) is None:
            missing.append(pkg)
    return missing


def install_with_mirrors(req_file: Path) -> int:
    for mirror in MIRRORS:
        cmd = [
            sys.executable,
            "-m",
            "pip",
            "install",
            "-r",
            str(req_file),
            "-i",
            mirror,
        ]
        print(f"\n[INFO] 尝试安装依赖，镜像: {mirror}")
        result = subprocess.run(cmd)
        if result.returncode == 0:
            print("[OK] 依赖安装成功")
            return 0
        print(f"[WARN] 镜像安装失败: {mirror}")
    return 1


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    req = root / "requirements.txt"
    if not req.exists():
        print(f"[ERROR] 找不到 requirements.txt: {req}")
        return 2

    packages = required_packages(req)
    miss = missing_modules(packages)
    if not miss:
        print("[OK] 依赖已完整，无需安装")
        return 0

    print("[INFO] 缺失依赖:", ", ".join(miss))
    code = install_with_mirrors(req)
    if code != 0:
        print("[ERROR] 自动安装失败，请检查网络/代理设置后重试")
        print("[TIP] 你可以手动执行: python -m pip install -r requirements.txt -i https://pypi.org/simple")
    return code


if __name__ == "__main__":
    raise SystemExit(main())
