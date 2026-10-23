# -*- coding: utf-8 -*-
"""通用工具：配置、路径、日志、AppleScript/Chrome 交互。"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
# 配置文件优先根目录；若使用“个人使用/”分包，则读个人目录里的配置
_ROOT_CONFIG = ROOT / "gk2027_config.json"
_PERSONAL_CONFIG = ROOT / "个人使用" / "gk2027_config.json"
CONFIG_PATH = _ROOT_CONFIG if _ROOT_CONFIG.exists() else (_PERSONAL_CONFIG if _PERSONAL_CONFIG.exists() else _ROOT_CONFIG)


def load_config() -> dict:
    # 新克隆的仓库没有 gk2027_config.json，先读脱敏示例；init 会写正式配置
    path = CONFIG_PATH if CONFIG_PATH.exists() else (ROOT / "gk2027_config.example.json")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


CFG = load_config()


def P(key: str) -> Path:
    """把配置里的相对路径转成绝对路径。"""
    return (ROOT / CFG[key]).resolve()


def ensure_dir(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def backup(path: Path) -> Path:
    """在同目录生成 .bak-时间戳 备份。"""
    if not path.exists():
        return path
    dst = path.with_name(path.name + time.strftime(".bak-%Y%m%d-%H%M%S"))
    shutil.copy2(path, dst)
    return dst


def fmt_int(n) -> str:
    try:
        return format(int(float(n)), ",")
    except Exception:
        return str(n)


def log(msg: str = "") -> None:
    print(msg, flush=True)


def banner(title: str) -> None:
    line = "=" * 64
    log("\n" + line)
    log(title)
    log(line)


def ask(prompt: str, default: str = "") -> str:
    suffix = f" [{default}]" if default else ""
    try:
        val = input(f"{prompt}{suffix}: ").strip()
    except EOFError:
        return default
    return val or default


def ask_yes(prompt: str, default: bool = False) -> bool:
    d = "Y/n" if default else "y/N"
    try:
        val = input(f"{prompt} [{d}]: ").strip().lower()
    except EOFError:
        return default
    if not val:
        return default
    return val in ("y", "yes", "是", "好")


class ChromeError(RuntimeError):
    pass


def run_osascript(script: str, timeout: int = 120) -> str:
    proc = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=timeout)
    if proc.returncode != 0:
        raise ChromeError(proc.stderr.strip() or proc.stdout.strip())
    return proc.stdout.strip()


def chrome_tabs() -> list[tuple[str, str]]:
    script = (
        'tell application "Google Chrome"\n'
        '  set out to ""\n'
        '  repeat with w in windows\n'
        '    repeat with t in tabs of w\n'
        '      set u to URL of t\n'
        '      if u contains "bm.scs.gov.cn" then set out to out & u & "|" & (title of t) & linefeed\n'
        '    end repeat\n'
        '  end repeat\n'
        '  return out\n'
        'end tell'
    )
    out = run_osascript(script, timeout=30)
    tabs = []
    for line in out.splitlines():
        if "|" in line:
            u, t = line.split("|", 1)
            tabs.append((u.strip(), t.strip()))
    return tabs


def chrome_js(js: str, tab_contains: str = "person_position.html", timeout: int = 120) -> str:
    """在 Chrome 指定标签页的主世界执行 JS，返回结果。

    通过“先写入临时文件、AppleScript 再 read”的方式，避免转义问题。
    """
    fd, js_path = tempfile.mkstemp(prefix="gk2027_js_", suffix=".js")
    os.close(fd)
    Path(js_path).write_text(js, encoding="utf-8")
    try:
        script = (
            'set jsCode to read POSIX file "%s" as «class utf8»\n'
            'tell application "Google Chrome"\n'
            '  repeat with w in windows\n'
            '    repeat with t in tabs of w\n'
            '      if URL of t contains "%s" then\n'
            '        return execute t javascript jsCode\n'
            '      end if\n'
            '    end repeat\n'
            '  end repeat\n'
            '  return "__NOT_FOUND__"\n'
            'end tell' % (js_path, tab_contains)
        )
        out = run_osascript(script, timeout=timeout)
    except ChromeError as e:
        msg = str(e)
        if "JavaScript" in msg and ("关闭" in msg or "turned off" in msg):
            raise ChromeError(
                "Chrome 禁止了 AppleScript 执行 JavaScript。\n"
                "请在 Chrome 菜单栏：显示 → 开发者 → 勾选“允许 Apple 事件中的 JavaScript”，然后重试。"
            ) from e
        raise
    finally:
        try:
            os.unlink(js_path)
        except OSError:
            pass
    if out == "__NOT_FOUND__":
        raise ChromeError(
            "没有找到已打开的职位查询页（person_position.html）。\n"
            "请在 Chrome 打开：个人中心 → 职位查询 → 按代码选择职位，保持页面打开后重试。"
        )
    return out


def chrome_open(url: str) -> None:
    run_osascript(f'tell application "Google Chrome" to open location "{url}"', timeout=30)


def check_chrome() -> bool:
    try:
        run_osascript('tell application "Google Chrome" to count windows', timeout=15)
        return True
    except Exception:
        return False
