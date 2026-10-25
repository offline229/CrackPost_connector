"""Windows 窗口 / 任务栏图标：可靠设置，避免 Tk 默认羽毛与模糊缩放。

参考 TablePet/src/utils/window_icon.py：iconbitmap + Win32 WM_SETICON，
并按窗口 DPI 生成对应尺寸的临时 .ico。
"""

from __future__ import annotations

import ctypes
import os
import shutil
import tempfile
import time
from pathlib import Path

WM_SETICON = 0x0080
ICON_SMALL = 0
ICON_BIG = 1
IMAGE_ICON = 1
LR_LOADFROMFILE = 0x00000010
GA_ROOT = 2
GA_ROOTOWNER = 3
SM_CXICON = 11
SM_CXSMICON = 49

_TEMP_ICO: dict[tuple[str, int], str] = {}
_HICON: dict[tuple[str, int], int] = {}


def _ascii_path_copy(src: Path, dest_name: str) -> str:
    try:
        str(src).encode("ascii")
        return str(src.resolve())
    except UnicodeEncodeError:
        pass
    dest = Path(tempfile.gettempdir()) / dest_name
    shutil.copy2(src, dest)
    return str(dest)


def _parse_hwnd(value) -> int:
    if value is None or value == "" or value == 0:
        return 0
    if isinstance(value, int):
        return value
    s = str(value).strip()
    try:
        return int(s, 16) if s.lower().startswith("0x") else int(s, 0)
    except Exception:
        return 0


def _window_dpi(hwnd: int) -> int:
    try:
        dpi = int(ctypes.windll.user32.GetDpiForWindow(int(hwnd)))
        if dpi > 0:
            return dpi
    except Exception:
        pass
    try:
        hdc = ctypes.windll.user32.GetDC(int(hwnd) if hwnd else 0)
        dpi = int(ctypes.windll.gdi32.GetDeviceCaps(hdc, 88))
        ctypes.windll.user32.ReleaseDC(int(hwnd) if hwnd else 0, hdc)
        if dpi > 0:
            return dpi
    except Exception:
        pass
    return 96


def _icon_sizes_for_hwnd(hwnd: int) -> tuple[int, int]:
    dpi = _window_dpi(hwnd)
    try:
        get_for_dpi = ctypes.windll.user32.GetSystemMetricsForDpi
        small = int(get_for_dpi(SM_CXSMICON, dpi) or 0)
        big = int(get_for_dpi(SM_CXICON, dpi) or 0)
        if small > 0 and big > 0:
            return small, big
    except Exception:
        pass
    small = max(16, int(round(16 * dpi / 96.0)))
    big = max(32, int(round(32 * dpi / 96.0)))
    return small, big


def _source_rgba(png_path: str | Path | None, ico_path: str | Path):
    from PIL import Image

    if png_path and Path(png_path).is_file():
        return Image.open(png_path).convert("RGBA")
    return Image.open(ico_path).convert("RGBA")


def _temp_ico_nearest(src_img, size: int, cache_key: str) -> str:
    from PIL import Image

    key = (cache_key, int(size))
    cached = _TEMP_ICO.get(key)
    if cached and os.path.isfile(cached):
        return cached
    img = src_img.resize((size, size), Image.Resampling.NEAREST)
    path = str(Path(tempfile.gettempdir()) / f"CrackPost_icon_{size}px.ico")
    img.save(path, format="ICO", sizes=[(size, size)])
    _TEMP_ICO[key] = path
    return path


def _load_hicon_size(path: str, size: int) -> int:
    key = (path, int(size))
    h = _HICON.get(key)
    if h:
        return h
    user32 = ctypes.windll.user32
    h = int(user32.LoadImageW(0, path, IMAGE_ICON, size, size, LR_LOADFROMFILE) or 0)
    if h:
        _HICON[key] = h
    return h


def _tk_hwnd_candidates(root) -> list[tuple[str, int]]:
    seen: set[int] = set()
    out: list[tuple[str, int]] = []

    def add(name: str, hwnd: int) -> None:
        hwnd = int(hwnd or 0)
        if hwnd and hwnd not in seen:
            seen.add(hwnd)
            out.append((name, hwnd))

    try:
        root.update_idletasks()
    except Exception:
        pass

    try:
        add("winfo_id", int(root.winfo_id()))
    except Exception:
        pass

    try:
        add("wm_frame", _parse_hwnd(root.wm_frame()))
    except Exception:
        pass

    try:
        user32 = ctypes.windll.user32
        base = out[0][1] if out else 0
        if base:
            add("parent", int(user32.GetParent(base) or 0))
            add("ancestor_root", int(user32.GetAncestor(base, GA_ROOT) or 0))
            add("ancestor_owner", int(user32.GetAncestor(base, GA_ROOTOWNER) or 0))
    except Exception as e:
        print(f"[window_icon] hwnd enumerate failed: {e}")

    return out


def apply_win32_icon(
    hwnd: int,
    ico_path: str | Path,
    png_path: str | Path | None = None,
) -> bool:
    if not hwnd:
        return False
    ico_path = Path(ico_path)
    if not ico_path.is_file() and not (png_path and Path(png_path).is_file()):
        return False

    small, big = _icon_sizes_for_hwnd(hwnd)
    user32 = ctypes.windll.user32
    h_small = 0
    h_big = 0
    try:
        from PIL import Image  # noqa: F401 — used via _source_rgba

        src = _source_rgba(png_path, ico_path)
        cache_key = str(Path(png_path or ico_path).resolve())
        small_ico = _temp_ico_nearest(src, small, cache_key)
        big_ico = _temp_ico_nearest(src, big, cache_key)
        h_small = _load_hicon_size(small_ico, small)
        h_big = _load_hicon_size(big_ico, big)
    except Exception as e:
        print(f"[window_icon] dpi icon build failed: {e}")

    if not h_small or not h_big:
        path = str(ico_path.resolve()) if ico_path.is_file() else ""
        if path:
            if not h_small:
                h_small = int(user32.LoadImageW(0, path, IMAGE_ICON, small, small, LR_LOADFROMFILE) or 0)
            if not h_big:
                h_big = int(user32.LoadImageW(0, path, IMAGE_ICON, big, big, LR_LOADFROMFILE) or 0)

    if not h_small and not h_big:
        return False

    if h_small:
        user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, h_small)
    if h_big:
        user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, h_big)
    return True


def apply_tk_window_icon(
    root,
    ico_path: str | Path,
    png_path: str | Path | None = None,
    *,
    schedule_retries: bool = True,
) -> bool:
    """给 Tk 根窗口固定自定义图标（标题栏 + 任务栏）。"""
    from PIL import Image  # noqa: F401

    ico_path = Path(ico_path)
    if not ico_path.is_file():
        print(f"[window_icon] missing ico: {ico_path}")
        return False

    png_path = Path(png_path) if png_path else None
    ok = False
    root._crackpost_icon_last_ok = getattr(root, "_crackpost_icon_last_ok", 0.0)  # noqa: SLF001

    def _win32(tag: str = "now", force: bool = False) -> bool:
        nonlocal ok
        now = time.monotonic()
        if not force and tag != "immediate":
            last = float(getattr(root, "_crackpost_icon_last_ok", 0.0) or 0.0)
            if now - last < 2.0:
                return True
        try:
            if not bool(root.winfo_exists()):
                return False
        except Exception:
            return False
        try:
            cands = _tk_hwnd_candidates(root)
            if not cands:
                return False
            ref_hwnd = next((h for n, h in cands if n == "wm_frame"), cands[0][1])
            small, big = _icon_sizes_for_hwnd(ref_hwnd)

            try:
                src = _source_rgba(png_path, ico_path)
                cache_key = str((png_path or ico_path).resolve())
                neat = _temp_ico_nearest(src, big, cache_key)
                neat_ascii = _ascii_path_copy(Path(neat), f"CrackPost_tk_{big}.ico")
                root.iconbitmap(neat_ascii)
                try:
                    root.iconbitmap(default=neat_ascii)
                except Exception:
                    pass
            except Exception as e:
                if tag == "immediate":
                    print(f"[window_icon] iconbitmap failed: {e}")

            any_ok = False
            for _name, hwnd in cands:
                if apply_win32_icon(hwnd, ico_path, png_path=png_path):
                    any_ok = True
            if any_ok:
                ok = True
                root._crackpost_icon_last_ok = now  # noqa: SLF001
                if tag == "immediate":
                    print(f"[window_icon] ok dpi_icon={small}/{big}")
            elif tag == "immediate":
                print(f"[window_icon] win32 failed dpi_icon={small}/{big}")
            return any_ok
        except Exception as e:
            if tag == "immediate":
                print(f"[window_icon] failed: {e}")
            return False

    root._crackpost_icon_apply = _win32  # noqa: SLF001

    def _retry_tagged(tag: str) -> None:
        try:
            if not bool(root.winfo_exists()):
                return
        except Exception:
            return
        fn = getattr(root, "_crackpost_icon_apply", None)
        if callable(fn):
            try:
                fn(tag)
            except Exception:
                pass

    if _win32("immediate", force=True):
        ok = True

    if not schedule_retries:
        return ok

    if not getattr(root, "_crackpost_icon_cmds_ready", False):
        def _cb_80():
            _retry_tagged("80ms")

        def _cb_400():
            _retry_tagged("400ms")

        def _cb_map():
            _retry_tagged("map")

        root._crackpost_icon_cb_80 = _cb_80  # noqa: SLF001
        root._crackpost_icon_cb_400 = _cb_400  # noqa: SLF001
        root._crackpost_icon_cb_map = _cb_map  # noqa: SLF001
        root._crackpost_icon_cmd_80 = root.register(_cb_80)  # noqa: SLF001
        root._crackpost_icon_cmd_400 = root.register(_cb_400)  # noqa: SLF001
        root._crackpost_icon_cmd_map = root.register(_cb_map)  # noqa: SLF001
        root._crackpost_icon_cmds_ready = True  # noqa: SLF001

    try:
        for attr, delay, cmd_attr in (
            ("_crackpost_icon_after_80", 80, "_crackpost_icon_cmd_80"),
            ("_crackpost_icon_after_400", 400, "_crackpost_icon_cmd_400"),
        ):
            old = getattr(root, attr, None)
            if old is not None:
                try:
                    root.after_cancel(old)
                except Exception:
                    pass
            cmd = getattr(root, cmd_attr, None)
            try:
                setattr(root, attr, root.tk.call("after", delay, cmd) if cmd else None)
            except Exception:
                setattr(root, attr, None)

        if not getattr(root, "_crackpost_icon_map_bound", False):
            root._crackpost_icon_map_bound = True  # noqa: SLF001

            def _on_map(_event=None):
                if _event is not None and _event.widget is not root:
                    return
                try:
                    if not bool(root.winfo_exists()):
                        return
                except Exception:
                    return
                old = getattr(root, "_crackpost_icon_after_map", None)
                if old is not None:
                    try:
                        root.after_cancel(old)
                    except Exception:
                        pass
                cmd = getattr(root, "_crackpost_icon_cmd_map", None)
                try:
                    root._crackpost_icon_after_map = (  # noqa: SLF001
                        root.tk.call("after", 40, cmd) if cmd else None
                    )
                except Exception:
                    root._crackpost_icon_after_map = None  # noqa: SLF001

            root._crackpost_icon_on_map = _on_map  # noqa: SLF001
            root.bind("<Map>", root._crackpost_icon_on_map, add="+")
    except Exception:
        pass
    return ok
