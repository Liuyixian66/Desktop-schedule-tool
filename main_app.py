# -*- coding: utf-8 -*-
"""
桌面课表工具 · 主程序（Pygame）
功能：
  1. 启动后读取 schedule_data.json，按当前日期 + 单双周推算今日课程；
  2. 两种布局：屏幕右侧竖排（默认）/ 屏幕顶部横排（可在配置程序中切换）；
     开启"按屏幕分辨率自动缩放"后：右侧模式高度占屏高 4/5，顶部模式宽度占屏宽 3/4；
  3. 始终保持置顶（周期性重申，防遮挡）；仅当其他程序处于【真正的全屏】
     （无边框且覆盖整个屏幕）时自动最小化，退出全屏自动恢复；
     普通最大化窗口、桌面、任务栏、任务管理器均不会触发最小化；
  4. 配置被修改时自动热加载（无需重启）；
  5. 系统托盘图标：左键单击最小化/恢复，右键菜单可最小化/恢复、打开配置程序、退出；
  6. 卡片右上角 "…" 打开配置程序、"×" 退出，按 Esc 也可退出。

用法：
  python main_app.py                启动
  python main_app.py --smoke [日期]  离屏渲染一帧保存预览（调试用，日期可选 YYYY-MM-DD）
  python main_app.py --autostart     注册开机自启后退出
  python main_app.py --no-autostart  取消开机自启后退出

适用：Python 3.13，Windows（依赖 pygame ≥ 2.0）。
"""

import argparse
import ctypes
import datetime
import os
import subprocess
import sys
import threading
from ctypes import wintypes

import schedule_common as sc

# ---------------- Win32 常量与句柄 ----------------
# 注意：必须用显式 WinDLL 实例（use_last_error=True）。
# 实测 ctypes.windll 默认实例调用 SetWindowPos(HWND_TOPMOST) 会静默返回 0，
# 导致置顶失效；显式实例调用正常返回 1。
try:
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    shell32 = ctypes.WinDLL("shell32", use_last_error=True)
except (AttributeError, OSError):  # 非 Windows
    user32 = None
    kernel32 = None
    shell32 = None

HWND_TOPMOST = -1
SDL_ALWAYS_ON_TOP = 0x00008000   # SDL_WINDOW_ALWAYS_ON_TOP
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOACTIVATE = 0x0010
SWP_SHOWWINDOW = 0x0040
SW_MINIMIZE = 6
SW_RESTORE = 9
SW_HIDE = 0
GWL_EXSTYLE = -20
GWL_STYLE = -16
WS_EX_TOOLWINDOW = 0x00000080   # 不在任务栏显示
WS_EX_LAYERED = 0x00080000      # 分层窗口（支持透明键色）
WS_EX_TOPMOST = 0x00000008
WS_CAPTION = 0x00C00000         # 标题栏（用于区分普通最大化窗口与全屏）
WS_THICKFRAME = 0x00040000      # 可调整边框
LWA_COLORKEY = 0x00000001
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

# 系统托盘常量
WM_USER = 0x0400
TRAY_CALLBACK = WM_USER + 20
WM_TIMER = 0x0113
TRAY_TIMER_ID = 1
TRAY_ID = 1
WM_LBUTTONUP = 0x0202
WM_RBUTTONUP = 0x0205
WM_COMMAND = 0x0111
WM_CLOSE = 0x0010
WM_DESTROY = 0x0002
WM_NULL = 0x0000
NIM_ADD = 0
NIM_DELETE = 2
NIF_MESSAGE = 0x1
NIF_ICON = 0x2
NIF_TIP = 0x4
MF_STRING = 0x0
MF_SEPARATOR = 0x800
IDM_TOGGLE = 1
IDM_CONFIG = 2
IDM_EXIT = 3
IDI_APPLICATION = 32512

# 为 64 位下指针安全设置参数类型
if user32 is not None:
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetClassNameW.restype = ctypes.c_int
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
    kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL

KEY_COLOR = (255, 0, 255)      # 透明键色：此颜色区域将完全透明
BG = (22, 28, 44)
BORDER = (52, 64, 92)
TEXT = (235, 238, 245)
SUB = (140, 152, 178)
ACCENT = (86, 200, 255)
NOTE = (255, 214, 90)
DIM = (104, 116, 142)

PAD = 12
ROW_H = 48
HEADER_H = 82
RADIUS = 16
CIRCLE = 34
MARGIN_RIGHT = 24
INFO_W = 200            # 顶部模式左侧信息区宽度

try:
    import pygame
except ImportError:
    pygame = None


# ---------------- 字体与绘图 ----------------
_fonts = {}


def _load_font(size, bold):
    candidates = [
        "C:/Windows/Fonts/msyhbd.ttc" if bold else "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/Dengb.ttf" if bold else "C:/Windows/Fonts/Deng.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            try:
                return pygame.font.Font(p, size)
            except Exception:
                pass
    return pygame.font.SysFont("microsoftyahei,simhei,arial", size, bold=bold)


def font(size, bold=False):
    key = (size, bold)
    if key not in _fonts:
        _fonts[key] = _load_font(size, bold)
    return _fonts[key]


def hex2rgb(s):
    s = (s or "#888888").lstrip("#")
    return tuple(int(s[i:i + 2], 16) for i in (0, 2, 4))


def luminance(c):
    r, g, b = c
    return (r * 299 + g * 587 + b * 114) // 1000


def draw_text(surf, s, f, color, pos, anchor="topleft"):
    img = f.render(s, True, color)
    r = img.get_rect(**{anchor: pos})
    surf.blit(img, r)
    return r


def draw_buttons(surface, w):
    """右上角按钮；返回 {gear: Rect, close: Rect}。"""
    gear_r = pygame.Rect(w - PAD - 60, 12, 22, 22)
    close_r = pygame.Rect(w - PAD - 34, 12, 22, 22)
    pygame.draw.circle(surface, (60, 72, 98), gear_r.center, 11)
    pygame.draw.circle(surface, (60, 72, 98), close_r.center, 11)
    draw_text(surface, "…", font(14, True), TEXT, gear_r.center, anchor="center")
    draw_text(surface, "×", font(14, True), TEXT, close_r.center, anchor="center")
    return {"gear": gear_r, "close": close_r}


def course_style(cfg, name):
    subj = cfg.get("subjects", {}).get(name) if name else None
    if subj:
        return hex2rgb(subj.get("color")), (subj.get("short") or name[0])[:2], name
    return DIM, "空", name or "空"


# ---------------- 布局尺寸 ----------------

def window_size(cfg, courses, note, mode, scale, sw, sh):
    """计算窗口尺寸；mode: 'right' 竖排 / 'top' 横排；scale: 是否按屏幕缩放。"""
    win = cfg.get("window", {})
    base_w = int(win.get("width", 232))
    if mode == "top":
        if scale:
            w = int(sw * float(win.get("width_ratio", 0.75)))
        else:
            w = int(win.get("top_width", 420))
        h = int(win.get("top_height", 116))
        return w, h
    # 右侧竖排
    if scale:
        h = int(sh * float(win.get("height_ratio", 0.8)))
        return base_w, h
    n = max(len(courses), 1)
    h = HEADER_H + n * ROW_H + PAD * 2 + (26 if note else 6)
    return base_w, h


def row_height(cfg, h, courses, note, scale):
    """右侧模式每行高度（缩放时按屏高均分）。"""
    n = max(len(courses), 1)
    if scale:
        return max(34, (h - HEADER_H - PAD * 2 - (26 if note else 0)) // n)
    return ROW_H


# ---------------- 绘制：右侧竖排 ----------------

def draw_ui_right(surface, cfg, today, scale):
    w, h = surface.get_size()
    surface.fill(KEY_COLOR)
    pygame.draw.rect(surface, BG, pygame.Rect(0, 0, w, h), border_radius=RADIUS)
    pygame.draw.rect(surface, BORDER, pygame.Rect(0, 0, w, h), width=2, border_radius=RADIUS)

    date_txt = "%d年%d月%d日 %s" % (today.year, today.month, today.day,
                                    sc.WEEKDAYS[today.weekday()])
    draw_text(surface, date_txt, font(15, True), TEXT, (PAD + 4, 14))
    parity, week = sc.week_parity(today, cfg)
    if week:
        week_txt = "第%d周 · %s" % (week, "单周" if parity == "single" else "双周")
    else:
        week_txt = "强制%s" % ("单周" if parity == "single" else "双周")
    draw_text(surface, week_txt, font(13), ACCENT, (PAD + 4, 42))
    pygame.draw.line(surface, BORDER, (PAD, HEADER_H - 8), (w - PAD, HEADER_H - 8))

    buttons = draw_buttons(surface, w)

    courses, note = sc.get_day_courses(today, cfg)
    rh = row_height(cfg, h, courses, note, scale)
    y = HEADER_H + 4
    if not courses:
        draw_text(surface, "今日无课", font(18, True), DIM, (w // 2, HEADER_H + 42),
                  anchor="center")
    for i, name in enumerate(courses):
        cy = y + rh // 2
        color, short, full = course_style(cfg, name)
        cx = PAD + CIRCLE // 2 + 4
        pygame.draw.circle(surface, color, (cx, cy), CIRCLE // 2)
        icon_color = (30, 30, 40) if luminance(color) > 150 else (255, 255, 255)
        draw_text(surface, short, font(15, True), icon_color, (cx, cy), anchor="center")

        nx = cx + CIRCLE // 2 + 12
        draw_text(surface, full, font(17), TEXT if name else DIM,
                  (nx, cy - 9), anchor="midleft")
        times = cfg.get("period_times") or []
        if name and i < len(times) and times[i]:
            draw_text(surface, times[i], font(12), SUB, (nx, cy + 10), anchor="midleft")
        draw_text(surface, "第%d节" % (i + 1), font(12), SUB,
                  (w - PAD - 4, cy), anchor="midright")
        y += rh

    if note:
        draw_text(surface, "调课：" + note, font(12), NOTE,
                  (PAD + 4, h - PAD - 2), anchor="bottomleft")
    return buttons


# ---------------- 绘制：顶部横排 ----------------

def draw_ui_top(surface, cfg, today):
    w, h = surface.get_size()
    surface.fill(KEY_COLOR)
    pygame.draw.rect(surface, BG, pygame.Rect(0, 0, w, h), border_radius=RADIUS)
    pygame.draw.rect(surface, BORDER, pygame.Rect(0, 0, w, h), width=2, border_radius=RADIUS)

    buttons = draw_buttons(surface, w)

    # 左侧信息区
    date_txt = "%d年%d月%d日 %s" % (today.year, today.month, today.day,
                                    sc.WEEKDAYS[today.weekday()])
    draw_text(surface, date_txt, font(15, True), TEXT, (PAD + 4, 10))
    parity, week = sc.week_parity(today, cfg)
    if week:
        week_txt = "第%d周 · %s" % (week, "单周" if parity == "single" else "双周")
    else:
        week_txt = "强制%s" % ("单周" if parity == "single" else "双周")
    draw_text(surface, week_txt, font(13), ACCENT, (PAD + 4, 32))
    courses, note = sc.get_day_courses(today, cfg)
    if note:
        note_txt = "调课：" + note
        if len(note_txt) > 14:
            note_txt = note_txt[:14] + "…"
        draw_text(surface, note_txt, font(11), NOTE, (PAD + 4, 56))
    pygame.draw.line(surface, BORDER, (INFO_W, 8), (INFO_W, h - 8))

    # 右侧课程横向排列
    n = max(len(courses), 1)
    area_x = INFO_W + PAD
    area_w = w - area_x - PAD
    bw = area_w / n
    if not courses:
        draw_text(surface, "今日无课", font(18, True), DIM,
                  (area_x + area_w // 2, h // 2 + 4), anchor="center")
        return buttons

    for i, name in enumerate(courses):
        bx = area_x + bw * i + bw / 2
        color, short, full = course_style(cfg, name)
        cy = h // 2 + 6
        pygame.draw.circle(surface, color, (int(bx), cy), 15)
        icon_color = (30, 30, 40) if luminance(color) > 150 else (255, 255, 255)
        draw_text(surface, short, font(13, True), icon_color, (int(bx), cy),
                  anchor="center")
        # 名称（窄块时截断）
        max_chars = max(2, int((bw - 8) / 14))
        label = full if len(full) <= max_chars else full[:max_chars] + "…"
        draw_text(surface, label, font(13), TEXT if name else DIM,
                  (int(bx), cy + 20), anchor="center")
        draw_text(surface, "第%d节" % (i + 1), font(10), SUB,
                  (int(bx), cy + 36), anchor="center")
    return buttons


# ---------------- 绘制入口 ----------------

def draw_ui(surface, cfg, today):
    mode = cfg.get("window", {}).get("mode", "right")
    if mode == "top":
        return draw_ui_top(surface, cfg, today)
    return draw_ui_right(surface, cfg, today, bool(cfg.get("window", {}).get("scale_to_screen")))


# ---------------- Windows 窗口控制 ----------------

def style_window(hwnd, x, y):
    """把 pygame 窗口改为：置顶、分层透明（键色镂空）、不在任务栏显示。
    置顶与定位分离：置顶用 SetWindowPos（SWP_NOMOVE），定位统一走 _place_window
    （pygame 官方 API，与 SDL 内部窗口状态同步，避免尺寸/模式变化后位置被 SDL 覆盖）。"""
    style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    user32.SetWindowLongW(hwnd, GWL_EXSTYLE,
                          style | WS_EX_TOOLWINDOW | WS_EX_LAYERED)
    color_key = KEY_COLOR[0] | (KEY_COLOR[1] << 8) | (KEY_COLOR[2] << 16)
    user32.SetLayeredWindowAttributes(hwnd, color_key, 0, LWA_COLORKEY)
    user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
    _place_window(hwnd, x, y)


def _place_window(hwnd, x, y):
    """设置窗口在屏幕上的位置（每次窗口尺寸/位置重建后必须调用）。

    顺序很重要：
      1) 先 pygame.display.set_window_position（pygame ≥ 2.5）与 SDL 内部状态
         同步，避免后续 SDL 覆盖位置导致偏移；
      2) 再 Win32 SetWindowPos 重申置顶——实测 pygame.set_window_position
         内部用 HWND_TOP 移动窗口，会把已设置的 WS_EX_TOPMOST 置顶状态清除，
         因此必须在定位之后补一次置顶（SWP_NOMOVE 不动位置）。
    旧版 pygame 无 set_window_position 时直接用 SetWindowPos 定位 + 置顶。"""
    fn = getattr(pygame.display, "set_window_position", None)
    if fn is not None:
        try:
            fn((x, y))
        except Exception:
            pass
    else:
        user32.SetWindowPos(hwnd, HWND_TOPMOST, x, y, 0, 0,
                            SWP_NOSIZE | SWP_SHOWWINDOW)
    user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW)


def window_process_name(hwnd):
    """返回窗口所属进程的可执行文件名（小写），失败返回 ''。"""
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if not pid.value:
        return ""
    h = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value)
    if not h:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(512)
        size = wintypes.DWORD(512)
        if kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
            return os.path.basename(buf.value).lower()
        return ""
    finally:
        kernel32.CloseHandle(h)


def window_class_name(hwnd):
    """返回窗口类名（用于识别桌面、任务栏等系统窗口）。"""
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def foreground_fullscreen(hwnd):
    """仅当前台程序处于【真正的全屏】时返回 True（触发课表最小化）。

    判定条件（同时满足）：
      1. 前台窗口覆盖整个屏幕（含任务栏区域，即超出工作区）；
      2. 该窗口是无边框的（没有标题栏/可调整边框样式）——
         带标准边框的最大化窗口（如浏览器、资源管理器最大化）不算全屏；
      3. 排除系统桌面 / 任务栏 / 资源管理器窗口（点击桌面不会误触发）；
      4. 排除 Windows 任务管理器（其"最大化"不触发）。
    """
    fg = user32.GetForegroundWindow()
    if not fg or fg == hwnd:
        return False
    name = window_process_name(fg)
    if name == "taskmgr.exe":
        return False
    cls = window_class_name(fg)
    if name == "explorer.exe" or cls in ("Progman", "WorkerW", "Shell_TrayWnd",
                                         "Shell_SecondaryTrayWnd"):
        return False  # 桌面、任务栏、资源管理器
    r = wintypes.RECT()
    user32.GetWindowRect(fg, ctypes.byref(r))
    sw = user32.GetSystemMetrics(0)
    sh = user32.GetSystemMetrics(1)
    # 必须覆盖整个屏幕（含任务栏区域）——普通最大化窗口不会覆盖任务栏
    if r.left > 0 or r.top > 0 or r.right < sw or r.bottom < sh - 1:
        return False
    # 带标题栏/可调整边框 → 是普通窗口的最大化，不是全屏
    style = user32.GetWindowLongW(fg, GWL_STYLE)
    if style & (WS_CAPTION | WS_THICKFRAME):
        return False
    return True


def side_position(mode, w, h):
    """按模式计算窗口位置：top=顶部居中；right=右侧垂直居中。"""
    sw = user32.GetSystemMetrics(0)
    sh = user32.GetSystemMetrics(1)
    if mode == "top":
        return (sw - w) // 2, 12
    return sw - w - MARGIN_RIGHT, (sh - h) // 2


# ---------------- 系统托盘 ----------------

WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_long, wintypes.HWND, wintypes.UINT,
                              wintypes.WPARAM, wintypes.LPARAM)


class WNDCLASSW(ctypes.Structure):
    _fields_ = [("style", wintypes.UINT),
                ("lpfnWndProc", WNDPROC),
                ("cbClsExtra", ctypes.c_int),
                ("cbWndExtra", ctypes.c_int),
                ("hInstance", wintypes.HINSTANCE),
                ("hIcon", wintypes.HANDLE),
                ("hCursor", wintypes.HANDLE),
                ("hbrBackground", wintypes.HANDLE),
                ("lpszMenuName", wintypes.LPCWSTR),
                ("lpszClassName", wintypes.LPCWSTR)]


class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD),
                ("hWnd", wintypes.HWND),
                ("uID", wintypes.UINT),
                ("uFlags", wintypes.UINT),
                ("uCallbackMessage", wintypes.UINT),
                ("hIcon", wintypes.HANDLE),
                ("szTip", wintypes.WCHAR * 128),
                ("dwState", wintypes.DWORD),
                ("dwStateMask", wintypes.DWORD),
                ("szInfo", wintypes.WCHAR * 256),
                ("uVersion", wintypes.UINT),
                ("szInfoTitle", wintypes.WCHAR * 64),
                ("dwInfoFlags", wintypes.DWORD),
                ("guidItem", ctypes.c_byte * 16),
                ("hBalloonIcon", wintypes.HANDLE)]


def _setup_win32_argtypes():
    """集中声明 Win32 函数签名。

    关键：不声明 argtypes 时 ctypes 按 32 位 int 传参，会导致
    SetWindowPos(HWND_TOPMOST) 等调用静默失败（返回 0），
    置顶因此从未生效。必须显式声明 HWND/WPARAM/LPARAM 等 64 位类型。
    """
    if user32 is None:
        return
    user32.GetSystemMetrics.argtypes = [ctypes.c_int]
    user32.GetSystemMetrics.restype = ctypes.c_int
    user32.GetForegroundWindow.argtypes = []
    user32.GetForegroundWindow.restype = wintypes.HWND
    user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
    user32.GetWindowRect.restype = wintypes.BOOL
    user32.GetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.GetWindowLongW.restype = ctypes.c_long
    user32.SetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_long]
    user32.SetWindowLongW.restype = ctypes.c_long
    user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                                    ctypes.c_int, ctypes.c_int, wintypes.UINT]
    user32.SetWindowPos.restype = wintypes.BOOL
    user32.IsIconic.argtypes = [wintypes.HWND]
    user32.IsIconic.restype = wintypes.BOOL
    user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.ShowWindow.restype = wintypes.BOOL
    user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetClassNameW.restype = ctypes.c_int
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND,
                                                ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
    user32.FindWindowW.restype = wintypes.HWND
    kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
    kernel32.GetModuleHandleW.restype = wintypes.HMODULE
    kernel32.GetModuleFileNameW.argtypes = [wintypes.HMODULE, wintypes.LPWSTR,
                                            wintypes.DWORD]
    kernel32.GetModuleFileNameW.restype = wintypes.DWORD
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD,
                                                    wintypes.LPWSTR,
                                                    ctypes.POINTER(wintypes.DWORD)]
    kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL


class TrayIcon:
    """系统托盘图标：左键单击 = 最小化/恢复；右键菜单 = 最小化/恢复、打开配置程序、退出。
    运行在独立线程中（自带 Win32 消息循环），不干扰 pygame 主循环。"""

    def __init__(self, main_hwnd, app_dir):
        self.main_hwnd = main_hwnd
        self.app_dir = app_dir
        self._hwnd = None
        self._thread = None
        self._proc = None          # 保持 WNDPROC 回调引用，防止被 GC
        self._lock = threading.Lock()
        self.manual_minimized = False   # 用户经托盘手动最小化时置 True，避免被自动恢复

    def start(self):
        if self._thread:
            return
        self._thread = threading.Thread(target=self._run, daemon=True,
                                        name="tray-icon")
        self._thread.start()

    def stop(self):
        if self._hwnd:
            user32.PostMessageW(self._hwnd, WM_CLOSE, 0, 0)

    def is_manual_minimized(self):
        with self._lock:
            return self.manual_minimized

    def _set_manual(self, v):
        with self._lock:
            self.manual_minimized = v

    # ---------------- 托盘线程 ----------------
    def _run(self):
        # 统一声明 Win32 函数签名（64 位指针安全，避免回调/参数溢出）
        user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT,
                                          wintypes.WPARAM, wintypes.LPARAM]
        user32.DefWindowProcW.restype = ctypes.c_long
        user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT,
                                        wintypes.WPARAM, wintypes.LPARAM]
        user32.PostMessageW.restype = wintypes.BOOL
        user32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASSW)]
        user32.RegisterClassW.restype = ctypes.c_ushort
        user32.CreateWindowExW.argtypes = [
            wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
            ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, ctypes.c_void_p]
        user32.CreateWindowExW.restype = wintypes.HWND
        user32.GetCursorPos.argtypes = [ctypes.POINTER(wintypes.POINT)]
        user32.SetForegroundWindow.argtypes = [wintypes.HWND]
        user32.SetForegroundWindow.restype = wintypes.BOOL
        user32.TrackPopupMenu.argtypes = [wintypes.HMENU, wintypes.UINT,
                                          ctypes.c_int, ctypes.c_int,
                                          ctypes.c_int, wintypes.HWND, ctypes.c_void_p]
        user32.TrackPopupMenu.restype = wintypes.BOOL
        user32.CreatePopupMenu.restype = wintypes.HMENU
        user32.AppendMenuW.argtypes = [wintypes.HMENU, wintypes.UINT,
                                       wintypes.UINT, wintypes.LPCWSTR]
        user32.AppendMenuW.restype = wintypes.BOOL
        user32.DestroyMenu.argtypes = [wintypes.HMENU]
        user32.DestroyMenu.restype = wintypes.BOOL
        user32.LoadIconW.argtypes = [wintypes.HINSTANCE, ctypes.c_void_p]
        user32.LoadIconW.restype = wintypes.HANDLE
        user32.SetTimer.argtypes = [wintypes.HWND, ctypes.c_size_t,
                                    wintypes.UINT, ctypes.c_void_p]
        user32.SetTimer.restype = ctypes.c_size_t
        shell32.Shell_NotifyIconW.argtypes = [wintypes.DWORD,
                                              ctypes.POINTER(NOTIFYICONDATAW)]
        shell32.Shell_NotifyIconW.restype = wintypes.BOOL

        self._proc = WNDPROC(self._wndproc)
        wc = WNDCLASSW()
        wc.lpfnWndProc = self._proc
        wc.hInstance = kernel32.GetModuleHandleW(None)
        wc.lpszClassName = "DoubaoTrayWin"
        if not user32.RegisterClassW(ctypes.byref(wc)):
            return
        hwnd = user32.CreateWindowExW(0, "DoubaoTrayWin", "桌面课表托盘", 0,
                                      0, 0, 0, 0, None, None, wc.hInstance, None)
        if not hwnd:
            return
        self._hwnd = hwnd
        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(nid)
        nid.hWnd = hwnd
        nid.uID = TRAY_ID
        nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP
        nid.uCallbackMessage = TRAY_CALLBACK
        nid.hIcon = user32.LoadIconW(None, IDI_APPLICATION)
        nid.szTip = "桌面课表"
        if not shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid)):
            return
        # 每 2 秒触发一次置顶重申（WM_TIMER）
        user32.SetTimer(hwnd, TRAY_TIMER_ID, 2000, None)
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
        shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(nid))
        user32.PostMessageW(self.main_hwnd, WM_CLOSE, 0, 0)  # 通知主窗口退出

    # ---------------- 托盘窗口过程 ----------------
    def _wndproc(self, hwnd, msg, wp, lp):
        if msg == WM_TIMER and wp == TRAY_TIMER_ID:
            # 由托盘线程周期重申置顶（跨线程调用，可靠）
            user32.SetWindowPos(self.main_hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW)
            return 0
        if msg == TRAY_CALLBACK:
            if lp == WM_LBUTTONUP:
                self._toggle()
            elif lp == WM_RBUTTONUP:
                self._menu()
            return 0
        if msg == WM_COMMAND:
            self._command(wp & 0xFFFF)
            return 0
        if msg == WM_CLOSE:
            user32.DestroyWindow(hwnd)
            return 0
        if msg == WM_DESTROY:
            user32.PostQuitMessage(0)
            return 0
        return user32.DefWindowProcW(hwnd, msg, wp, lp)

    # ---------------- 动作 ----------------
    def _toggle(self):
        if user32.IsIconic(self.main_hwnd):
            user32.ShowWindow(self.main_hwnd, SW_RESTORE)
            self._set_manual(False)
            user32.SetWindowPos(self.main_hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW)
        else:
            user32.ShowWindow(self.main_hwnd, SW_MINIMIZE)
            self._set_manual(True)

    def _menu(self):
        menu = user32.CreatePopupMenu()
        user32.AppendMenuW(menu, MF_STRING, IDM_TOGGLE, "最小化 / 恢复")
        user32.AppendMenuW(menu, MF_STRING, IDM_CONFIG, "打开配置程序")
        user32.AppendMenuW(menu, MF_SEPARATOR, 0, None)
        user32.AppendMenuW(menu, MF_STRING, IDM_EXIT, "退出")
        pt = wintypes.POINT()
        user32.GetCursorPos(ctypes.byref(pt))
        user32.SetForegroundWindow(self._hwnd)
        user32.TrackPopupMenu(menu, 0, pt.x, pt.y, 0, self._hwnd, None)
        user32.PostMessageW(self._hwnd, WM_NULL, 0, 0)
        user32.DestroyMenu(menu)

    def _command(self, cmd):
        if cmd == IDM_TOGGLE:
            self._toggle()
        elif cmd == IDM_CONFIG:
            if getattr(sys, "frozen", False):
                subprocess.Popen([os.path.join(self.app_dir, "配置程序.exe")])
            else:
                subprocess.Popen(
                    [sys.executable, os.path.join(self.app_dir, "config_app.py")])
        elif cmd == IDM_EXIT:
            user32.DestroyWindow(self._hwnd)  # → 消息循环退出 → 删除图标 → 通知主窗口关闭


# ---------------- 主循环 ----------------

def run():
    if pygame is None:
        _fatal("缺少 pygame 库，请先运行：\npip install pygame")
        return
    _enable_dpi_awareness()

    pygame.init()
    cfg = sc.load_config()
    today = datetime.date.today()
    win = cfg.get("window", {})
    mode = win.get("mode", "right")
    scale = bool(win.get("scale_to_screen"))
    sw = user32.GetSystemMetrics(0)
    sh = user32.GetSystemMetrics(1)
    courses, note = sc.get_day_courses(today, cfg)
    w, h = window_size(cfg, courses, note, mode, scale, sw, sh)
    x, y = side_position(mode, w, h)
    os.environ["SDL_VIDEO_WINDOW_POS"] = "%d,%d" % (x, y)

    screen = pygame.display.set_mode((w, h), pygame.NOFRAME | SDL_ALWAYS_ON_TOP)
    pygame.display.set_caption("桌面课表")
    hwnd = int(pygame.display.get_wm_info()["window"])
    style_window(hwnd, x, y)

    tray = TrayIcon(hwnd, sc.APP_DIR)
    tray.start()

    clock = pygame.time.Clock()
    try:
        cfg_mtime = os.path.getmtime(sc.CONFIG_PATH)
    except OSError:
        cfg_mtime = 0
    today_s = today.strftime("%Y-%m-%d")
    buttons = {}
    poll = 0
    running = True

    while running:
        for e in pygame.event.get():
            if e.type == pygame.QUIT:
                running = False
            elif e.type == pygame.KEYDOWN and e.key == pygame.K_ESCAPE:
                running = False
            elif e.type == pygame.MOUSEBUTTONDOWN and e.button == 1:
                if buttons.get("close") and buttons["close"].collidepoint(e.pos):
                    running = False
                elif buttons.get("gear") and buttons["gear"].collidepoint(e.pos):
                    if getattr(sys, "frozen", False):
                        subprocess.Popen([os.path.join(sc.APP_DIR, "配置程序.exe")])
                    else:
                        subprocess.Popen(
                            [sys.executable, os.path.join(sc.APP_DIR, "config_app.py")])

        # 配置文件热加载
        try:
            m = os.path.getmtime(sc.CONFIG_PATH)
        except OSError:
            m = 0
        if m != cfg_mtime:
            cfg_mtime = m
            cfg = sc.load_config()

        # 跨日自动刷新
        s = datetime.date.today().strftime("%Y-%m-%d")
        if s != today_s:
            today_s = s
            today = datetime.date.today()

        # 尺寸/位置/模式变化 → 重建窗口（先更新环境变量，重建后再定位）
        win = cfg.get("window", {})
        mode = win.get("mode", "right")
        scale = bool(win.get("scale_to_screen"))
        courses, note = sc.get_day_courses(today, cfg)
        nw, nh = window_size(cfg, courses, note, mode, scale, sw, sh)
        nx, ny = side_position(mode, nw, nh)
        if (nw, nh) != screen.get_size() or (nx, ny) != (x, y):
            os.environ["SDL_VIDEO_WINDOW_POS"] = "%d,%d" % (nx, ny)
            screen = pygame.display.set_mode((nw, nh), pygame.NOFRAME)
            hwnd = int(pygame.display.get_wm_info()["window"])
            x, y, w, h = nx, ny, nw, nh
            tray.main_hwnd = hwnd   # SDL 可能重建窗口，同步托盘句柄
            style_window(hwnd, x, y)

        # 置顶保持 + 全屏检测：约每 0.4 秒一次
        # （置顶会被其他置顶窗口抢占，需周期性重申；仅真正无边框全屏触发最小化）
        poll += 1
        if poll >= 12:
            poll = 0
            user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW)
            if user32.IsIconic(hwnd):
                if not tray.is_manual_minimized() and not foreground_fullscreen(hwnd):
                    user32.ShowWindow(hwnd, SW_RESTORE)
                    user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW)
                    _place_window(hwnd, x, y)
            else:
                if foreground_fullscreen(hwnd):
                    user32.ShowWindow(hwnd, SW_MINIMIZE)
                    tray._set_manual(False)   # 全屏自动最小化，全屏退出后允许自动恢复

        buttons = draw_ui(screen, cfg, today)
        pygame.display.flip()
        clock.tick(30)

    tray.stop()
    pygame.quit()


def smoke(date_arg=None):
    """离屏渲染一帧并保存为 PNG，用于验证布局（不弹出窗口）。"""
    if pygame is None:
        print("SMOKE_FAIL missing pygame")
        return
    _enable_dpi_awareness()
    pygame.init()
    cfg = sc.load_config()
    today = datetime.date.today()
    if date_arg and date_arg != "auto":
        today = sc.parse_date(date_arg)
    mode = cfg.get("window", {}).get("mode", "right")
    scale = bool(cfg.get("window", {}).get("scale_to_screen"))
    sw = user32.GetSystemMetrics(0)
    sh = user32.GetSystemMetrics(1)
    courses, note = sc.get_day_courses(today, cfg)
    w, h = window_size(cfg, courses, note, mode, scale, sw, sh)
    surf = pygame.Surface((w, h))
    draw_ui(surf, cfg, today)
    out = os.path.join(sc.APP_DIR, "_smoke_preview.png")
    pygame.image.save(surf, out)
    parity, week = sc.week_parity(today, cfg)
    print("SMOKE_OK size=%dx%d mode=%s scale=%s courses=%d week=%s parity=%s preview=%s" %
          (w, h, mode, scale, len(courses), week, parity, out))


def _enable_dpi_awareness():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            user32.SetProcessDPIAware()
        except Exception:
            pass


def _fatal(msg):
    try:
        ctypes.windll.user32.MessageBoxW(None, msg, "桌面课表", 0x10)
    except Exception:
        print(msg)


def main():
    ap = argparse.ArgumentParser(description="桌面课表主程序")
    ap.add_argument("--smoke", nargs="?", const="auto", metavar="YYYY-MM-DD",
                    help="离屏渲染一帧并保存预览（调试用）")
    ap.add_argument("--autostart", action="store_true", help="注册开机自启后退出")
    ap.add_argument("--no-autostart", action="store_true", help="取消开机自启后退出")
    a = ap.parse_args()

    if a.autostart:
        print(sc.install_autostart())
        return
    if a.no_autostart:
        sc.remove_autostart()
        print("autostart removed")
        return
    if a.smoke is not None:
        smoke(a.smoke)
        return
    run()


_setup_win32_argtypes()

if __name__ == "__main__":
    main()
