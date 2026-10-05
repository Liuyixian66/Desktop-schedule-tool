# -*- coding: utf-8 -*-
"""
桌面课表工具 · 主程序（Pygame）
功能：
  1. 启动后读取 schedule_data.json，按当前日期 + 单双周推算今日课程；
  2. 两种布局：屏幕右侧竖排（默认）/ 屏幕顶部横排（可在配置程序中切换）；
     开启"按屏幕分辨率自动缩放"后：右侧模式高度占屏高 4/5，顶部模式宽度占屏宽 3/4；
  3. 检测到其他程序进入全屏时自动最小化，退出全屏后自动恢复
     （Windows 任务管理器全屏/最大化时不视为全屏，不触发最小化）；
  4. 配置被修改时自动热加载（无需重启）；
  5. 卡片右上角 "…" 打开配置程序、"×" 退出，按 Esc 也可退出。

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
from ctypes import wintypes

import schedule_common as sc

# ---------------- Win32 常量与句柄 ----------------
try:
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
except AttributeError:  # 非 Windows
    user32 = None
    kernel32 = None

HWND_TOPMOST = -1
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_SHOWWINDOW = 0x0040
SW_MINIMIZE = 6
SW_RESTORE = 9
GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080   # 不在任务栏显示
WS_EX_LAYERED = 0x00080000      # 分层窗口（支持透明键色）
LWA_COLORKEY = 0x00000001
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

# 为 64 位下指针安全设置参数类型
if user32 is not None:
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
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
    优先 pygame.display.set_window_position（pygame ≥ 2.5，与 SDL 内部状态同步，
    不会被 SDL 后续覆盖）；旧版 pygame 回退 Win32 SetWindowPos。"""
    fn = getattr(pygame.display, "set_window_position", None)
    if fn is not None:
        try:
            fn((x, y))
            return
        except Exception:
            pass
    user32.SetWindowPos(hwnd, HWND_TOPMOST, x, y, 0, 0,
                        SWP_NOSIZE | SWP_SHOWWINDOW)


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


def foreground_fullscreen(hwnd):
    """判断当前前台窗口是否覆盖整个主屏幕（视为全屏程序）。
    排除：本窗口自身、Windows 任务管理器（taskmgr.exe）。"""
    fg = user32.GetForegroundWindow()
    if not fg or fg == hwnd:
        return False
    if window_process_name(fg) == "taskmgr.exe":
        return False
    r = wintypes.RECT()
    user32.GetWindowRect(fg, ctypes.byref(r))
    sw = user32.GetSystemMetrics(0)
    sh = user32.GetSystemMetrics(1)
    return (r.right - r.left) >= sw and (r.bottom - r.top) >= (sh - 2)


def side_position(mode, w, h):
    """按模式计算窗口位置：top=顶部居中；right=右侧垂直居中。"""
    sw = user32.GetSystemMetrics(0)
    sh = user32.GetSystemMetrics(1)
    if mode == "top":
        return (sw - w) // 2, 12
    return sw - w - MARGIN_RIGHT, (sh - h) // 2


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

    screen = pygame.display.set_mode((w, h), pygame.NOFRAME)
    pygame.display.set_caption("桌面课表")
    hwnd = int(pygame.display.get_wm_info()["window"])
    style_window(hwnd, x, y)

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
            style_window(hwnd, x, y)

        # 全屏检测：约每 0.4 秒一次（任务管理器不触发最小化）
        poll += 1
        if poll >= 12:
            poll = 0
            if user32.IsIconic(hwnd):
                if not foreground_fullscreen(hwnd):
                    user32.ShowWindow(hwnd, SW_RESTORE)
                    _place_window(hwnd, x, y)
            else:
                if foreground_fullscreen(hwnd):
                    user32.ShowWindow(hwnd, SW_MINIMIZE)

        buttons = draw_ui(screen, cfg, today)
        pygame.display.flip()
        clock.tick(30)

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


if __name__ == "__main__":
    main()
