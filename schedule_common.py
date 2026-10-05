# -*- coding: utf-8 -*-
"""
桌面课表工具 · 公共模块
负责：配置文件的读写、单双周（单双周）推算、临时调课优先级、开机自启注册。
主程序(main_app.py)与配置程序(config_app.py)共用本模块。

适用：Python 3.13（本机 3.14 亦可运行），Windows。
"""

import copy
import datetime
import json
import os
import sys

# 程序所在目录 = 配置文件的存放目录
# 说明：打包为 exe 后（sys.frozen），模块位于临时解压目录，必须改用 exe 所在目录，
#       否则配置会写到临时目录、重启后丢失。
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "schedule_data.json")

# 开机自启注册表信息
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
AUTOSTART_NAME = "DoubaoCourseReminder"

# 星期顺序（与主程序显示、配置程序列顺序保持一致）
WEEKDAYS = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]

# 16 个科目及默认配色（color 用于主程序图标圆块，short 为图标内文字）
DEFAULT_SUBJECTS = {
    "语文":   {"color": "#E74C3C", "short": "语"},
    "数学":   {"color": "#3498DB", "short": "数"},
    "英语":   {"color": "#F39C12", "short": "英"},
    "物理":   {"color": "#8E44AD", "short": "物"},
    "化学":   {"color": "#16A085", "short": "化"},
    "生物":   {"color": "#27AE60", "short": "生"},
    "政治":   {"color": "#C0392B", "short": "政"},
    "历史":   {"color": "#7F8C8D", "short": "史"},
    "地理":   {"color": "#1ABC9C", "short": "地"},
    "体育":   {"color": "#E67E22", "short": "体"},
    "音乐":   {"color": "#9B59B6", "short": "音"},
    "美术":   {"color": "#E84393", "short": "美"},
    "心理":   {"color": "#6C5CE7", "short": "心"},
    "自习":   {"color": "#95A5A6", "short": "自"},
    "体育模块": {"color": "#F1C40F", "short": "模"},
    "校本":   {"color": "#2ECC71", "short": "校"},
}

DEFAULT_CONFIG = {
    "semester_start_date": None,   # "YYYY-MM-DD"，建议为开学第一周的周一
    "week_override": None,         # None=自动推算；"single"=强制单周；"double"=强制双周
    "subjects": DEFAULT_SUBJECTS,  # {科目: {color, short}}
    "period_times": [],            # 每节课时间段（可选），如 ["08:00-08:45", ...]
    "schedule": {                  # 正常课表：单周/双周 × 星期 → [课程列表]
        "single": {d: [] for d in WEEKDAYS},
        "double": {d: [] for d in WEEKDAYS},
    },
    "temp_adjustments": {},        # 临时调课：{"YYYY-MM-DD": {"periods": [...], "note": "..."}}
    "default_periods": 10,         # 每日默认节数（配置程序空白课表也显示这么多行）
    "window": {
        "width": 232,             # 右侧模式卡片宽度
        "mode": "right",          # right=屏幕右侧竖排；top=屏幕顶部横排
        "scale_to_screen": False, # 开启后按屏幕分辨率自动缩放
        "height_ratio": 0.8,      # 右侧模式：高度占屏幕高度的比例（4/5）
        "width_ratio": 0.75,      # 顶部模式：宽度占屏幕宽度的比例（3/4）
        "top_height": 116,        # 顶部模式卡片高度
        "top_width": 420,         # 顶部模式未缩放时的固定宽度
    },
}


def load_defaults():
    """返回一份全新的默认配置（含自动推算的学期起始日）。"""
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    cfg["semester_start_date"] = _default_start().strftime("%Y-%m-%d")
    return cfg


def _default_start():
    """默认学期起始日：本周一。"""
    today = datetime.date.today()
    return today - datetime.timedelta(days=today.weekday())


def deep_merge(base, patch):
    """递归合并 dict（导入外部 JSON 时使用），列表/标量直接覆盖。"""
    for k, v in patch.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            deep_merge(base[k], v)
        else:
            base[k] = v


def load_config(path=CONFIG_PATH):
    """读取配置；文件缺失或损坏时回退为默认配置并落盘。"""
    cfg = load_defaults()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                deep_merge(cfg, data)
        except Exception:
            pass  # 配置损坏：使用默认配置
    if not cfg.get("semester_start_date"):
        cfg["semester_start_date"] = _default_start().strftime("%Y-%m-%d")
        save_config(cfg, path)
    return cfg


def save_config(cfg, path=CONFIG_PATH):
    """原子写入配置（先写临时文件再替换，避免写一半损坏）。"""
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def parse_date(s):
    """'YYYY-MM-DD' → date。格式错误会抛 ValueError。"""
    return datetime.datetime.strptime(s, "%Y-%m-%d").date()


def week_parity(date_obj, config):
    """
    计算某日的单双周。
    返回 (parity, week_number)：
      parity ∈ {'single','double'}
      week_number：自动模式下的第几周；强制单/双周时为 None。
    规则：week = (date - semester_start).days // 7 + 1；奇数周=单周，偶数周=双周。
    """
    override = config.get("week_override")
    if override in ("single", "double"):
        return override, None
    start = parse_date(config["semester_start_date"])
    delta = (date_obj - start).days
    week = 1 if delta < 0 else delta // 7 + 1
    return ("single" if week % 2 == 1 else "double"), week


def get_day_courses(date_obj, config):
    """
    取某日课程。返回 (courses列表, 备注)。
    优先级：当日临时调整 > 正常课表（按单双周 + 星期）。
    """
    key = date_obj.strftime("%Y-%m-%d")
    adj = config.get("temp_adjustments", {}).get(key)
    if adj:
        return list(adj.get("periods") or []), adj.get("note") or ""
    parity, _ = week_parity(date_obj, config)
    wk = "single" if parity == "single" else "double"
    weekday = WEEKDAYS[date_obj.weekday()]
    return list(config["schedule"].get(wk, {}).get(weekday, []) or []), ""


# ---------------- 开机自启（注册表 HKCU\\...\\Run） ----------------

def install_autostart():
    """注册开机自启，返回写入的命令行。
    exe 模式直接注册 exe；脚本模式使用 pythonw 以免弹黑窗口。"""
    import winreg
    if getattr(sys, "frozen", False):
        exe = sys.executable  # 打包后即为 桌面课表.exe
        cmd = '"%s"' % exe
    else:
        exe = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        if not os.path.exists(exe):
            exe = sys.executable
        cmd = '"%s" "%s"' % (exe, os.path.join(APP_DIR, "main_app.py"))
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
        winreg.SetValueEx(k, AUTOSTART_NAME, 0, winreg.REG_SZ, cmd)
    return cmd


def remove_autostart():
    """取消开机自启。"""
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.DeleteValue(k, AUTOSTART_NAME)
    except FileNotFoundError:
        pass


def is_autostart_installed():
    """查询开机自启是否已注册。"""
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
            winreg.QueryValueEx(k, AUTOSTART_NAME)
        return True
    except FileNotFoundError:
        return False


if __name__ == "__main__":
    # 自检：打印今天与下周一的课程信息
    cfg = load_config()
    today = datetime.date.today()
    courses, note = get_day_courses(today, cfg)
    parity, week = week_parity(today, cfg)
    print("today:", today, WEEKDAYS[today.weekday()])
    print("week :", week, parity)
    print("courses:", courses)
    print("note  :", note or "(无)")
