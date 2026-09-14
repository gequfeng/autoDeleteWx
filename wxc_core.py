# -*- coding: utf-8 -*-
"""
企业微信外部客户清理工具 - 核心库
=====================================
整合：窗口定位、截图、SendInput 模拟点击、OpenCV 模板匹配、
      RapidOCR 时间识别、时间解析与"近半年"判断、消息列表扫描。

运行环境：Windows 10/11 + 企业微信桌面版（普通员工账号即可）
依赖：opencv-python, numpy, mss, Pillow, pywin32, psutil, openpyxl, rapidocr
"""
import ctypes
import datetime
import hashlib
import json
import os
import re
import sys
import time

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

import cv2
import numpy as np
import mss
import win32gui
import win32con
import win32process
import psutil

# ---------- 路径配置 ----------
# 打包成单文件 exe 后，数据文件（calibration.json、templates）与 exe 同目录
if getattr(sys, "frozen", False):
    BASE = os.path.dirname(os.path.abspath(sys.argv[0]))
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
TEMPLATES = os.path.join(BASE, "templates")
RECON = os.path.join(BASE, "recon")
OUTPUT = os.path.join(BASE, "outputs")
CALIB_FILE = os.path.join(BASE, "calibration.json")
CONFIG_FILE = os.path.join(BASE, "config.json")
for _d in (TEMPLATES, RECON, OUTPUT):
    os.makedirs(_d, exist_ok=True)

# ---------- 日志 ----------
def log(msg):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


# 记录已警告过的缺失文件路径，避免刷屏
_MISSING_FILE_WARNED = set()


# ---------- 窗口操作 ----------
def find_main_window():
    """查找企业微信主窗口。返回 (hwnd, rect) 或 None；检测到登录窗口时抛出异常。"""
    candidates = []

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        cls = win32gui.GetClassName(hwnd)
        title = win32gui.GetWindowText(hwnd)
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        w, h = r - l, b - t
        if w < 600 or h < 400:
            return
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        try:
            proc = psutil.Process(pid).name().lower()
        except Exception:
            proc = ""
        score = 0
        if cls == "WeWorkWindow" and title == "企业微信":
            score += 1000
        if proc == "wxwork.exe":
            score += 100
        if score > 0:
            candidates.append((score, hwnd, (l, t, r, b), cls, title, proc))

    win32gui.EnumWindows(cb, None)

    if not candidates:
        # 检测登录窗口，给出更明确的提示
        def login_cb(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            cls = win32gui.GetClassName(hwnd)
            title = win32gui.GetWindowText(hwnd)
            if cls == "WeChatLogin" and title == "企业微信":
                candidates.append((-1, hwnd, None, cls, title, ""))
        win32gui.EnumWindows(login_cb, None)
        if candidates:
            raise RuntimeError("检测到企业微信【登录窗口】，请先在电脑上完成登录并最大化主窗口。")
        return None

    candidates.sort(reverse=True)
    _, hwnd, rect, cls, title, proc = candidates[0]
    log(f"找到主窗口: {title} | 类: {cls} | 位置: {rect}")
    return hwnd, rect


def activate(hwnd, maximize=True):
    """
    激活并前置窗口，返回窗口最新 rect。
    关键：最小化窗口的 GetWindowRect 返回 (-32000, -32000)，
    必须先恢复/最大化再重新取坐标，否则截图全空。
    """
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.3)
        if maximize:
            # 强制最大化，保证消息列表区域与校准截图一致
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.5)
        rect = win32gui.GetWindowRect(hwnd)
        log(f"窗口已激活: {rect}")
        return rect
    except Exception as e:
        log(f"[warn] 激活窗口失败: {e}")
        try:
            return win32gui.GetWindowRect(hwnd)
        except Exception:
            return None


def ensure_visible(hwnd, rect):
    """
    统一入口：确保窗口可见、最大化，返回最新 rect。
    供扫描/删除流程在调用 find_main_window() 后使用。
    """
    new_rect = activate(hwnd)
    return new_rect if new_rect else rect


# ---------- 截图 ----------
def window_shot(rect):
    """截取窗口指定区域。rect = (left, top, right, bottom)。"""
    l, t, r, b = rect
    # mss 8.x 用 mss.mss() 工厂函数；9.x 改名 MSS() 又改回 mss.mss()。
    # 跨版本最稳的是先试 mss.mss()，失败再 try MSS()。
    mss_factory = getattr(mss, "mss", None) or getattr(mss, "MSS", None)
    if mss_factory is None:
        raise RuntimeError(f"当前 mss 版本不支持已知 API: {mss.__version__}")
    with mss_factory() as s:
        img = np.array(s.grab({"left": l, "top": t, "width": r - l, "height": b - t}))
    if img.shape[2] == 4:
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
    return img


def save_shot(img, name, subdir=RECON):
    path = os.path.join(subdir, name)
    cv2_imwrite(path, img)
    return path


def cv2_imread(path, flags=cv2.IMREAD_COLOR):
    """支持中文路径的 cv2 读取；缺失文件只警告一次。"""
    if not path:
        return None
    if not os.path.exists(path):
        if path not in _MISSING_FILE_WARNED:
            _MISSING_FILE_WARNED.add(path)
            log(f"[warn] 读取图片失败 {path}: 文件不存在")
        return None
    try:
        data = np.fromfile(path, dtype=np.uint8)
        if data.size == 0:
            return None
        return cv2.imdecode(data, flags)
    except Exception as e:
        log(f"[warn] 读取图片失败 {path}: {e}")
        return None


def cv2_imwrite(path, img, params=None):
    """支持中文路径的 cv2 保存（解决 OpenCV 在 Windows 非 ASCII 路径上静默失败）。"""
    try:
        ext = os.path.splitext(path)[1]
        if not ext:
            ext = ".png"
        d = os.path.dirname(path)
        if d:
            os.makedirs(d, exist_ok=True)
        ok, buf = cv2.imencode(ext, img, params or [])
        if ok:
            buf.tofile(path)
            return True
        log(f"[warn] 编码图片失败: {path}")
        return False
    except Exception as e:
        log(f"[warn] 保存图片失败 {path}: {e}")
        return False


# ---------- 模板匹配 ----------
def find_template(screen_bgr, tpl_bgr, threshold=0.6):
    """在 screen 中匹配模板，返回 (中心点(x,y), 相似度) 或 (None, score)。"""
    if tpl_bgr is None or screen_bgr is None:
        return None, 0.0
    h, w = tpl_bgr.shape[:2]
    if screen_bgr.shape[0] < h or screen_bgr.shape[1] < w:
        return None, 0.0
    res = cv2.matchTemplate(screen_bgr, tpl_bgr, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(res)
    if max_val < threshold:
        return None, max_val
    return (max_loc[0] + w // 2, max_loc[1] + h // 2), max_val


def find_all_templates(screen_bgr, tpl_bgr, threshold=0.6, max_boxes=50):
    """找所有匹配位置（去重），返回 [(中心点, 相似度), ...]。"""
    if tpl_bgr is None or screen_bgr is None:
        return []
    th, tw = tpl_bgr.shape[:2]
    if screen_bgr.shape[0] < th or screen_bgr.shape[1] < tw:
        return []
    res = cv2.matchTemplate(screen_bgr, tpl_bgr, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= threshold)
    pts = list(zip(*loc[::-1]))  # (x, y)
    if not pts:
        return []
    # 按 y 排序，合并重叠区域
    boxes = []
    for x, y in pts:
        boxes.append((x + tw // 2, y + th // 2))
    # 简单去重：按距离
    merged = []
    for cx, cy in sorted(boxes, key=lambda p: (p[1], p[0])):
        dup = False
        for mx, my in merged:
            if abs(cx - mx) < tw and abs(cy - my) < th:
                dup = True
                break
        if not dup:
            merged.append((cx, cy))
        if len(merged) >= max_boxes:
            break
    return merged


# ---------- 点击/滚轮操作 ----------
class _MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long), ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong), ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong), ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))
    ]


class _INPUT_I(ctypes.Union):
    _fields_ = [("mi", _MOUSEINPUT)]


class _INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong), ("ii", _INPUT_I)]


def _send_mouse(flags, x, y, mouse_data=0):
    """底层 SendInput：绝对坐标移动/按键/滚轮。"""
    sw = ctypes.windll.user32.GetSystemMetrics(0)
    sh = ctypes.windll.user32.GetSystemMetrics(1)
    nx = int(x * 65535 / sw)
    ny = int(y * 65535 / sh)
    inp = _INPUT(type=0)
    inp.ii.mi.dx = nx
    inp.ii.mi.dy = ny
    inp.ii.mi.mouseData = mouse_data & 0xFFFFFFFF
    inp.ii.mi.dwFlags = 0x8000 | flags
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(_INPUT))


def send_click(x, y, double=False):
    """使用 SendInput 在屏幕绝对坐标 (x,y) 移动并点击左键。
    必须 MOVE + DOWN + UP 三步，Qt 自绘窗口才会响应。"""
    for flag in (0x0001, 0x0002, 0x0004):  # MOVE, LEFT DOWN, LEFT UP
        _send_mouse(flag, x, y)
        time.sleep(0.05)
    if double:
        for flag in (0x0002, 0x0004):
            _send_mouse(flag, x, y)
            time.sleep(0.05)


def send_right_click(x, y):
    """使用 SendInput 在屏幕绝对坐标 (x,y) 右键单击。"""
    for flag in (0x0001, 0x0008, 0x0010):  # MOVE, RIGHT DOWN, RIGHT UP
        _send_mouse(flag, x, y)
        time.sleep(0.05)


def scroll_at(x, y, notches=3, up=False):
    """在屏幕绝对坐标 (x,y) 处滚动鼠标滚轮。notches 为格数，默认向下。"""
    _send_mouse(0x0001, x, y)  # 先把光标移到目标位置
    time.sleep(0.05)
    n = max(1, abs(int(notches)))
    delta = (120 if up else -120) * n  # MOUSEEVENTF_WHEEL，负值为向下
    _send_mouse(0x0800, x, y, mouse_data=delta)
    time.sleep(0.05)


def click_rel(rect, rel_x, rel_y, double=False):
    """点击窗口内相对坐标（rect 左上角为原点）。"""
    send_click(rect[0] + rel_x, rect[1] + rel_y, double=double)


# ---------- 外部客户识别 ----------
def _load_template(tpl):
    if isinstance(tpl, str):
        return cv2_imread(tpl)
    return tpl


def _item_has_blue_bg(item_img, threshold=0.25):
    """检测条目是否有蓝色高亮背景（当前选中项）。"""
    hsv = cv2.cvtColor(item_img, cv2.COLOR_BGR2HSV)
    # 企业微信选中背景偏蓝
    lower_blue = np.array([90, 20, 120])
    upper_blue = np.array([130, 255, 255])
    mask = cv2.inRange(hsv, lower_blue, upper_blue)
    ratio = np.sum(mask > 0) / (mask.size + 1e-6)
    return ratio > threshold


def _preprocess_for_blue_ocr(img_bgr):
    """蓝色背景白色文字：增强对比，转成黑字白底，提升 OCR 准确率。"""
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    # 反色：白字 -> 黑字
    inv = cv2.bitwise_not(gray)
    # 自适应阈值增强
    _, binary = cv2.threshold(inv, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    # 膨胀让细字变粗
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    binary = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)
    # 转回 3 通道
    return cv2.cvtColor(binary, cv2.COLOR_GRAY2BGR)


def _item_text_contains_at_wechat(item_img):
    """对条目做 OCR，检查是否含 '@微信' 字样（支持蓝色选中项）。"""
    h, w = item_img.shape[:2]
    # 取全宽（外部联系人列表没有时间在最右侧；消息列表里时间也不含 @微信）
    region = item_img
    # 普通 OCR
    results = ocr_text(region)
    text = "".join([t for t, _c, _b in results])
    if "@微信" in text:
        return True
    # 蓝色背景增强后 OCR
    if _item_has_blue_bg(item_img):
        enhanced = _preprocess_for_blue_ocr(region)
        results2 = ocr_text(enhanced)
        text2 = "".join([t for t, _c, _b in results2])
        if "@微信" in text2:
            return True
    return False


def is_external_contact(item_img, at_templates=None):
    """
    判断会话条目是否为外部客户（带 @微信 标签）。
    at_templates: 从校准得到的标签模板列表（图像或路径）；为空时用默认模板 + 绿色检测兜底。
    额外处理：当前选中项蓝色背景上的白色 @微信 标签。
    """
    # 1. 模板匹配（绿色 @微信 标签）
    if at_templates:
        for tpl in at_templates:
            tpl = _load_template(tpl)
            if tpl is None or tpl.shape[0] > item_img.shape[0] or tpl.shape[1] > item_img.shape[1]:
                continue
            res = cv2.matchTemplate(item_img, tpl, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, _ = cv2.minMaxLoc(res)
            if max_val >= 0.60:
                return True
    else:
        for tpl_name in ("at_wechat.png", "at_wechat_highlight.png"):
            tpl = cv2_imread(os.path.join(TEMPLATES, tpl_name))
            if tpl is None or tpl.shape[0] > item_img.shape[0] or tpl.shape[1] > item_img.shape[1]:
                continue
            res = cv2.matchTemplate(item_img, tpl, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, _ = cv2.minMaxLoc(res)
            if max_val >= 0.60:
                return True

    # 2. 兜底：绿色标签检测（未选中项）
    hsv = cv2.cvtColor(item_img, cv2.COLOR_BGR2HSV)
    lower_green = np.array([35, 40, 40])
    upper_green = np.array([85, 255, 255])
    green_mask = cv2.inRange(hsv, lower_green, upper_green)
    green_ratio = np.sum(green_mask > 0) / (green_mask.size + 1e-6)
    if green_ratio > 0.015:
        return True

    # 3. 蓝色选中背景项：OCR 识别 @微信
    if _item_has_blue_bg(item_img):
        if _item_text_contains_at_wechat(item_img):
            return True

    return False


def find_at_tag_positions(screen_bgr, at_templates=None, threshold=0.6):
    """在全屏/区域图中找出所有 @微信 标签的中心点。返回 [(x, y), ...]"""
    if at_templates:
        all_pts = []
        for tpl in at_templates:
            if tpl is None:
                continue
            all_pts.extend(find_all_templates(screen_bgr, tpl, threshold=threshold))
        # 去重
        seen, uniq = set(), []
        for p in all_pts:
            key = (p[0] // 10, p[1] // 10)
            if key not in seen:
                seen.add(key)
                uniq.append(p)
        if uniq:
            return uniq
    # 兜底：绿色横向矩形区域质心（@微信 标签是横向小矩形，宽高比 >= 1.5）
    hsv = cv2.cvtColor(screen_bgr, cv2.COLOR_BGR2HSV)
    lower_green = np.array([35, 40, 40])
    upper_green = np.array([85, 255, 255])
    mask = cv2.inRange(hsv, lower_green, upper_green)
    # 形态学闭运算，让断开的标签连成块
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    pts = []
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        # @微信 标签：横向 25~90，高 12~30，宽高比 >= 1.5
        if w < 20 or h < 10 or w > 100 or h > 32:
            continue
        if h > 0 and w / h < 1.5:
            continue
        pts.append((x + w // 2, y + h // 2))
    return sorted(pts, key=lambda p: (p[1], p[0]))


def _find_at_tag_positions_by_ocr(screen_bgr):
    """
    用 OCR 直接定位 "@微信" 文本的中心点。比模板匹配更稳健，避免模板污染。
    限定在中间面板区域 (x 200~700, y >= 100) 内识别，返回 [(x, y), ...]。
    文本 "@微信" 容错：OCR 经常识别为 "国微信" "@ 微信" "@微信" "口微信" 等。
    同时也支持 "HY@微信" "名字@微信" 这种名字+标签连写的情况。
    """
    h, w = screen_bgr.shape[:2]
    # 只对中间面板做 OCR，加快速度
    x1, x2 = max(0, 200), min(w, 720)
    region = screen_bgr[100:, x1:x2]
    try:
        results = ocr_text(region)
    except Exception as e:
        log(f"[warn] OCR 失败: {e}")
        return []
    pts = []
    for txt, conf, box in results:
        if not box:
            continue
        t = txt.replace(" ", "")
        # 必须包含 "微信"
        if "微信" not in t:
            continue
        # 过滤明显不是标签的文本
        if len(t) > 15:
            continue
        bx, by, bw, bh = box
        cx = bx + bw // 2 + x1
        cy = by + bh // 2 + 100
        # 文本框高度合理性：10~30 (单行文字)
        if not (10 <= bh <= 32):
            continue
        # 过滤掉分组标题和其他误识别
        noise_texts = ("微信联系人", "外部联系人", "测试测试", "新的联系人",
                       "清晰标记你的外部联系人", "添加联系人", "添加")
        if any(n in t for n in noise_texts):
            continue
        # 必须以 @/国/口/同 等可识别符号开头（避免误识"微信联系人"等）
        # 或整个文本就是 "@微信"
        if not any(c in t for c in "@国口同") and t != "@微信":
            continue
        pts.append((cx, cy))
    # 按 y 排序
    pts.sort(key=lambda p: (p[1], p[0]))
    # 去重：相同 y 区间（±20）保留一个
    uniq = []
    for p in pts:
        if not any(abs(p[1] - q[1]) < 20 and abs(p[0] - q[0]) < 30 for q in uniq):
            uniq.append(p)
    return uniq


# ---------- OCR 时间识别 ----------
_OCR_ENGINE = None


def get_ocr():
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        try:
            from rapidocr import RapidOCR
            _OCR_ENGINE = RapidOCR()
        except Exception as e:
            log(f"[warn] RapidOCR 初始化失败: {e}")
            _OCR_ENGINE = False
    return _OCR_ENGINE if _OCR_ENGINE else None


def ocr_text(img_bgr):
    """对图像做 OCR，返回 [(文本, 置信度, (x,y,w,h)), ...]"""
    engine = get_ocr()
    if engine is None:
        return []
    try:
        out = engine(img_bgr)
        txts = list(out.txts) if out.txts else []
        scores = list(out.scores) if out.scores else []
        boxes = out.boxes if out.boxes is not None else []
        results = []
        for i, t in enumerate(txts):
            conf = scores[i] if i < len(scores) else 0.0
            box = boxes[i] if i < len(boxes) else None
            x, y, w, h = 0, 0, 0, 0
            if box is not None and len(box) == 4:
                xs = [p[0] for p in box]
                ys = [p[1] for p in box]
                x, y = int(min(xs)), int(min(ys))
                w, h = int(max(xs) - min(xs)), int(max(ys) - min(ys))
            results.append((t, float(conf), (x, y, w, h)))
        return results
    except Exception as e:
        log(f"[warn] OCR 识别失败: {e}")
        return []


# ---------- 时间解析与半年判断 ----------
WEEKDAYS = {"星期一": 0, "周二": 1, "星期二": 1, "周三": 2, "星期三": 2, "周四": 3,
            "星期四": 3, "周五": 4, "星期五": 4, "周六": 5, "星期六": 5, "周日": 6,
            "星期天": 6, "周日": 6}


def parse_last_contact(text, today=None):
    """
    解析企业微信消息列表的时间文本，返回 datetime.date 或 None。
    支持格式：
      HH:MM / H:MM            -> 今天
      昨天                     -> 昨天
      星期X / 周X              -> 本周内最近的一天
      X月X日                  -> 今年（若晚于今天则视为去年）
      X/X                    -> 今年（若晚于今天则视为去年）
      YYYY/X/X  / YYYY年X月X日 -> 直接解析
    """
    if today is None:
        today = datetime.date.today()
    t = (text or "").strip()
    if not t:
        return None

    # 1. HH:MM
    m = re.match(r"^(\d{1,2}):(\d{2})$", t)
    if m:
        return today

    # 2. 昨天
    if t == "昨天":
        return today - datetime.timedelta(days=1)

    # 3. 星期X / 周X
    for wd, idx in WEEKDAYS.items():
        if t == wd or t.endswith(wd):
            diff = (today.weekday() - idx) % 7
            if diff == 0:
                diff = 7  # 如果显示的正好是今天，但一般今天会显示时间，保守取 7 天前
            return today - datetime.timedelta(days=diff)

    # 4. X月X日
    m = re.match(r"^(\d{1,2})月(\d{1,2})日?$", t)
    if m:
        month, day = int(m.group(1)), int(m.group(2))
        try:
            d = datetime.date(today.year, month, day)
            if d > today:  # 未来日期 -> 去年
                d = datetime.date(today.year - 1, month, day)
            return d
        except ValueError:
            return None

    # 5. X/X
    m = re.match(r"^(\d{1,2})/(\d{1,2})$", t)
    if m:
        month, day = int(m.group(1)), int(m.group(2))
        try:
            d = datetime.date(today.year, month, day)
            if d > today:
                d = datetime.date(today.year - 1, month, day)
            return d
        except ValueError:
            return None

    # 6. YYYY/X/X
    m = re.match(r"^(\d{4})[/年.](\d{1,2})[/月.](\d{1,2})日?$", t)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None

    # 7. YYYY-MM-DD
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", t)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None

    return None


def months_before(today, n_months):
    """返回 n 个月前的日期（同日，跨月时取当月最后一天）。"""
    total = today.year * 12 + (today.month - 1) - n_months
    y, m = divmod(total, 12)
    m += 1
    import calendar
    last_day = calendar.monthrange(y, m)[1]
    return datetime.date(y, m, min(today.day, last_day))


# 清理基准日期：最后联系时间早于此日期才算需要清理（不再按半年动态计算）
CUTOFF_DATE = datetime.date(2026, 1, 1)
CUTOFF_LABEL = "2026.1.1前未联系"


def is_over_half_year(last_date, today=None):
    """最后联系时间是否早于 2026-01-01（固定基准日期，today 参数仅为兼容保留）。"""
    if last_date is None:
        return None  # 无法判断
    return last_date < CUTOFF_DATE


# ---------- 名字/时间提取 ----------
_TIME_RE = re.compile(
    r"(\d{1,2}[:：]\d{2}|昨天|星期[一二三四五六天日]|周[一二三四五六天日]|"
    r"\d{1,2}月\d{1,2}日?|\d{4}[/年.]\d{1,2}[/月.]\d{1,2}日?|\d{1,2}/\d{1,2})"
)


def _strip_tag(t):
    """去掉外部客户标签，保留名字。"""
    t = t.replace(" ", "")
    # 去掉 @微信 / 微信 标签（通常出现在末尾）
    if "@微信" in t:
        t = t.split("@微信")[0]
    elif t.endswith("微信") and len(t) > 2:
        t = t[:-2]
    t = t.rstrip("@")
    return t.strip()


def _ocr_variants_for_time(crop_gray):
    """为时间区域生成多个 OCR 预处理变体。"""
    inv = cv2.bitwise_not(crop_gray)
    _, otsu = cv2.threshold(inv, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    out = [
        cv2.cvtColor(crop_gray, cv2.COLOR_GRAY2BGR),
        cv2.cvtColor(inv, cv2.COLOR_GRAY2BGR),
        cv2.cvtColor(otsu, cv2.COLOR_GRAY2BGR),
    ]
    # 2x 放大对细小时间文字更有效
    out.append(cv2.resize(out[1], None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC))
    out.append(cv2.resize(out[2], None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC))
    return out


def extract_customer_info(item_img, is_blue=None):
    """
    从单条会话条目中提取 (名字, 时间文本)。
    策略：整条 OCR + 文字框位置分类；时间单独处理右下角小区域。
    """
    h, w = item_img.shape[:2]
    if is_blue is None:
        is_blue = _item_has_blue_bg(item_img)

    # ---- 时间：取右上角区域做 OCR ----
    # 企业微信消息列表中，时间与名字在同一行，位于条目右上角
    time_text = ""
    tx = int(w * 0.55)
    ty_end = int(h * 0.45)
    crop = item_img[:ty_end, tx:]
    if crop.size:
        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
        for v in _ocr_variants_for_time(gray):
            for t, _c, _b in ocr_text(v):
                t2 = t.replace(" ", "").replace("：", ":")
                m = _TIME_RE.search(t2)
                if m:
                    time_text = m.group(0)
                    break
            if time_text:
                break

    # ---- 名字：整条 OCR，按位置分类 ----
    name = ""
    sources = [item_img]
    if is_blue:
        # 蓝色背景白字：反色+otsu 有助于把名字和标签拆开
        gray = cv2.cvtColor(item_img, cv2.COLOR_BGR2GRAY)
        inv = cv2.bitwise_not(gray)
        _, otsu = cv2.threshold(inv, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        sources.append(
            cv2.resize(cv2.cvtColor(otsu, cv2.COLOR_GRAY2BGR), None,
                       fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        )

    all_texts = []
    for src in sources:
        for t, conf, box in ocr_text(src):
            x, y, bw, bh = box
            all_texts.append((y, y + bh / 2, x, t.strip(), float(conf)))

    # 排序：y 小的在前
    all_texts.sort(key=lambda r: r[0])

    # 策略 1：文本里直接包含 @微信，去掉标签即名字
    for y, cy, x, t, conf in all_texts:
        t2 = t.replace(" ", "")
        if "@微信" in t2 or ("微信" in t2 and len(t2) <= 10):
            part = _strip_tag(t2)
            if part:
                name = part
                break

    # 策略 2：在名字行（上半部）找最靠上的非时间、非标签文本
    if not name:
        for y, cy, x, t, conf in all_texts:
            t2 = t.replace(" ", "")
            # 过滤顶部边界噪声（低置信/极靠上）
            if y < 8 or conf < 0.65:
                continue
            # 只取上半部（名字行），下半部通常是消息预览/时间
            if y > h * 0.55:
                continue
            # 跳过时间
            if _TIME_RE.fullmatch(t2):
                continue
            # 跳过孤立标签/符号/未读角标
            if t2 in ("@微信", "微信", "@", "GETF") or (t2.isdigit() and len(t2) <= 2):
                continue
            part = _strip_tag(t2)
            if part:
                name = part
                break

    return name, time_text


# ---------- 外部联系人导航 ----------
def _get_calib_template(calib, img_key, path_key):
    """
    从 calib 中取模板图像：优先取已加载的 *_img 键；
    没有则按 path_key 的路径加载（相对路径按 BASE 解析）并缓存回 calib。
    解决 calibrate() 返回的字典只有路径、没有图像的问题。
    """
    if not calib:
        return None
    img = calib.get(img_key)
    if img is not None:
        return img
    path = calib.get(path_key)
    if path:
        if not os.path.isabs(path):
            path = os.path.join(BASE, path)
        img = cv2_imread(path)
        if img is not None:
            calib[img_key] = img
    return img


def _get_calib_templates_list(calib):
    """从 calib 中取 @微信 标签模板列表（同样兼容路径形式）。"""
    if not calib:
        return None
    imgs = calib.get("at_template_imgs")
    if imgs:
        return imgs
    paths = calib.get("at_templates")
    if paths:
        imgs = []
        for p in paths:
            if not p:
                continue
            fp = p if os.path.isabs(p) else os.path.join(BASE, p)
            im = cv2_imread(fp)
            if im is not None:
                imgs.append(im)
        if imgs:
            calib["at_template_imgs"] = imgs
        return imgs if imgs else None
    return None


def find_green_wechat_icons(screen_bgr, wechat_icon_template=None, threshold=0.6):
    """
    在整张截图中查找绿色微信图标位置。
    优先用模板匹配；没有模板时用颜色+形状检测兜底。
    返回 [(cx, cy), ...]，按 y 排序。
    """
    pts = []
    if wechat_icon_template is not None:
        pts = find_all_templates(screen_bgr, wechat_icon_template, threshold=threshold)
    if pts:
        return sorted(pts, key=lambda p: p[1])

    # 兜底：颜色+形状检测
    hsv = cv2.cvtColor(screen_bgr, cv2.COLOR_BGR2HSV)
    lower_green = np.array([35, 40, 40])
    upper_green = np.array([85, 255, 255])
    mask = cv2.inRange(hsv, lower_green, upper_green)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        if 14 <= w <= 34 and 14 <= h <= 34:
            ratio = max(w, h) / (min(w, h) + 1e-6)
            if ratio <= 1.6:
                pts.append((x + w // 2, y + h // 2))
    return sorted(pts, key=lambda p: p[1])


def click_at(hwnd, rect, x, y):
    """在窗口绝对坐标 (x,y) 处点击（x,y 相对于屏幕）。"""
    send_click(x, y)


def navigate_to_external_contacts(hwnd, rect, calib=None):
    """
    导航到通讯录页面。
    流程：点击左侧导航栏通讯录图标，等待页面加载完成。
    注意：外部联系人入口在中间面板的"微信联系人"文件夹，不在左侧，交由 expand_wechat_contacts_folder 处理。
    """
    full = window_shot(rect)
    save_shot(full, "nav_step0_full.png")

    # 统一使用受限的导航栏点击，避免全屏低阈值匹配误点微盘
    if not click_contacts_nav(hwnd, rect, calib):
        log("[warn] 未匹配到通讯录图标模板，尝试点击默认位置")
        click_rel(rect, 60, 250)
        time.sleep(1.0)
    else:
        time.sleep(1.0)

    rect = ensure_visible(hwnd, rect)
    return rect


def _folder_region_has_contacts(screen_bgr, at_templates, folder_y, folder_x):
    """
    判断"微信联系人"文件夹是否已经展开：
    在文件夹下方中间面板搜索 @微信 标签模板或绿色标签像素。
    """
    h, w = screen_bgr.shape[:2]
    x1, x2 = max(0, folder_x - 120), min(w, folder_x + 160)
    y1 = min(h - 50, folder_y + 30)
    y2 = h - 10
    if y1 >= y2 or x1 >= x2:
        return False
    region = screen_bgr[y1:y2, x1:x2]
    # 1. 模板匹配
    if at_templates:
        for tpl in at_templates:
            if tpl is None:
                continue
            pts = find_all_templates(region, tpl, threshold=0.6)
            if pts:
                return True
    # 2. 绿色像素兜底
    hsv = cv2.cvtColor(region, cv2.COLOR_BGR2HSV)
    mask = cv2.inRange(hsv, np.array([35, 40, 40]), np.array([85, 255, 255]))
    ratio = np.sum(mask > 0) / mask.size
    return ratio > 0.01


def expand_wechat_contacts_folder(hwnd, rect, calib=None):
    """
    在通讯录页面，点击展开中间面板的"微信联系人"分组。
    返回最新的 rect。
    """
    rect = ensure_visible(hwnd, rect)
    full = window_shot(rect)
    save_shot(full, "nav_step2_external.png")
    wechat_tpl = _get_calib_template(calib, "wechat_icon_template_img", "wechat_icon_template")
    at_templates = _get_calib_templates_list(calib)
    pts = find_green_wechat_icons(full, wechat_icon_template=wechat_tpl, threshold=0.5)
    h, w = full.shape[:2]
    # 中间分组面板里的文件夹图标：x 在 250~550，y 在 80~350（考虑上方可能有其他分组）
    folder_pts = [p for p in pts if 250 <= p[0] <= min(550, w - 20) and 80 <= p[1] <= min(350, h - 100)]
    if folder_pts:
        # 取最靠上的文件夹图标
        target = sorted(folder_pts, key=lambda p: p[1])[0]
        # 如果已经展开（下方能看到 @微信 标签），不要再点，避免折叠
        if _folder_region_has_contacts(full, at_templates, target[1], target[0]):
            log(f"'微信联系人'分组已展开: {target}")
        else:
            log(f"点击展开'微信联系人'分组: {target}")
            click_at(hwnd, rect, target[0], target[1])
            time.sleep(0.8)
    else:
        log("[warn] 未找到'微信联系人'分组图标，假设已展开")
    return ensure_visible(hwnd, rect)


def extract_external_contact_name(item_img):
    """
    从外部联系人列表条目中提取名字。
    布局：头像 | 名字 | @微信 标签（可能同一行或换行）
    返回名字字符串。
    """
    h, w = item_img.shape[:2]
    # 外部联系人列表条目不宽，直接全宽 OCR
    region = item_img
    results = ocr_text(region)
    if not results:
        return ""

    # 按 y 排序
    results.sort(key=lambda r: r[2][1])

    # 优先找包含 @微信 的文本，去掉标签即名字
    for t, conf, box in results:
        t2 = t.replace(" ", "").strip()
        if "@微信" in t2:
            name = _strip_tag(t2)
            if name:
                return name

    # 否则取最上方的非标签、非"共X位客户"、非数字短串文本
    for t, conf, box in results:
        t2 = t.replace(" ", "").strip()
        if not t2:
            continue
        if t2 in ("@微信", "微信", "@"):
            continue
        if "位客户" in t2 or "客户" in t2:
            continue
        if t2.isdigit():
            continue
        # 过滤纯 emoji / 特殊符号
        if len(t2) <= 1 and not any("\u4e00" <= c <= "\u9fff" for c in t2):
            continue
        # 可能名字和 @微信 分开，返回名字
        if "微信" not in t2:
            return t2
    return ""


# ---------- 外部联系人列表扫描 ----------
def scan_external_contacts_list(hwnd, rect, calib=None):
    """
    扫描通讯录 -> 外部联系人 -> 微信联系人 列表。
    返回外部客户记录列表（名字、时间字段为空，因外部联系人列表不显示时间）。
    """
    rect = navigate_to_external_contacts(hwnd, rect, calib)
    rect = expand_wechat_contacts_folder(hwnd, rect, calib)

    l, t, r, b = rect
    full = window_shot(rect)
    save_shot(full, "scan_ext_full.png")

    # 外部联系人列表区域：中间面板
    win_h = b - t
    if calib and calib.get("list_region"):
        lx, ly, rx, _ = calib["list_region"]
        # 如果校准的是消息列表区域（x<160），改用外部联系人默认区域
        if lx < 150:
            lx, ly, rx = 250, 170, 540
    else:
        lx, ly, rx = 250, 170, 540
    ry = win_h - 10
    list_img = full[ly:ry, lx:rx]
    save_shot(list_img, "scan_ext_list_region.png")
    log(f"外部联系人列表区域: ({lx},{ly})-({rx},{ry})")

    # 外部联系人条目高度约 50~60px
    item_height = calib.get("item_height") if calib else None
    if item_height and item_height > 70:
        item_height = 55
    items = split_sessions(list_img, item_height=item_height, min_height=45)
    log(f"分割出 {len(items)} 个候选条目")

    at_templates = _get_calib_templates_list(calib)

    records = []
    skipped_system = 0
    for idx, (y1, y2) in enumerate(items):
        item = list_img[y1:y2, :]
        if not is_external_contact(item, at_templates=at_templates):
            continue
        name = extract_external_contact_name(item)
        # 过滤系统账号、空名、以及明显不是人名的文件名/菜单噪声
        if not name:
            continue
        clean = name.replace(" ", "")
        if "企业微信团队" in clean:
            log(f"跳过系统账号 #{idx}: {name}")
            skipped_system += 1
            continue
        # 过滤常见误识别：文件扩展名、快捷方式、菜单项等
        noise_terms = (".zip", ".exe", "快捷方式", "Jenkins", "预定会议", "加入会议", "添加日历")
        if any(t in clean for t in noise_terms):
            log(f"跳过噪声条目 #{idx}: {name}")
            continue

        records.append({
            "name": name,
            "time_text": "",
            "last_date": "",
            "over_half_year": None,
            "note": "外部客户（通讯录-外部联系人）",
        })
        log(f"外部客户 #{idx}: 名字={name!r}")

    log(f"扫描完成，有效外部客户 {len(records)} 人，跳过系统账号/噪声 {skipped_system} 条")
    return records


def scan_current_external_contacts_page(hwnd, rect, calib=None):
    """
    直接扫描当前已打开的外部联系人页面（需客户手动定位到该页面）。
    不再自动点击通讯录或展开文件夹，而是根据 @微信 标签位置直接定位每个联系人条目。
    返回外部客户记录列表。
    """
    rect = ensure_visible(hwnd, rect)
    full = window_shot(rect)
    save_shot(full, "scan_current_full.png")

    h, w = full.shape[:2]
    # 优先用 OCR 直接找 "@微信" 文本作为锚点（不受模板污染影响）
    pts = _find_at_tag_positions_by_ocr(full)
    if not pts:
        # OCR 没识别到时，再退回绿色横向矩形检测（更稳健，不依赖模板）
        pts = find_at_tag_positions(full, at_templates=None, threshold=0.55)
    # 只保留中间面板的标签（外部联系人列表），过滤顶部"标签"分组标题、左侧侧栏
    mid_pts = [p for p in pts if 200 <= p[0] <= min(w - 50, 700) and p[1] >= 150]
    if not mid_pts:
        log("[warn] 当前页面未识别到外部联系人 @微信 标签，请确认已展开'微信联系人'分组")
        return []

    mid_pts.sort(key=lambda p: p[1])
    log(f"中间面板识别到 {len(mid_pts)} 个 @微信 标签位置: {mid_pts}")

    item_height = 60
    if calib and calib.get("item_height"):
        item_height = max(50, min(80, calib["item_height"]))

    records = []
    seen_y = set()
    skipped_system = 0

    for cx, cy in mid_pts:
        # y 接近的标签合并为同一条目
        if any(abs(cy - sy) < item_height // 2 for sy in seen_y):
            continue
        seen_y.add(cy)

        y1 = max(0, cy - item_height // 2)
        y2 = min(h, cy + item_height // 2)
        x1 = max(0, cx - 150)
        x2 = min(w, cx + 60)
        item = full[y1:y2, x1:x2]
        save_shot(item, f"scan_current_item_{cy}.png")

        # 已通过 @微信 标签位置定位；只要有该标签即视为外部客户，不再做额外绿色校验
        name = extract_external_contact_name(item)
        clean = name.replace(" ", "") if name else ""
        if "企业微信团队" in clean:
            log(f"跳过系统账号: {name}")
            skipped_system += 1
            continue
        # 只跳过明显的非联系人噪声；空名字不跳过
        noise_terms = (".zip", ".exe", "快捷方式", "Jenkins", "预定会议", "加入会议", "添加日历", "位客户")
        if any(t in clean for t in noise_terms):
            log(f"跳过噪声条目: {name}")
            continue

        display_name = name if name else "（未识别名称）"
        records.append({
            "name": display_name,
            "time_text": "",
            "last_date": "",
            "over_half_year": None,
            "note": "外部客户（通讯录-外部联系人）",
        })
        log(f"外部客户: 名字={display_name!r}")

    log(f"扫描完成，有效外部客户 {len(records)} 人，跳过系统账号/噪声 {skipped_system} 条")
    return records


# ---------- 聊天时间分割线解析 ----------
# 企业微信聊天窗口中的时间分割线格式：
#   14:32 / 昨天 14:32 / 星期一 09:00 / 8月30日 15:01 / 2025/8/30 10:00
_TIME_DIVIDER_RE = re.compile(
    r"^(昨天|前天)?"
    r"(星期[一二三四五六天日]|周[一二三四五六天日])?"
    r"(\d{4}[/年.\-]\d{1,2}[/月.\-]\d{1,2}日?|\d{1,2}月\d{1,2}日?|\d{1,2}/\d{1,2})?"
    r"\d{1,2}:\d{2}(:\d{2})?(AM|PM|上午|下午)?$"
)


def _looks_like_chat_time(text):
    """判断 OCR 文本是否像聊天时间分割线。"""
    t = (text or "").replace(" ", "").replace("：", ":").replace("凌晨", "").replace("早上", "")
    return bool(_TIME_DIVIDER_RE.match(t))


def parse_chat_time_divider(text, today=None):
    """
    解析聊天时间分割线文本，返回 datetime.date 或 None。
    "14:32" -> 今天；"昨天 09:10" -> 昨天；"星期一 09:00" -> 本周；
    "8月30日 15:01" -> 今年8月30日；"2025/8/30 10:00" -> 2025-08-30。
    """
    if today is None:
        today = datetime.date.today()
    t = (text or "").replace(" ", "").replace("：", ":").replace("凌晨", "").replace("早上", "").replace("上午", "").replace("下午", "")
    if not t:
        return None
    m = re.search(r"\d{1,2}:\d{2}(:\d{2})?(AM|PM|上午|下午)?$", t)
    time_part = m.group(0) if m else ""
    date_part = t[:m.start()] if m else t
    if date_part in ("昨天", "前天"):
        base = today - datetime.timedelta(days=1 if date_part == "昨天" else 2)
        return base if time_part else None
    if not date_part:
        # 纯时间（如 "14:32"）-> 今天
        return today if time_part else None
    d = parse_last_contact(date_part, today=today)
    return d


def read_chat_last_time(hwnd, rect):
    """
    读取当前打开聊天窗口中"最后一条消息时间"。
    扫描聊天内容区域（右侧、输入框以上），取最下方的时间分割线。
    返回 (last_date 或 None, 时间文本)。
    无任何时间分割线时返回 (None, "")。
    """
    full = window_shot(rect)
    h, w = full.shape[:2]
    # 聊天内容区域：避开左侧会话列表（x<320），避开底部输入框
    x0 = max(320, int(w * 0.28))
    # 时间分割线通常在聊天内容最上方（刚进入聊天时只有一条系统提示），y0 必须包含它
    y0 = max(45, int(h * 0.05))
    y1 = max(y0 + 50, h - 150)
    region = full[y0:y1, x0:max(x0 + 50, w - 30)]
    save_shot(region, "check_chat_region.png")
    results = ocr_text(region)
    best_y, best_text = -1, ""
    any_text = bool(results)
    for t, conf, box in results:
        if not _looks_like_chat_time(t):
            continue
        y = box[1]
        if y > best_y:
            best_y, best_text = y, t.strip()
    if best_text:
        return parse_chat_time_divider(best_text), best_text
    # 仍然没有：可能是灰底时间条对比度太低，OCR 漏识别。
    # 截取顶部窄条再跑一次，作为兜底。
    if not best_text:
        top_region = full[max(40, int(h * 0.04)):min(int(h * 0.25), y1), x0:max(x0 + 50, w - 30)]
        if top_region.size:
            save_shot(top_region, "check_chat_region_top.png")
            for t, conf, box in ocr_text(top_region):
                if _looks_like_chat_time(t):
                    best_text = t.strip()
                    break
    if best_text:
        return parse_chat_time_divider(best_text), best_text
    # 没有时间分割线：区分"空会话"和"识别失败"
    return (None, "" if not any_text else "（有内容未识别到时间）")


# ---------- 通讯录导航 ----------
def _contacts_page_ok(full, verbose=False):
    """
    验证当前窗口是否在通讯录外部联系人列表页。
    判据1：中间面板顶部固定区域出现列表页标题"外部联系人"/"微信联系人"
           （聊天页该区域是会话名/搜索框，不会出现这两个标题）。
    判据2（兜底）：中部面板存在 >=2 行 @微信 标签。
    """
    h, w = full.shape[:2]
    # 列表页标题在中间面板顶部固定区域（标题+分组头），限制窄带避免命中聊天正文
    band_y1 = max(260, int(h * 0.25))
    band = full[30:band_y1, int(w * 0.15):int(w * 0.6)]
    try:
        results = ocr_text(band)
    except Exception:
        results = []
    for t, conf, box in results:
        t2 = (t or "").replace(" ", "")
        if t2 == "外部联系人" or t2.startswith("微信联系人"):
            if verbose:
                log(f"页面校验通过：命中列表页标题 {t2!r}")
            return True
    # @微信 标签兜底：要求至少 2 行（聊天页偶发单个"微信"文本误识别）
    pts = _find_at_tag_positions_by_ocr(full)
    mid = [p for p in pts
           if 200 <= p[0] <= min(w - 50, 700) and 150 <= p[1] <= h - 90]
    rows = _cluster_tag_rows(mid)
    if len(rows) >= 2:
        if verbose:
            log(f"页面校验通过：命中 {len(rows)} 行 @微信 标签")
        return True
    return False


def _get_nav_geometry(w):
    """
    根据窗口宽度返回导航栏几何参数。
    企业微信左侧导航栏：图标在左、文字在右水平排列；
    icon_col 只覆盖图标列，避免中间面板干扰。
    """
    # 整个导航栏（图标+文字）宽约 70~90px
    nav_w = min(90, max(70, int(w * 0.05)))
    # 图标列：图标本身约 24~32px，加上左侧边距
    icon_col = min(45, max(35, nav_w - 40))
    return nav_w, icon_col


def _click_nav_item_by_text(hwnd, rect, text, icon_left_offset=30):
    """
    在窗口最左侧导航栏里 OCR 精确匹配 text，点击其左侧图标。
    企业微信导航栏为图标-文字水平排列，文字在图标右侧，
    因此应点击文字左侧 icon_left_offset 像素处（图标中心）。
    返回是否点击成功。
    """
    full = window_shot(rect)
    h, w = full.shape[:2]
    nav_w, icon_col = _get_nav_geometry(w)
    nav_region = full[:, :nav_w]
    try:
        results = ocr_text(nav_region)
    except Exception as e:
        log(f"[warn] OCR 读取导航栏失败: {e}")
        return False
    for t, conf, box in results:
        if (t or "").replace(" ", "") == text:
            x, y, bw, bh = box
            # 点击文字左侧的图标中心，y 与文字中心对齐
            cx = max(icon_col // 2, x - icon_left_offset)
            cy = y + bh // 2
            # 防止越界
            cx = min(cx, nav_w - 10)
            click_at(hwnd, rect, rect[0] + cx, rect[1] + cy)
            log(f"点击导航'{text}'图标（OCR 文字{x},{y} -> 图标{cx},{cy}, conf={conf:.2f}）")
            return True
    return False


def _ocr_nav_icon_pos(full, text, icon_left_offset=30):
    """
    在导航栏内 OCR 精确匹配 text，返回其左侧图标的中心坐标 (cx, cy)。
    找不到返回 None。不做任何点击。
    """
    h, w = full.shape[:2]
    nav_w, icon_col = _get_nav_geometry(w)
    nav_region = full[:, :nav_w]
    try:
        results = ocr_text(nav_region)
    except Exception:
        return None
    for t, conf, box in results:
        if (t or "").replace(" ", "") == text:
            x, y, bw, bh = box
            cx = min(max(icon_col // 2, x - icon_left_offset), nav_w - 10)
            cy = y + bh // 2
            return (cx, cy)
    return None


def click_contacts_nav(hwnd, rect, calib=None):
    """
    点击左侧导航栏"通讯录"图标（不验证页面）。
    严格限定在最左侧图标列内搜索，避免误点中间面板或右侧内容。
    模板优先（阈值 0.75），但模板命中后与 OCR 计算的图标中心交叉校验：
    偏差 > 15px 时以 OCR 坐标为准（模板可能框偏），既防误点又保证点中热区。
    OCR 不可用时退回纯模板；模板未命中才走 OCR 兜底。
    返回是否点击成功。
    """
    full = window_shot(rect)
    h, w = full.shape[:2]
    nav_w, icon_col = _get_nav_geometry(w)
    icon_region = full[:, :icon_col]

    # 0. 先 OCR 定位"通讯录"文字的图标中心（仅用于校验/修正，不点击）
    ocr_pos = _ocr_nav_icon_pos(full, "通讯录", icon_left_offset=30)

    # 1. 模板匹配优先（仅在左侧图标列内）
    tpl = _get_calib_template(calib, "contacts_icon_template_img", "contacts_icon_template")
    if tpl is not None:
        pt, score = find_template(icon_region, tpl, threshold=0.75)
        if pt:
            if ocr_pos and (abs(pt[0] - ocr_pos[0]) > 15 or abs(pt[1] - ocr_pos[1]) > 15):
                log(f"模板中心({pt[0]},{pt[1]})与OCR图标中心({ocr_pos[0]},{ocr_pos[1]})偏差>15px，用OCR坐标修正")
                pt = ocr_pos
            click_at(hwnd, rect, rect[0] + pt[0], rect[1] + pt[1])
            log(f"点击通讯录导航图标（模板 score={score:.2f} -> {pt[0]},{pt[1]}，图标列内）")
            return True

    # 2. OCR 文字定位兜底：在导航栏内精确匹配"通讯录"，点击文字左侧图标
    if _click_nav_item_by_text(hwnd, rect, "通讯录", icon_left_offset=30):
        return True

    log("[warn] 未能定位通讯录导航图标")
    return False


def _click_contacts_nav_by_ocr(hwnd, rect):
    """用通用文字定位点击'通讯录'（用于模板点错后的纠正）。"""
    return _click_nav_item_by_text(hwnd, rect, "通讯录", icon_left_offset=30)


def ensure_contacts_page(hwnd, rect, calib=None, retries=3, force_click=False):
    """
    确保当前窗口回到通讯录外部联系人列表页。
    每次点击后用 OCR 验证页面；点错了自动换方式重试。
    force_click=True：跳过初始"是否已在列表页"检查，直接点导航
    （用于刚从聊天页返回等明确不在列表页的场景，避免聊天页被误判）。
    返回 (是否成功, 最新rect)。
    """
    # 先保存初始位置，用于"先点别的菜单复位"策略
    for attempt in range(1, retries + 1):
        rect = ensure_visible(hwnd, rect)
        if not (force_click and attempt == 1):
            full = window_shot(rect)
            if _contacts_page_ok(full, verbose=True):
                return True, rect  # 已经在列表页
        # 不在列表页 -> 尝试返回
        if attempt == 1:
            click_contacts_nav(hwnd, rect, calib)
        elif attempt == 2:
            # 模板可能点错：换纯 OCR 点击"通讯录"
            if not _click_contacts_nav_by_ocr(hwnd, rect):
                click_contacts_nav(hwnd, rect, calib)
        else:
            # 用户建议：点其它菜单（消息）复位，再点通讯录
            log(f"[warn] 第{attempt-1}次尝试仍未回到列表页，先点击'消息'复位后再点'通讯录'")
            if _click_nav_item_by_text(hwnd, rect, "消息", icon_left_offset=30):
                time.sleep(0.8)
            click_contacts_nav(hwnd, rect, calib)
        time.sleep(1.0)
        rect = ensure_visible(hwnd, rect)
        full = window_shot(rect)
        if _contacts_page_ok(full, verbose=True):
            log(f"已回到通讯录列表页（第{attempt}次尝试）")
            return True, rect
        log(f"[warn] 第{attempt}次点击后仍未回到通讯录列表页，重试")
    log("[warn] 多次尝试后仍未回到通讯录列表页")
    return False, rect


# ---------- 滚动扫描 + 逐个检查聊天时间 ----------
def img_hash(img):
    """截图指纹（缩略后 md5），用于判断滚动是否已到列表底部。"""
    try:
        small = cv2.resize(img, (160, 90))
        return hashlib.md5(np.ascontiguousarray(small).tobytes()).hexdigest()
    except Exception:
        return ""


def _cluster_tag_rows(pts, row_tol=28):
    """把同一行（y 接近）的 @微信 标签合并，每行保留一个。按 y 排序。"""
    rows = []
    for p in sorted(pts, key=lambda p: p[1]):
        if rows and abs(p[1] - rows[-1][1]) <= row_tol:
            continue
        rows.append(p)
    return rows


def _extract_row_name(full, cx, cy):
    """从整屏截图中，以 @微信 标签 (cx,cy) 为中心提取该行联系人名字。"""
    h, w = full.shape[:2]
    y1, y2 = max(0, cy - 32), min(h, cy + 32)
    x1 = max(0, cx - 170)
    x2 = min(w, cx + 60)
    name = extract_external_contact_name(full[y1:y2, x1:x2])
    if not name:
        x1 = max(0, cx - 240)
        name = extract_external_contact_name(full[y1:y2, x1:x2])
    return (name or "").strip()


def open_chat_and_read_time(hwnd, rect, calib, row_x, row_y):
    """
    对列表中 (row_x, row_y) 处的联系人：
      点击行 -> 资料卡 -> 点"发消息"进入聊天 -> 读最后消息时间 -> 返回通讯录。
    注意：不会真正发送任何消息，只是打开发消息的聊天界面查看时间。
    返回 (last_date 或 None, 时间文本, note)。
    """
    l, t, r, b = rect
    w = r - l
    # 1. 点击联系人行（标签左侧的名字区域，同行高度）
    send_click(l + max(40, row_x - 60), t + row_y)
    time.sleep(0.8)

    # 2. 在右侧资料卡找"发消息"按钮
    full = window_shot(rect)
    save_shot(full, "check_profile.png")
    btn = None
    x0 = int(w * 0.45)
    for txt, conf, box in ocr_text(full[:, x0:]):
        if conf < 0.5 or not box:
            continue
        if "发消息" in (txt or "").replace(" ", ""):
            bx, by, bw, bh = box
            btn = (x0 + bx + bw // 2, by + bh // 2)
            break
    if btn:
        send_click(l + btn[0], t + btn[1])
        log(f"点击'发消息'进入聊天: {btn}")
        time.sleep(1.0)
    else:
        log("[warn] 未找到'发消息'按钮，尝试直接读取当前页面时间")

    # 3. 读聊天最后消息时间
    try:
        last_date, time_text = read_chat_last_time(hwnd, rect)
        note = ""
        if last_date is None:
            # 按用户要求：未读到任何时间记录（空会话/识别失败）统一视为需清理（2026.1.1前）
            note = "未读到聊天时间，按2026.1.1前处理" if time_text == "" else "聊天时间识别失败，按2026.1.1前处理"
    except Exception as e:
        log(f"[warn] 读取聊天时间异常: {e}，按2026.1.1前处理")
        last_date, time_text, note = None, "", "读取聊天时间异常，按2026.1.1前处理"

    # 4. 返回通讯录：刚从聊天页返回，明确不在列表页，强制点击导航（不依赖"先验证"捷径）
    ok, rect = ensure_contacts_page(hwnd, rect, calib, force_click=True)
    if not ok:
        log("[warn] 未能确认返回通讯录列表页，继续扫描（若连续失败请人工点击通讯录）")
    time.sleep(0.5)
    return last_date, time_text, note


# 联系人名字噪声过滤（文件名/菜单项等误识别）
_NAME_NOISE_TERMS = (".zip", ".exe", "快捷方式", "Jenkins", "预定会议", "加入会议",
                     "添加日历", "位客户", "企业微信团队")


def scan_contacts_with_scroll(hwnd, rect, calib=None, check_last_time=True,
                               progress=None, should_stop=None):
    """
    滚动扫描"微信联系人"列表中的全部外部客户（支持 1000+ 人）。

    流程（每轮循环）：
      1. 截图当前屏，OCR 找 @微信 标签 -> 按行聚类 -> 提取名字
      2. 找到第一个未处理的联系人：
         - check_last_time=True：点击进入聊天读最后消息时间，判断是否超半年，再返回
         - 否则只记录名字
      3. 本屏全部处理完则向下滚动，直到连续多屏无新联系人（到底部）

    参数：
      progress: 回调 progress(已处理数, 当前名字)
      should_stop: 返回 True 时安全中止
    返回记录列表（name/time_text/last_date/over_half_year/note）。

    预计耗时：检查时间模式下每位客户约 4~6 秒，1000 人约 1~1.5 小时。
    """
    rect = ensure_visible(hwnd, rect)
    l, t, r, b = rect
    w = r - l
    win_h = b - t
    records = []
    processed_names = set()
    processed_ys = []  # 已处理行的 y 坐标，防止选中态下同一行 OCR 名字变化而重复
    no_new_rounds = 0
    last_hash = ""
    stable_screens = 0
    expanded_retry = False
    max_rounds = 8000

    if check_last_time:
        log("提示：逐个检查聊天时间较慢（每人约 4~6 秒），1000 人约需 1~1.5 小时，可随时点击【停止】")

    round_idx = 0
    while round_idx < max_rounds:
        round_idx += 1
        if should_stop and should_stop():
            log("收到停止请求，安全中止扫描")
            break

        full = window_shot(rect)
        save_shot(full, "scroll_screen.png")
        h = full.shape[:2][0]

        # 找当前屏的 @微信 标签
        pts = _find_at_tag_positions_by_ocr(full)
        if not pts:
            pts = find_at_tag_positions(full, at_templates=None, threshold=0.55)
        mid = [p for p in pts
               if 200 <= p[0] <= min(w - 50, 700) and 150 <= p[1] <= h - 90]
        rows = _cluster_tag_rows(mid)

        # 找第一个未处理的行
        target = None
        for cx, cy in rows:
            name = _extract_row_name(full, cx, cy)
            clean = name.replace(" ", "")
            if not clean:
                continue  # 名字未识别的行跳过（避免重复处理死循环）
            if any(nz in clean for nz in _NAME_NOISE_TERMS):
                continue
            if clean in processed_names:
                continue
            # 同一物理行（y 接近）已处理过，跳过。点击后联系人会高亮，
            # OCR 可能把高亮/图标干扰读成另一个名字（如 HY -> HY引）。
            if any(abs(cy - py) <= 30 for py in processed_ys):
                continue
            target = (cx, cy, name, clean)
            break

        if target:
            cx, cy, name, key = target
            processed_names.add(key)
            processed_ys.append(cy)
            if check_last_time:
                last_date, time_text, note = open_chat_and_read_time(
                    hwnd, rect, calib, cx, cy)
                # 用户要求：只要没读到有效时间记录，统一视为需清理（2026.1.1前）
                if last_date is None:
                    over = True
                else:
                    over = is_over_half_year(last_date)
                if not note:
                    note = "外部客户（微信联系人）"
            else:
                last_date, time_text, over = None, "", None
                note = "外部客户（微信联系人）｜未检查聊天时间"
            records.append({
                "name": name,
                "time_text": time_text,
                "last_date": last_date.isoformat() if last_date else "",
                "over_half_year": over,
                "note": note,
                "y": cy,  # 临时记录 y，用于后处理去重
            })
            log(f"[{len(records)}] {name} 最后联系={time_text or '未知'} "
                f"2026.1.1前未联系={'是' if over else ('否' if over is False else '未知')}")
            if progress:
                progress(len(records), name)
            no_new_rounds = 0
            continue  # 页面已变化，重新截图

        # 本屏没有未处理条目 -> 滚动
        cur_hash = img_hash(full)
        if cur_hash and cur_hash == last_hash:
            stable_screens += 1
            if stable_screens >= 2:
                # 到底前先验证还在列表页，防止误导航走后把错误页面当"底部"
                if not _contacts_page_ok(full):
                    log("检测到已离开通讯录列表页，尝试导航回来")
                    ok, rect = ensure_contacts_page(hwnd, rect, calib)
                    if ok:
                        expand_wechat_contacts_folder(hwnd, rect, calib)
                        stable_screens = 0
                        last_hash = ""
                        continue
                log("已滚动到列表底部，扫描结束")
                break
        else:
            stable_screens = 0
        last_hash = cur_hash

        no_new_rounds += 1
        if no_new_rounds >= 3 and not expanded_retry and not rows:
            # 连续几屏都看不到标签：先验证是否还在通讯录列表页（可能被误导航走）
            ok, rect = ensure_contacts_page(hwnd, rect, calib)
            if ok:
                log("重新展开'微信联系人'分组")
                expand_wechat_contacts_folder(hwnd, rect, calib)
            expanded_retry = True
            rect = ensure_visible(hwnd, rect)
            continue
        if no_new_rounds > 60:
            log("连续多屏未发现新联系人，结束扫描")
            break

        scroll_at(l + 420, t + win_h // 2, notches=5)
        # 滚动后行位置全部变化：行高固定（65px），绝对 y 坐标会周期性复用，
        # 若不清空，新屏露出的新行 y 会撞上已处理行的 y 而被误判为"已处理"，导致漏扫。
        # y 去重只在"未滚动"时可信（防同一行高亮后 OCR 名字变体重复处理）。
        processed_ys.clear()
        time.sleep(0.45)

    # 后处理：合并重复。同一行被点击后会高亮，OCR 可能读出不同名字
    # （HY / HY引，人... / 人.）。同屏（未滚动）时靠 y 坐标相同判断；
    # 滚动后绝对 y 已不可靠，改用"时间相同 + 名字互相包含"兜底合并。
    merged = []
    for rec in records:
        dup = False
        for m in merged:
            y_same = abs(rec.get("y", -9999) - m.get("y", -9999)) <= 40
            time_same = rec["time_text"] and rec["time_text"] == m["time_text"]
            name_rel = (rec["name"] in m["name"]) or (m["name"] in rec["name"])
            if time_same and (y_same or name_rel):
                # 保留名字更完整/更长的一条
                if len(rec["name"]) > len(m["name"]):
                    m["name"] = rec["name"]
                dup = True
                break
        if not dup:
            merged.append(rec)
    records = merged

    # 去掉临时 y 字段后再返回
    for rec in records:
        rec.pop("y", None)

    # 按名字去重兜底（OCR 可能对同一人产生完全相同的重复）
    seen = set()
    uniq = []
    for rec in records:
        if rec["name"] in seen:
            continue
        seen.add(rec["name"])
        uniq.append(rec)
    if len(uniq) != len(records):
        log(f"去重：{len(records)} -> {len(uniq)} 人")
    log(f"滚动扫描完成，共识别 {len(uniq)} 位外部客户")
    return uniq


# ---------- 消息列表扫描 ----------
def split_sessions(list_img, item_height=None, min_height=60):
    """
    将消息列表区域按固定条目高度分割。
    优先使用校准得到的 item_height；否则按 65px 估算并向上取整到可整除。
    """
    h, w = list_img.shape[:2]
    if item_height is None:
        item_height = 65
    item_height = max(min_height, int(item_height))
    items = []
    y = 0
    while y + item_height <= h:
        items.append((y, y + item_height))
        y += item_height
    return items


def extract_name_time(item_img):
    """
    从会话条目中截取名字区域和时间区域。
    返回 (name_region_bgr, time_region_bgr)。
    消息条目布局：头像(左) | 名字 + 消息预览 | 时间(右)
    名字区域取较宽范围（0.18~0.80），避免截掉 @微信 标签。
    """
    h, w = item_img.shape[:2]
    # 名字在上半部分；@微信 标签可能较靠右，时间在最右侧
    name_region = item_img[:h * 5 // 8, int(w * 0.18):int(w * 0.80)]
    time_region = item_img[:, int(w * 0.80):]
    return name_region, time_region


def ocr_name(name_region):
    """识别名字区域第一行（名字本身）。"""
    results = ocr_text(name_region)
    if not results:
        return ""
    # 按 y 坐标排序（优先最上方的文本）
    results.sort(key=lambda item: item[2][1])
    # 过滤掉 @微信 标签文本、emoji 等
    parts = []
    for t, conf, box in results:
        tt = t.replace(" ", "").strip()
        if not tt:
            continue
        if "@微信" in tt or tt == "@":
            continue
        # 过滤纯数字/短串？先不过滤
        parts.append(tt)
    # 取最上方非标签文本作为名字
    return parts[0] if parts else ""


def scan_message_list(hwnd, rect, calib=None):
    """
    扫描当前消息列表，返回外部客户记录列表。
    calib: 校准结果 dict（含 at_templates, list_region, item_height 等）
    """
    l, t, r, b = rect
    full = window_shot(rect)
    save_shot(full, "scan_full.png")

    # 消息列表区域：x 范围用校准值，y 下边界用当前窗口高度（列表通常占满窗口）
    win_h = b - t
    if calib and calib.get("list_region"):
        lx, ly, rx, _ = calib["list_region"]
        ry = win_h - 10
    else:
        lx, ly = 60, 60
        rx, ry = 340, win_h - 20
    list_img = full[ly:ry, lx:rx]
    save_shot(list_img, "scan_list_region.png")
    log(f"消息列表区域: ({lx},{ly})-({rx},{ry})")

    # 条目高度：优先校准值
    item_height = calib.get("item_height") if calib else None
    items = split_sessions(list_img, item_height=item_height)
    log(f"分割出 {len(items)} 个候选条目")

    at_templates = calib.get("at_templates") if calib else None

    records = []
    for idx, (y1, y2) in enumerate(items):
        item = list_img[y1:y2, :]
        if not is_external_contact(item, at_templates=at_templates):
            continue
        # 名字/时间提取：整条 OCR + 位置分类
        blue = _item_has_blue_bg(item)
        name, time_text = extract_customer_info(item, is_blue=blue)

        # 过滤系统内置账号
        if name and "企业微信团队" in name.replace(" ", ""):
            log(f"跳过系统账号 #{idx}: {name}")
            continue

        last_date = parse_last_contact(time_text)
        over = is_over_half_year(last_date)

        records.append({
            "name": name,
            "time_text": time_text,
            "last_date": last_date.isoformat() if last_date else "",
            "over_half_year": over,
            "note": "外部客户（消息列表）" + ("[选中项]" if blue else ""),
        })
        log(f"外部客户 #{idx}: 名字={name!r} 时间={time_text!r} 2026.1.1前未联系={over}")

    return records


# ---------- Excel 导出 ----------
def export_excel(records, filename="2026.1.1前未联系客户名单.xlsx"):
    """将记录导出为 Excel。"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        raise RuntimeError("缺少 openpyxl，请运行: pip install openpyxl")

    wb = Workbook()
    ws = wb.active
    ws.title = "客户名单"
    headers = ["序号", "客户名称", "最后联系时间", "识别日期", "2026.1.1前未联系", "备注"]
    ws.append(headers)
    for i, r in enumerate(records, 1):
        over = r.get("over_half_year")
        if over is True:
            over_txt = "是（建议删除）"
        elif over is False:
            over_txt = "否"
        else:
            over_txt = "无法判断（待确认）"
        ws.append([
            i,
            r.get("name", ""),
            r.get("time_text", ""),
            r.get("last_date", ""),
            over_txt,
            r.get("note", ""),
        ])

    header_fill = PatternFill("solid", fgColor="B4C7DC")
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    path = os.path.join(OUTPUT, filename)
    wb.save(path)
    log(f"Excel 已保存: {path}")
    return path


# ---------- 单元自测 ----------
if __name__ == "__main__":
    print("===== 时间解析自测 =====")
    today = datetime.date(2026, 8, 31)  # 周一
    cases = [
        ("15:23", today, False),
        ("昨天", datetime.date(2026, 8, 30), False),
        ("星期一", datetime.date(2026, 8, 24), False),
        ("8月10日", datetime.date(2026, 8, 10), False),
        ("2月1日", datetime.date(2026, 2, 1), True),
        ("2025/12/1", datetime.date(2025, 12, 1), True),
        ("2025年3月15日", datetime.date(2025, 3, 15), True),
        ("2026/8/1", datetime.date(2026, 8, 1), False),
    ]
    for text, expected_date, expected_over in cases:
        d = parse_last_contact(text, today=today)
        over = is_over_half_year(d, today=today)
        ok = (d == expected_date and over == expected_over)
        print(f"{'OK ' if ok else 'FAIL'} {text!r:20} -> {d}  2026.1.1前={over}")
    print("清理基准日期:", CUTOFF_DATE)

    print("===== 聊天时间分割线自测 =====")
    today = datetime.date(2026, 9, 1)  # 周二
    divider_cases = [
        ("14:32", today, False),
        ("昨天 09:10", datetime.date(2026, 8, 31), False),
        ("星期一 14:30", datetime.date(2026, 8, 31), False),
        ("8月30日 15:01", datetime.date(2026, 8, 30), False),
        ("3月5日 10:00", datetime.date(2026, 3, 5), True),
        ("2025/8/30 10:00", datetime.date(2025, 8, 30), True),
        ("2025年12月1日 09:00", datetime.date(2025, 12, 1), True),
        ("上午 10:23", today, False),
        ("你好", None, None),
        ("8月30日", datetime.date(2026, 8, 30), False),  # 无时间部分时按纯日期解析
    ]
    for text, expected_date, expected_over in divider_cases:
        d = parse_chat_time_divider(text, today=today)
        over = is_over_half_year(d, today=today) if d else None
        ok = (d == expected_date)
        print(f"{'OK ' if ok else 'FAIL'} {text!r:24} -> {d}  2026.1.1前={over}")

    print("===== 分割线格式判定自测 =====")
    for t, expect in [("14:32", True), ("昨天14:32", True), ("星期一09:00", True),
                      ("8月30日15:01", True), ("2025/8/3010:00", True),
                      ("你好", False), ("2026-08-30", False), ("今天天气不错", False)]:
        got = _looks_like_chat_time(t)
        print(f"{'OK ' if got == expect else 'FAIL'} {t!r:20} -> {got}")
