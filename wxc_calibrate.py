# -*- coding: utf-8 -*-
"""
企业微信外部客户清理工具 - 截图校准模块
=====================================
读取 exe 同目录下的校准截图（1.png ~ 4.png 或任意 png/jpg），
自动提取关键界面元素，生成 calibration.json 供扫描流程使用。

校准目标：
  1. 通讯录图标模板（contacts_icon）
  2. @微信 标签模板（at_tag_templates）
  3. 消息列表区域（list_region）
  4. 会话条目高度（item_height）
"""
import ctypes
import datetime
import json
import os
import sys

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

import cv2
import numpy as np
from wxc_core import (BASE, TEMPLATES, RECON, save_shot, cv2_imread, cv2_imwrite, log,
                      find_all_templates, is_external_contact, ocr_text)

CALIB_FILE = os.path.join(BASE, "calibration.json")


def read_calibration_images(exe_dir=None):
    """读取 exe 同目录下的校准截图。返回 [(文件名, img_bgr), ...]。"""
    if exe_dir is None:
        exe_dir = BASE
    exts = (".png", ".jpg", ".jpeg", ".bmp")
    files = []
    for f in sorted(os.listdir(exe_dir)):
        if f.lower().endswith(exts):
            path = os.path.join(exe_dir, f)
            img = cv2_imread(path)
            if img is not None:
                files.append((f, img))
    # 优先把 1.png ~ 4.png 排到前面
    def sort_key(item):
        name = item[0]
        m = __import__("re").match(r"(\d+)\.", name)
        if m:
            return (0, int(m.group(1)), name)
        return (1, 0, name)
    files.sort(key=sort_key)
    return files


def detect_sidebar_icons(img, sidebar_width=80):
    """
    检测左侧导航栏中的图标。
    返回按 y 排序的图标 bounding box 列表 [(x, y, w, h), ...]。
    """
    h, w = img.shape[:2]
    strip = img[:, :min(sidebar_width, w)]
    gray = cv2.cvtColor(strip, cv2.COLOR_BGR2GRAY)
    # 企业微信侧边栏深色背景；图标为亮色。用自适应阈值把图标分离出来。
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    _, binary = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    # 形态学过滤，连接图标内部断点
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    binary = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)
    contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    icons = []
    for c in contours:
        x, y, cw, ch = cv2.boundingRect(c)
        # 过滤噪点和过大的块
        if cw < 12 or ch < 12 or cw > 60 or ch > 60:
            continue
        # 中心必须在侧边栏中间区域（排除最边缘）
        cx = x + cw // 2
        if cx < 10 or cx > sidebar_width - 10:
            continue
        # 图标填充率
        area = cv2.contourArea(c)
        if area < 80:
            continue
        icons.append((x, y, cw, ch))

    # 合并同一图标可能出现的重复检测（按 y 聚类）
    icons.sort(key=lambda b: b[1])
    merged = []
    for b in icons:
        x, y, cw, ch = b
        dup = False
        for mx, my, mw, mh in merged:
            if abs(y - my) < ch and abs(x - mx) < cw:
                # 合并为并集
                nx = min(x, mx)
                ny = min(y, my)
                nr = max(x + cw, mx + mw)
                nb = max(y + ch, my + mh)
                merged[merged.index((mx, my, mw, mh))] = (nx, ny, nr - nx, nb - ny)
                dup = True
                break
        if not dup:
            merged.append(b)

    merged.sort(key=lambda b: b[1])
    return merged


def identify_contacts_icon(icons, img):
    """
    从导航栏图标中识别“通讯录”图标（旧兜底策略）。
    策略：企业微信侧边栏图标顺序不固定（不同版本/权限顺序不同）。
    兜底时取最下方 y 较大的导航图标（通常是通讯录或分组）。
    """
    if not icons:
        return None
    # 去掉顶部头像/搜索区域，y < 40 的不要
    nav_icons = [b for b in icons if b[1] >= 40]
    nav_icons.sort(key=lambda b: b[1])
    if len(nav_icons) >= 2:
        # 取最靠下的主功能图标，更可能是通讯录
        return nav_icons[-1]
    return nav_icons[0] if nav_icons else icons[0]


def identify_contacts_icon_by_ocr(img, icons, sidebar_width=80):
    """
    通过 OCR 识别侧边栏文字“通讯录”，然后定位其左侧图标。
    企业微信导航栏为图标-文字水平排列，文字在图标右侧，
    因此应找文字左侧且 y 中心对齐的图标。
    返回 bbox (x,y,w,h) 或 None。
    """
    if not icons:
        return None
    nav_icons = [b for b in icons if b[0] < sidebar_width and b[1] >= 40]
    if not nav_icons:
        return None

    # 只 OCR 左侧 sidebar 文字区域，减少干扰
    sidebar = img[:, :sidebar_width]
    try:
        results = ocr_text(sidebar)
    except Exception as e:
        log(f"[warn] 侧边栏 OCR 失败: {e}")
        return None

    for txt, conf, box in results:
        if txt and "通讯录" in txt:
            tx, ty, tw, th = box
            ty_cy = ty + th / 2
            best = None
            best_dx = 9999
            for ix, iy, iw, ih in nav_icons:
                # 图标应在文字左侧，且 y 中心接近（允许图标略大）
                if ix + iw <= tx + 5:
                    dx = tx - (ix + iw)
                    dy = abs((iy + ih / 2) - ty_cy)
                    if dy <= max(ih, th, 18) and dx < best_dx:
                        best_dx = dx
                        best = (ix, iy, iw, ih)
            if best:
                return best
    return None


def _is_contacts_page(img, at_positions=None):
    """
    判断截图是否处于“通讯录 -> 外部联系人”页面。
    依据：中间面板出现绿色微信文件夹图标，或中间区域有 @微信 标签。
    """
    from wxc_core import find_green_wechat_icons
    h, w = img.shape[:2]
    # 1. 中间面板有绿色微信图标（文件夹）
    pts = find_green_wechat_icons(img, wechat_icon_template=None, threshold=0.5)
    folder_pts = [p for p in pts if 250 <= p[0] <= min(600, w - 20) and 80 <= p[1] <= min(400, h - 100)]
    if folder_pts:
        return True
    # 2. 中间面板有 @微信 标签
    if at_positions:
        mid_tags = [p for p in at_positions if 250 <= p[0] <= min(600, w - 20) and 100 <= p[1] <= h - 50]
        if len(mid_tags) >= 2:
            return True
    return False


def _detect_selected_sidebar_icon(img, icons, padding=8):
    """
    在侧边栏图标中找出当前高亮（蓝色背景）的那个。
    返回 bbox (x,y,w,h) 或 None。
    """
    if not icons:
        return None
    h, w = img.shape[:2]
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    # 选中态背景：浅蓝/蓝灰
    lower_blue1 = np.array([90, 30, 180])
    upper_blue1 = np.array([130, 200, 255])
    mask1 = cv2.inRange(hsv, lower_blue1, upper_blue1)
    # 再覆盖一种偏灰蓝的选中背景
    lower_blue2 = np.array([100, 20, 140])
    upper_blue2 = np.array([140, 120, 220])
    mask2 = cv2.inRange(hsv, lower_blue2, upper_blue2)
    mask = cv2.bitwise_or(mask1, mask2)

    best = None
    best_ratio = 0.0
    for x, y, cw, ch in icons:
        x1 = max(0, x - padding)
        y1 = max(0, y - padding)
        x2 = min(w, x + cw + padding)
        y2 = min(h, y + ch + padding)
        region = mask[y1:y2, x1:x2]
        if region.size == 0:
            continue
        ratio = np.sum(region > 0) / region.size
        if ratio > best_ratio:
            best_ratio = ratio
            best = (x, y, cw, ch)
    # 阈值放宽一点，避免截图压缩导致蓝色偏淡
    if best_ratio < 0.10:
        return None
    return best


def crop_icon_template(img, bbox, padding=6):
    """从 bbox 裁剪图标模板，加少量 padding。"""
    x, y, w, h = bbox
    x1 = max(0, x - padding)
    y1 = max(0, y - padding)
    x2 = min(img.shape[1], x + w + padding)
    y2 = min(img.shape[0], y + h + padding)
    return img[y1:y2, x1:x2]


def extract_green_wechat_icon_template(img, min_size=14, max_size=34):
    """
    从截图中提取绿色微信图标模板（外部联系人入口 / "微信联系人" 分组图标）。
    返回最佳模板图像或 None。
    策略：找绿色区域中接近正方形的图标块，优先选左侧导航栏里 y>=40 的那个，
    其次选中间分组面板里的图标。
    """
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    lower_green = np.array([35, 40, 40])
    upper_green = np.array([85, 255, 255])
    mask = cv2.inRange(hsv, lower_green, upper_green)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    candidates = []
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        if w < min_size or h < min_size or w > max_size or h > max_size:
            continue
        # 要求接近正方形（微信图标是圆的，宽高比接近 1）
        ratio = max(w, h) / (min(w, h) + 1e-6)
        if ratio > 1.6:
            continue
        area = cv2.contourArea(c)
        if area < 80:
            continue
        cx, cy = x + w // 2, y + h // 2
        # 左侧导航栏里的图标优先（x 很小，y >= 40）
        score = 0
        if x < 90 and y >= 40:
            score += 100
        # 中间分组面板的图标次之（x 在 250~400 之间）
        if 250 <= cx <= 450 and 50 <= cy <= 130:
            score += 50
        # 倾向于更大的实心区域
        score += area / 10.0
        candidates.append((score, x, y, w, h))

    if not candidates:
        return None
    candidates.sort(reverse=True)
    _, x, y, w, h = candidates[0]
    pad = 2
    x1 = max(0, x - pad)
    y1 = max(0, y - pad)
    x2 = min(img.shape[1], x + w + pad)
    y2 = min(img.shape[0], y + h + pad)
    return img[y1:y2, x1:x2]


def extract_at_templates(img, min_conf=4):
    """
    从截图中提取 @微信 标签模板。
    通过绿色区域检测定位标签，裁剪并去重。
    返回模板列表 [template_bgr, ...] 和标签中心点列表 [(x, y), ...]。
    """
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    # 绿色 @微信 标签
    lower_green = np.array([35, 40, 40])
    upper_green = np.array([85, 255, 255])
    mask = cv2.inRange(hsv, lower_green, upper_green)
    # 形态学连接
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    templates = []
    positions = []
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        if w < 8 or h < 8 or w > 80 or h > 40:
            continue
        # 过滤接近正方形的绿色图标（如微信图标 w≈h），@微信 标签是扁的
        ratio = max(w, h) / (min(w, h) + 1e-6)
        if ratio < 1.5:
            continue
        cx, cy = x + w // 2, y + h // 2
        # 避免重复
        dup = False
        for px, py in positions:
            if abs(cx - px) < 15 and abs(cy - py) < 10:
                dup = True
                break
        if dup:
            continue
        positions.append((cx, cy))
        # 裁剪模板（加少量边距）
        pad = 3
        x1 = max(0, x - pad)
        y1 = max(0, y - pad)
        x2 = min(img.shape[1], x + w + pad)
        y2 = min(img.shape[0], y + h + pad)
        templates.append(img[y1:y2, x1:x2])

    # 如果没有绿色标签，尝试用 OCR 检测含 "@微信" 文字的区域并截图（兜底）
    if not templates:
        # 简单兜底：不做复杂 OCR，留给后续扫描时颜色兜底
        pass

    return templates, positions


def infer_list_region(at_positions, img_shape, sidebar_width=70):
    """
    根据 @微信 标签位置推断列表区域。
    支持两种布局：
      1. 消息列表（左侧，x 40~360）
      2. 通讯录 -> 外部联系人列表（中间面板，x 250~650）
    返回 [x1, y1, x2, y2]（相对于截图左上角）。
    """
    if not at_positions:
        return [60, 60, 340, img_shape[0] - 20]

    # 分类：左侧消息列表 vs 中间外部联系人列表
    left_tags = [(x, y) for x, y in at_positions if 40 <= x <= 360 and y >= 45]
    mid_tags = [(x, y) for x, y in at_positions if 250 <= x <= 650 and y >= 45]

    # 优先使用数量多的一方
    if len(mid_tags) >= len(left_tags) and mid_tags:
        # 外部联系人中间列表
        xs = [p[0] for p in mid_tags]
        ys = [p[1] for p in mid_tags]
        x_med = sorted(xs)[len(xs) // 2]
        # 列表左边界：标签中位数往左约 130px（头像+名字宽度）
        x1 = max(160, x_med - 130)
        x2 = min(img_shape[1], max(xs) + 80)
        y1 = max(0, min(ys) - 80)  # 分组标题占用空间
        y2 = img_shape[0] - 10
        if x2 - x1 < 180:
            x2 = x1 + 360
        return [x1, y1, x2, y2]

    if left_tags:
        # 消息列表
        xs = [p[0] for p in left_tags]
        ys = [p[1] for p in left_tags]
        x_med = sorted(xs)[len(xs) // 2]
        x1 = max(sidebar_width, x_med - 55)
        x2 = min(img_shape[1], max(xs) + 80)
        y1 = max(0, min(ys) - 40)
        y2 = img_shape[0] - 10
        if x2 - x1 < 180:
            x1 = max(sidebar_width, x_med - 90)
            x2 = x1 + 280
        return [x1, y1, x2, y2]

    # 默认区域
    return [60, 60, 340, img_shape[0] - 20]


def estimate_item_height(at_positions):
    """根据相邻 @微信 标签的 y 间距估算条目高度。支持消息列表和中间外部联系人列表。"""
    # 同时支持左侧消息列表和中间外部联系人列表
    valid = [(x, y) for x, y in at_positions if (40 <= x <= 360 or 250 <= x <= 650) and y >= 45]
    if len(valid) < 2:
        return 65
    ys = sorted([p[1] for p in valid])
    diffs = [ys[i + 1] - ys[i] for i in range(len(ys) - 1)]
    # 过滤异常大/小的间距（可能跨列表分组）
    diffs = [d for d in diffs if 45 <= d <= 90]
    if not diffs:
        return 65
    # 取众数（四舍五入到 5 的倍数）更稳健，避免跨图异常值拉高估计
    rounded = [int(round(d / 5.0) * 5) for d in diffs]
    counts = {}
    for d in rounded:
        counts[d] = counts.get(d, 0) + 1
    mode = max(counts, key=counts.get)
    return max(55, min(80, mode))


def calibrate(exe_dir=None, save_debug=True):
    """
    主校准流程。读取 exe 同目录截图，生成 calibration.json。
    返回校准结果字典。
    """
    images = read_calibration_images(exe_dir)
    if not images:
        raise RuntimeError(f"未在 {exe_dir or BASE} 找到校准截图（1.png~4.png 或任意 png/jpg）。")

    log(f"读取到 {len(images)} 张校准截图: {[f for f, _ in images]}")

    best_contacts_bbox = None
    best_contacts_img = None
    best_ocr_bbox = None
    best_ocr_img = None
    best_selected_bbox = None
    best_selected_img = None
    best_fallback_bbox = None
    best_fallback_img = None
    all_at_templates = []
    all_at_positions = []
    per_image_positions = []

    for fname, img in images:
        if save_debug:
            save_shot(img, f"calib_{fname}")

        # 先提取 @微信 标签（用于判断是否为通讯录页面）
        tpls, positions = extract_at_templates(img)
        all_at_templates.extend(tpls)
        all_at_positions.extend(positions)
        per_image_positions.append(positions)

        # 1. 检测导航栏图标
        icons = detect_sidebar_icons(img)
        if not icons:
            continue

        # 策略 A（最稳）：OCR 识别“通讯录”文字定位图标，不依赖当前页面
        ocr_bbox = identify_contacts_icon_by_ocr(img, icons)
        if ocr_bbox:
            log(f"[{fname}] OCR 定位到通讯录图标: {ocr_bbox}")
            if best_ocr_bbox is None or ocr_bbox[3] > best_ocr_bbox[3]:
                best_ocr_bbox = ocr_bbox
                best_ocr_img = img

        # 策略 B：若截图是通讯录/外部联系人页面，取当前高亮的侧边栏图标作为通讯录
        if _is_contacts_page(img, at_positions=positions):
            selected = _detect_selected_sidebar_icon(img, icons)
            if selected:
                log(f"[{fname}] 识别到通讯录页面，高亮图标: {selected}")
                if best_selected_bbox is None or selected[3] > best_selected_bbox[3]:
                    best_selected_bbox = selected
                    best_selected_img = img

        # 策略 C（兜底）：按位置取最下方主图标
        cbbox = identify_contacts_icon(icons, img)
        if cbbox and (best_fallback_bbox is None or cbbox[3] > best_fallback_bbox[3]):
            best_fallback_bbox = cbbox
            best_fallback_img = img

    # 优先级：OCR > 高亮 > 兜底
    if best_ocr_bbox is not None:
        best_contacts_bbox = best_ocr_bbox
        best_contacts_img = best_ocr_img
        log("使用 OCR 定位的通讯录图标作为模板")
    elif best_selected_bbox is not None:
        best_contacts_bbox = best_selected_bbox
        best_contacts_img = best_selected_img
        log("使用通讯录页面高亮图标作为通讯录模板")
    elif best_fallback_bbox is not None:
        best_contacts_bbox = best_fallback_bbox
        best_contacts_img = best_fallback_img
        log("[warn] 未通过 OCR/高亮定位通讯录，使用兜底位置图标")
    else:
        log("[warn] 未在校准截图中定位到通讯录图标")

    if best_contacts_bbox is None:
        raise RuntimeError("未能在校准截图中定位到通讯录图标。")

    contacts_template = crop_icon_template(best_contacts_img, best_contacts_bbox)
    contacts_icon_path = os.path.join(TEMPLATES, "contacts_icon.png")
    cv2_imwrite(contacts_icon_path, contacts_template)
    log(f"保存通讯录图标模板: {contacts_icon_path} ({contacts_template.shape[1]}x{contacts_template.shape[0]})")

    # 保存标签模板
    at_tag_paths = []
    for i, tpl in enumerate(all_at_templates[:6]):  # 最多存 6 个
        path = os.path.join(TEMPLATES, f"at_calib_{i:02d}.png")
        cv2_imwrite(path, tpl)
        at_tag_paths.append(path)
    log(f"保存 {len(at_tag_paths)} 个 @微信 标签模板")

    # 推断消息列表/外部联系人列表区域
    list_region = infer_list_region(all_at_positions, best_contacts_img.shape)

    # 按单图分别估算条目高度，再取中位数，避免跨图混合导致异常值
    height_estimates = [estimate_item_height(pos) for pos in per_image_positions]
    height_estimates = [h for h in height_estimates if h]
    if height_estimates:
        item_height = sorted(height_estimates)[len(height_estimates) // 2]
    else:
        item_height = 65

    # 提取绿色微信图标模板（外部联系人入口 / 微信联系人分组）
    wechat_icon_template = None
    wechat_icon_path = None
    for fname, img in images:
        tpl = extract_green_wechat_icon_template(img)
        if tpl is not None:
            wechat_icon_template = tpl
            log(f"[{fname}] 提取到绿色微信图标模板 {tpl.shape[1]}x{tpl.shape[0]}")
            break
    if wechat_icon_template is not None:
        wechat_icon_path = os.path.join(TEMPLATES, "wechat_icon.png")
        cv2_imwrite(wechat_icon_path, wechat_icon_template)
        log(f"保存绿色微信图标模板: {wechat_icon_path} ({wechat_icon_template.shape[1]}x{wechat_icon_template.shape[0]})")
    else:
        log("[warn] 未能从校准截图中提取绿色微信图标模板")

    log(f"推断列表区域: {list_region}, 条目高度: {item_height}px")

    calib = {
        "generated_at": datetime.datetime.now().isoformat(),
        "source_images": [f for f, _ in images],
        # 存相对路径（相对 calibration.json 所在目录），文件夹整体搬家也不失效
        "contacts_icon_template": os.path.join("templates", "contacts_icon.png"),
        "wechat_icon_template": os.path.join("templates", "wechat_icon.png") if wechat_icon_path else None,
        "at_templates": [os.path.join("templates", os.path.basename(p)) for p in at_tag_paths],
        "list_region": list_region,
        "item_height": item_height,
        "window_size_hint": [best_contacts_img.shape[1], best_contacts_img.shape[0]],
        # 校准后立即扫描可直接使用（GUI 不会重新 load_calibration）
        "contacts_icon_template_img": contacts_template,
        "wechat_icon_template_img": wechat_icon_template,
        "at_template_imgs": all_at_templates[:6],
    }

    # 写文件时剔除内存图像（numpy 数组不可 JSON 序列化），加载时再按路径读取
    serializable = {k: v for k, v in calib.items() if not k.endswith("_img") and not k.endswith("_imgs")}
    with open(CALIB_FILE, "w", encoding="utf-8") as f:
        json.dump(serializable, f, ensure_ascii=False, indent=2)
    log(f"校准文件已保存: {CALIB_FILE}")

    return calib


# ---------- 旧版兼容：_is_contacts_page 曾命名为 _is_contacts_page，保留即可 ----------


def load_calibration():
    """加载 calibration.json；不存在则返回 None。

    兼容旧版本（模板存绝对路径）：若路径不存在，
    尝试按 calibration.json 所在目录重新解析。
    """
    if not os.path.exists(CALIB_FILE):
        return None
    with open(CALIB_FILE, "r", encoding="utf-8") as f:
        calib = json.load(f)

    calib_dir = os.path.dirname(os.path.abspath(CALIB_FILE))

    def resolve(p):
        if not p:
            return None
        if os.path.isabs(p) and os.path.exists(p):
            return p
        # 相对路径（新版本）或失效绝对路径（旧版本）→ 按 calib 目录解析
        cand = os.path.join(calib_dir, p)
        if not os.path.exists(cand):
            # 旧版本绝对路径的兜底：取文件名拼到 calib 目录
            cand = os.path.join(calib_dir, os.path.basename(p))
        return cand if os.path.exists(cand) else None

    calib["contacts_icon_template_img"] = cv2_imread(resolve(calib.get("contacts_icon_template", "")))
    calib["wechat_icon_template_img"] = cv2_imread(resolve(calib.get("wechat_icon_template", "")))
    calib["at_template_imgs"] = [cv2_imread(resolve(p)) for p in calib.get("at_templates", [])
                                 if resolve(p)]
    return calib


# ---------- 测试 ----------
if __name__ == "__main__":
    # 用 recon/explore 里的真实截图做校准测试（模拟 exe 同目录）
    import shutil
    test_dir = os.path.join(RECON, "calib_test")
    os.makedirs(test_dir, exist_ok=True)
    # 复制 recon/explore 里的部分截图作为 1.png~4.png
    src_dir = os.path.join(RECON, "explore")
    for i, src in enumerate(["03_external_contacts.png", "04_wechat_contacts_expanded.png",
                              "07_chat_hy.png", "08_chat_window_hy.png"], 1):
        sp = os.path.join(src_dir, src)
        if os.path.exists(sp):
            shutil.copy(sp, os.path.join(test_dir, f"{i}.png"))
    calib = calibrate(test_dir, save_debug=True)
    serializable = {k: v for k, v in calib.items() if not k.endswith("_img") and not k.endswith("_imgs")}
    print(json.dumps(serializable, ensure_ascii=False, indent=2))
