# -*- coding: utf-8 -*-
"""
企业微信外部客户清理工具 - 自动删除模块
=========================================
根据扫描结果（按名字），在企业微信客户端自动滚动查找联系人、
在列表行上右键选择"删除外部联系人"并确认删除。

注意事项：
  1. 删除依赖当前窗口位置与校准数据一致；窗口大小/缩放变化后请重新校准。
  2. 建议先小批量测试，确认删除路径正确后再批量运行。
  3. 普通员工删除外部客户通常只是个人视角移除；企业侧数据可能仍保留。
"""
import time

from wxc_core import (
    log, activate, ensure_visible, window_shot, send_click, send_right_click, ocr_text,
    scroll_at, img_hash,
    click_contacts_nav, ensure_contacts_page, expand_wechat_contacts_folder,
    _find_at_tag_positions_by_ocr, _cluster_tag_rows, _extract_row_name,
    _NAME_NOISE_TERMS,
)


def _find_text_box(img_bgr, keywords, min_conf=0.5):
    """OCR 后查找包含任一关键词的文本框，返回 (x, y, w, h) 或 None。"""
    if isinstance(keywords, str):
        keywords = [keywords]
    results = ocr_text(img_bgr)
    for t, conf, (x, y, w, h) in results:
        if conf < min_conf:
            continue
        for kw in keywords:
            if kw in t:
                return (x, y, w, h)
    return None


def _scroll_list_top(hwnd, rect):
    """把联系人列表滚动回顶部。"""
    l, t, _, _ = rect
    h = rect[3] - rect[1]
    for _ in range(2):
        scroll_at(l + 420, t + h // 2, notches=60, up=True)
        time.sleep(0.4)


def _find_contact_row(full, rect, target_name, fuzzy=False):
    """
    在截图中查找目标联系人所在行。
    返回该行 @微信 标签坐标 (cx, cy) 或 None。
    fuzzy=False：只认去空格后完全相等的名字（第一轮，防误删）。
    fuzzy=True ：本屏无全等时退到互为包含（容忍 OCR 细微差异）。
    """
    h, w = full.shape[:2]
    pts = _find_at_tag_positions_by_ocr(full)
    if not pts:
        from wxc_core import find_at_tag_positions
        pts = find_at_tag_positions(full, at_templates=None, threshold=0.55)
    mid = [p for p in pts
           if 200 <= p[0] <= min(w - 50, 700) and 150 <= p[1] <= h - 90]
    tgt = (target_name or "").replace(" ", "")
    if not tgt:
        return None
    exact, fuzzy_hit = None, None
    for cx, cy in _cluster_tag_rows(mid):
        name = _extract_row_name(full, cx, cy)
        clean = name.replace(" ", "")
        if not clean or any(nz in clean for nz in _NAME_NOISE_TERMS):
            continue
        if clean == tgt:
            exact = (cx, cy)
            break  # 全等命中，直接用
        if fuzzy_hit is None and len(clean) >= 2 and (clean in tgt or tgt in clean):
            fuzzy_hit = (cx, cy)
    if exact:
        return exact
    return fuzzy_hit if fuzzy else None


def _click_menu_item(hwnd, rect, keywords, desc):
    """
    在右键弹出的上下文菜单中查找并点击指定菜单项。
    菜单通常出现在光标附近，先对全屏或中心区域 OCR。
    """
    l, t = rect[0], rect[1]
    time.sleep(0.4)
    full = window_shot(rect)
    # 菜单一般出现在窗口中间偏左区域，优先搜索左半部分
    h, w = full.shape[:2]
    left_half = full[:, :max(w // 2, 600)]
    box = _find_text_box(left_half, keywords)
    if box:
        bx, by, bw, bh = box
        send_click(l + bx + bw // 2, t + by + bh // 2)
        log(f"点击菜单项: {desc}")
        return True
    # 兜底：全屏
    box = _find_text_box(full, keywords)
    if box:
        bx, by, bw, bh = box
        send_click(l + bx + bw // 2, t + by + bh // 2)
        log(f"点击菜单项（全屏兜底）: {desc}")
        return True
    return False


def _confirm_dialog(hwnd, rect):
    """点击删除确认弹窗中的确定按钮。返回是否成功。"""
    l, t = rect[0], rect[1]
    time.sleep(0.5)
    confirm_img = window_shot(rect)
    h, w = confirm_img.shape[:2]
    # 按钮一般在弹窗下方区域（窗口中下部），先只精确匹配"确定"
    lower_region = confirm_img[int(h * 0.55):, :]
    box = _find_text_box(lower_region, ["确定"])
    if box:
        bx, by, bw, bh = box
        by += int(h * 0.55)  # 还原为整图坐标
        send_click(l + bx + bw // 2, t + by + bh // 2)
        log("点击确认删除（确定）")
        time.sleep(0.8)
        return True
    # 兜底："确认"
    box = _find_text_box(lower_region, ["确认"])
    if box:
        bx, by, bw, bh = box
        by += int(h * 0.55)
        send_click(l + bx + bw // 2, t + by + bh // 2)
        log("点击确认删除（确认）")
        time.sleep(0.8)
        return True
    # 再兜底：全屏找"删除"（但可能误点正文，这里仅作最后尝试）
    box = _find_text_box(confirm_img, ["删除"])
    if box:
        bx, by, bw, bh = box
        send_click(l + bx + bw // 2, t + by + bh // 2)
        log("点击确认删除（删除）")
        time.sleep(0.8)
        return True
    return False


def delete_contact(hwnd, rect, record, calib=None):
    """
    按名字删除一位外部客户。
    流程：回通讯录 -> 展开微信联系人 -> 滚动查找目标 -> 右键行 ->
          选择"删除外部联系人" -> 确认弹窗。
    返回 (ok: bool, info: str)。
    """
    name = record.get("name") or ""
    if not name or "未识别" in name:
        return False, "记录缺少有效名字，无法定位"

    new_rect = activate(hwnd)
    if new_rect:
        rect = new_rect
    time.sleep(0.3)

    l, t, r, b = rect
    try:
        # 1. 回到通讯录页面（带验证，误点自动纠正）并展开分组
        ok, rect = ensure_contacts_page(hwnd, rect, calib)
        if not ok:
            return False, "未能回到通讯录列表页"
        time.sleep(0.8)
        rect = ensure_visible(hwnd, rect)
        expand_wechat_contacts_folder(hwnd, rect, calib)
        time.sleep(0.4)

        # 2. 从顶部开始滚动查找目标联系人
        #    第一轮只认全等名字；到底没找到才第二轮模糊匹配
        #    （防止"HY"误删"HY引"这类包含关系误删）
        found = None
        for fuzzy in (False, True):
            _scroll_list_top(hwnd, rect)
            last_hash = ""
            for attempt in range(300):
                full = window_shot(rect)
                found = _find_contact_row(full, rect, name, fuzzy=fuzzy)
                if found:
                    break
                cur = img_hash(full)
                if cur and cur == last_hash:
                    found = None
                    break  # 到底了，进入下一轮或返回失败
                last_hash = cur
                scroll_at(l + 420, t + (b - t) // 2, notches=5)
                time.sleep(0.4)
            if found:
                log(f"定位 {name}（{'模糊' if fuzzy else '全等'}匹配）")
                break
            if not fuzzy:
                log(f"全列表未找到与 {name} 全等的名字，改用模糊匹配重找")
        if not found:
            return False, "已滚动到底部仍未找到该联系人（可能已被删除）"

        # 3. 在目标联系人行上右键，呼出上下文菜单
        cx, cy = found
        row_x = max(60, cx - 80)      # 点击名字/行中央，避免太靠左点到侧边栏
        row_y = cy
        send_right_click(l + row_x, t + row_y)
        log(f"右键联系人行: {name} ({row_x},{row_y})")
        time.sleep(0.6)

        # 4. 点击"删除外部联系人"
        if not _click_menu_item(hwnd, rect, ["删除外部联系人", "删除外部", "删除联系人"], "删除外部联系人"):
            # 可能菜单没出来，尝试再右键一次
            send_right_click(l + row_x, t + row_y)
            time.sleep(0.6)
            if not _click_menu_item(hwnd, rect, ["删除外部联系人", "删除外部", "删除联系人"], "删除外部联系人"):
                return False, "未找到右键菜单中的删除入口"

        # 5. 确认弹窗
        if not _confirm_dialog(hwnd, rect):
            return False, "未找到确认按钮"
        return True, "已删除"
    except Exception as e:
        return False, f"删除过程异常: {e}"


if __name__ == "__main__":
    print("删除模块已加载")
