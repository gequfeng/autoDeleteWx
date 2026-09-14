# -*- coding: utf-8 -*-
"""
企业微信外部客户清理工具 - GUI 主程序
=====================================
提供：截图校准、消息列表扫描、客户勾选、Excel 导出、自动删除。

运行方式：
  开发：  .venv\Scripts\python.exe wxwork_cleanup.py
  交付：  将 wxwork_cleanup.exe 与 1.png~4.png 放在同一目录运行。
"""
import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog

# 路径：打包后数据文件与 exe 同目录
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

from wxc_core import (
    find_main_window, activate, scan_contacts_with_scroll, export_excel,
    log as core_log, CALIB_FILE, OUTPUT,
)
from wxc_calibrate import calibrate, load_calibration
import wxc_delete  # 确保 PyInstaller 能打包该模块


# ---------- 日志重定向到 GUI ----------
class GUILogger:
    def __init__(self, text_widget):
        self.widget = text_widget

    def __call__(self, msg):
        core_log(msg)
        self._append(msg)

    def _append(self, msg):
        # 让主线程更新
        if self.widget.winfo_exists():
            self.widget.after(0, lambda: self._do_append(msg))

    def _do_append(self, msg):
        try:
            self.widget.configure(state="normal")
            self.widget.insert(tk.END, f"{msg}\n")
            self.widget.see(tk.END)
            self.widget.configure(state="disabled")
        except tk.TclError:
            pass


# ---------- 主窗口 ----------
class WxWorkCleanupGUI:
    CHECK_COL = "选择"

    def __init__(self, root):
        self.root = root
        root.title("企业微信外部客户清理工具")
        root.geometry("980x720")
        root.minsize(820, 560)
        try:
            root.iconbitmap(os.path.join(APP_DIR, "app.ico"))
        except Exception:
            pass

        self.records = []
        self.calib = None
        self.stop_event = threading.Event()
        self.scanning = False

        self._build_ui()
        self.logger = GUILogger(self.log_text)
        # 把核心库/校准模块的 log 函数也重定向到 GUI
        import wxc_core
        wxc_core.log = self.logger
        import wxc_calibrate
        wxc_calibrate.log = self.logger

        self.logger("程序启动")
        self.logger(f"工作目录: {APP_DIR}")
        self._load_calib_silent()

    def _build_ui(self):
        # 顶部按钮区
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill=tk.X)

        ttk.Button(top, text="📷 截图校准", command=self.on_calibrate).pack(side=tk.LEFT, padx=5)
        self.scan_btn = ttk.Button(top, text="🔍 扫描客户", command=self.on_scan)
        self.scan_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = ttk.Button(top, text="⏹ 停止", command=self.on_stop_scan, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        self.check_time_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="逐个检查聊天时间（较慢）",
                        variable=self.check_time_var).pack(side=tk.LEFT, padx=(10, 2))
        ttk.Button(top, text="☑ 全选", command=lambda: self.set_all_checks(True)).pack(side=tk.LEFT, padx=5)
        ttk.Button(top, text="☐ 取消全选", command=lambda: self.set_all_checks(False)).pack(side=tk.LEFT, padx=5)
        ttk.Button(top, text="☑ 选中2026.1.1前未联系", command=self.check_over_half_year).pack(side=tk.LEFT, padx=5)
        ttk.Button(top, text="📤 导出 Excel", command=self.on_export).pack(side=tk.LEFT, padx=5)
        ttk.Button(top, text="🗑 删除选中", command=self.on_delete).pack(side=tk.LEFT, padx=5)
        ttk.Button(top, text="❌ 退出", command=self.root.quit).pack(side=tk.RIGHT, padx=5)

        # 表格区
        tree_frame = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        tree_frame.pack(fill=tk.BOTH, expand=True)

        cols = (self.CHECK_COL, "客户名称", "最后联系时间", "识别日期", "2026.1.1前未联系", "备注")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="browse")
        for c in cols:
            self.tree.heading(c, text=c)
        self.tree.column(self.CHECK_COL, width=55, anchor="center")
        self.tree.column("客户名称", width=180)
        self.tree.column("最后联系时间", width=110, anchor="center")
        self.tree.column("识别日期", width=110, anchor="center")
        self.tree.column("2026.1.1前未联系", width=130, anchor="center")
        self.tree.column("备注", width=220)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # 点击首列切换勾选
        self.tree.bind("<ButtonRelease-1>", self.on_tree_click)

        # 状态栏
        self.status = ttk.Label(self.root, text="就绪", padding=(10, 2))
        self.status.pack(fill=tk.X)

        # 日志区
        self.log_text = scrolledtext.ScrolledText(self.root, height=9, state="disabled", wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=False, padx=10, pady=(0, 10))

    def _set_status(self, text):
        self.root.after(0, lambda: self.status.config(text=text))

    def _load_calib_silent(self):
        try:
            self.calib = load_calibration()
            if self.calib:
                self.logger(f"已加载校准文件: {CALIB_FILE}")
        except Exception as e:
            self.logger(f"[warn] 加载校准文件失败: {e}")

    # ---------- 事件处理 ----------
    def on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        if col == "#1":  # 首列
            item = self.tree.identify_row(event.y)
            if item:
                self._toggle_item(item)

    def _toggle_item(self, item):
        idx = int(item)
        rec = self.records[idx]
        rec["checked"] = not rec.get("checked", False)
        self.tree.set(item, self.CHECK_COL, "☑" if rec["checked"] else "☐")

    def set_all_checks(self, state):
        for i, rec in enumerate(self.records):
            rec["checked"] = state
            self.tree.set(str(i), self.CHECK_COL, "☑" if state else "☐")

    def check_over_half_year(self):
        """只勾选 2026.1.1 前未联系的客户。"""
        count = 0
        for i, rec in enumerate(self.records):
            over = bool(rec.get("over_half_year"))
            rec["checked"] = over
            self.tree.set(str(i), self.CHECK_COL, "☑" if over else "☐")
            if over:
                count += 1
        self._set_status(f"已选中 2026.1.1 前未联系客户 {count} 人")
        self.logger(f"勾选 2026.1.1 前未联系客户: {count} 人")

    # ---------- 校准 ----------
    def on_calibrate(self):
        self._set_status("正在校准...")
        threading.Thread(target=self._calibrate_thread, daemon=True).start()

    def _calibrate_thread(self):
        try:
            self.calib = calibrate(APP_DIR)
            self.logger("校准完成")
            self._set_status(f"校准完成 | 列表区域 {self.calib.get('list_region')} | 条目高度 {self.calib.get('item_height')}")
        except Exception as e:
            self.logger(f"[error] 校准失败: {e}")
            self._set_status("校准失败")

    # ---------- 扫描 ----------
    def on_scan(self):
        if self.scanning:
            messagebox.showinfo("扫描进行中", "当前正在扫描，请点击【停止】中止。")
            return
        if not self.calib:
            # 尝试加载
            self._load_calib_silent()
            if not self.calib:
                messagebox.showwarning("未校准", "未找到 calibration.json，请先点击【截图校准】。\n"
                                      "将 1.png~4.png 与本程序放在同一目录即可。")
                return
        if self.check_time_var.get():
            messagebox.showinfo(
                "扫描提示",
                "将执行滚动扫描 + 逐个检查聊天时间：\n\n"
                "1. 请先在企业微信中手动打开\n"
                "    通讯录 → 外部联系人 → 微信联系人\n"
                "2. 扫描过程中会自动滚动列表、逐个点开联系人查看聊天时间，\n"
                "    不会真正发送任何消息；期间请勿操作鼠标键盘\n"
                "3. 1000 位客户约需 1~1.5 小时，可随时点击【停止】安全中止\n\n"
                "确认开始扫描？")
        self._set_status("正在扫描...")
        self.scanning = True
        self.stop_event.clear()
        self.scan_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        threading.Thread(target=self._scan_thread, daemon=True).start()

    def on_stop_scan(self):
        if not self.scanning:
            return
        self.stop_event.set()
        self._set_status("正在停止（等待当前联系人处理完成）...")
        self.logger("已请求停止，等待当前联系人处理完成后中止...")

    def _scan_thread(self):
        try:
            found = find_main_window()
            if not found:
                self.logger("[error] 未找到企业微信主窗口，请确认已登录并最大化。")
                self._set_status("未找到企业微信窗口")
                return
            hwnd, rect = found
            self.logger(f"窗口: {rect}")
            rect = activate(hwnd)  # 恢复/最大化后重新取坐标（最小化窗口 rect 为 -32000）
            if not rect:
                self.logger("[error] 无法获取窗口坐标，请手动打开企业微信窗口后重试。")
                self._set_status("窗口不可用")
                return

            def progress(n, name):
                self._set_status(f"扫描中... 已识别 {n} 人（当前: {name}）")

            records = scan_contacts_with_scroll(
                hwnd, rect, calib=self.calib,
                check_last_time=self.check_time_var.get(),
                progress=progress,
                should_stop=self.stop_event.is_set)
            # 2026.1.1 前未联系的默认勾选（作为待删除候选）
            for r in records:
                r["checked"] = r.get("over_half_year", False)
            self.root.after(0, lambda: self.populate_tree(records))
            self.logger(f"扫描完成，共 {len(records)} 位外部客户")
            self._set_status(f"扫描完成 | 外部客户 {len(records)} 人")
        except Exception as e:
            self.logger(f"[error] 扫描失败: {e}")
            self.logger(traceback.format_exc())
            self._set_status("扫描失败")
        finally:
            self.scanning = False
            self.root.after(0, lambda: (
                self.scan_btn.config(state=tk.NORMAL),
                self.stop_btn.config(state=tk.DISABLED)))

    def populate_tree(self, records):
        self.records = records
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, r in enumerate(records):
            over = r.get("over_half_year")
            over_txt = "是" if over is True else ("否" if over is False else "无法判断")
            check = "☑" if r.get("checked") else "☐"
            self.tree.insert(
                "", tk.END, iid=str(i),
                values=(
                    check,
                    r.get("name", "") or "(未识别)",
                    r.get("time_text", ""),
                    r.get("last_date", ""),
                    over_txt,
                    r.get("note", ""),
                ),
            )

    # ---------- 导出 ----------
    def on_export(self):
        selected = [r for r in self.records if r.get("checked")]
        if not selected:
            selected = self.records
        if not selected:
            messagebox.showinfo("无数据", "当前没有可导出的记录。")
            return
        try:
            path = export_excel(selected)
            self.logger(f"已导出: {path}")
            messagebox.showinfo("导出成功", f"已保存到:\n{path}")
        except Exception as e:
            self.logger(f"[error] 导出失败: {e}")
            messagebox.showerror("导出失败", str(e))

    # ---------- 删除 ----------
    def on_delete(self):
        targets = [r for r in self.records if r.get("checked")]
        if not targets:
            messagebox.showwarning("未选择", "请先勾选要删除的客户。")
            return
        names = "\n".join([f"  • {r.get('name') or '(未识别)'}" for r in targets])
        if not messagebox.askyesno("确认删除", f"即将自动删除以下 {len(targets)} 位外部客户：\n{names}\n\n"
                                    "删除过程中请勿操作鼠标键盘，是否继续？"):
            return
        self._set_status("正在删除...")
        threading.Thread(target=self._delete_thread, args=(targets,), daemon=True).start()

    def _delete_thread(self, targets):
        try:
            found = find_main_window()
            if not found:
                self.logger("[error] 未找到企业微信窗口")
                self._set_status("未找到企业微信窗口")
                return
            hwnd, rect = found
            rect = activate(hwnd)
            if not rect:
                self.logger("[error] 无法获取窗口坐标，请手动打开企业微信窗口后重试。")
                self._set_status("窗口不可用")
                return
            # 逐个按名字滚动查找并删除
            for r in targets:
                name = r.get("name") or "(未识别)"
                self.logger(f"准备删除: {name}")
                # 实际删除委托给 wxc_delete
                try:
                    import wxc_delete
                    ok, info = wxc_delete.delete_contact(hwnd, rect, r, self.calib)
                    if ok:
                        self.logger(f"已删除: {name}")
                    else:
                        self.logger(f"[warn] 删除 {name} 未成功: {info}")
                        break  # 失败则停止，避免误操作
                except Exception as e:
                    self.logger(f"[error] 删除 {name} 出错: {e}")
                    break
            self._set_status("删除流程结束")
        except Exception as e:
            self.logger(f"[error] 删除流程异常: {e}")
            self.logger(traceback.format_exc())
            self._set_status("删除异常")


def main():
    root = tk.Tk()
    WxWorkCleanupGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
