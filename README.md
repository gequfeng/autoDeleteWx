# 企业微信外部客户清理工具

基于屏幕 OCR + 模板匹配的企业微信「外部联系人」批量清理辅助工具。
自动扫描外部客户列表、读取每个客户的最后联系时间，标记 **2026-01-01 前未联系** 的客户，并支持自动勾选与批量删除。

> 仅适用于 Windows 桌面版企业微信（wxwork.exe），采用屏幕识别方式，不注入、不篡改客户端。

## 功能

- 🔍 **自动导航**：自动定位并点击左侧「通讯录」→「外部联系人」，带模板 + OCR 双重校验防误点
- 📋 **列表扫描**：OCR 识别外部客户列表（头像/名字/`@微信` 标签），支持 1000+ 长列表自动滚动
- 🕐 **时间读取**：逐个打开会话，OCR 右下角最后联系时间
- ☑️ **一键标记**：2026-01-01 前未联系（含无时间记录）自动勾选
- 🗑️ **批量删除**：自动删除选中联系人
- 📤 **Excel 导出**：扫描结果导出 xlsx 备查

## 文件结构

```
wxwork_cleanup.py    GUI 主程序（tkinter）
wxc_core.py          核心库：窗口检测、截图、OCR、扫描、导出
wxc_calibrate.py     截图校准工具（生成 calibration.json + templates/）
wxc_delete.py        自动删除联系人
wxwork_cleanup.spec        Win10 打包配置（PyInstaller）
wxwork_cleanup_win7.spec   Win7 兼容打包配置
calibration.json     校准结果（运行校准后生成，随包附带一份参考）
```

## 环境要求

### Win10 / Win11
- Python 3.11+
- 依赖见 `requirements.txt`

### Win7
- Python 3.8 + onnxruntime 1.14.1 + rapidocr 3.5.0（见 `wxwork_cleanup_win7.spec`）
- 系统需先补装：KB4474419 → KB4490628 → KB2533623 → KB2999226 + VC++ 运行库

## 使用步骤

1. **校准**（每台机器首次使用 / 企业微信窗口尺寸变化后）：
   ```
   python wxc_calibrate.py
   ```
   按提示截图 `1.png~4.png`，生成 `calibration.json` 与 `templates/`。

2. **扫描**：
   ```
   python wxwork_cleanup.py
   ```
   打开 GUI → 检测企业微信窗口 → 扫描外部客户 → 导出/勾选 → 删除。

3. **打包 exe**：
   ```
   pyinstaller wxwork_cleanup.spec          # Win10 版
   pyinstaller wxwork_cleanup_win7.spec     # Win7 版（需在 py38 环境下执行）
   ```

## 工作原理要点

- 企业微信为 Qt 自绘窗口，鼠标点击必须用 `SendInput` 发送 MOVE → LEFT DOWN → LEFT UP 三步
- 「通讯录」导航：模板匹配 + 同屏 OCR 交叉校正坐标（偏差 >15px 以 OCR 为准）；图标在文字左侧 ~30px 处点击
- 外部客户识别：`@微信` 模板匹配 → 绿色 HSV 兜底 → 蓝色背景 OCR
- 时间提取：整条 item OCR + 位置分类；时间区域反色 + Otsu 二值化
- 基准日期固定为 **2026-01-01**（`CUTOFF_DATE`），无时间记录的联系人按「未联系」处理
- 长列表滚动后清空 y 坐标去重缓存，后处理按「名字互相包含 + 时间相同」合并去重

## 风险提示

⚠️ 删除操作不可撤销。请先「扫描 + 导出 Excel」确认名单后再执行删除，建议先小批量试删。
