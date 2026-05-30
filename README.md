# PPT Touch Controller 触控屏 PPT 查看器

> 专为教学触控大屏设计的 PowerPoint 演示文稿控制工具。全屏放映时，在 PPT 上方显示半透明悬浮按钮，支持触控翻页和拖拽调整位置。

## ✨ 功能

- **透明悬浮按钮** — 按钮浮动在 PPT 全屏窗口上方，背景完全透传点击
- **触控翻页** — 点击 `<` `>` 大按钮翻页（上一页/下一页），兼容触控和鼠标
- **拖拽定位** — 长按中间 `≡` 手柄 500ms 进入拖拽模式，移动整个按钮组到任意位置
- **位置记忆** — 拖拽后自动保存坐标，下次启动恢复
- **左右手模式** — 支持切换按钮顺序（左手：下一页在左）
- **可调外观** — 按钮大小 (60-120px)、透明度 (30-100%)、颜色主题可配置
- **无 PowerPoint 降级** — 未安装 PowerPoint 时，自动使用内置简易查看器渲染静态幻灯片
- **单实例** — 重复打开 .pptx 文件自动复用已有进程

## 📋 运行环境

| 要求 | 说明 |
|------|------|
| Windows 10/11 | 64 位 |
| Python 3.9+ | 推荐 3.11+ |
| Microsoft PowerPoint | 推荐 Office 2016+（非强制，无 PPT 可用降级模式） |

## 🚀 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 运行（打开文件对话框）
python src/main.py

# 3. 或直接指定文件
python src/main.py "C:\path\to\slides.pptx"
```

也可双击 `run.bat` 启动（需确认 Python 路径）。

## 🎮 使用方式

| 操作 | 效果 |
|------|------|
| 点击 `<` | 上一页 / 上一个动画 |
| 点击 `>` | 下一页 / 下一个动画 |
| 长按 ≡ 500ms 后拖拽 | 移动按钮组位置 |
| 点击 ≡ （短按） | 无操作（仅拖拽） |
| PPT 放映结束 | 按钮自动隐藏 |

### 按钮布局

```
   ┌────┐   ┌──┐   ┌────┐
   │ ◀  │   │≡≡│   │ ▶  │
   └────┘   └──┘   └────┘
   翻页按钮  拖拽手柄  翻页按钮
   (80px)   (40px)  (80px)
```

## 🏗️ 项目结构

```
ppt-touch-controller/
├── src/
│   ├── main.py               # 入口，应用生命周期、单实例 IPC
│   ├── overlay_window.py     # 透明悬浮窗、TouchButton、DragHandle、设置面板
│   ├── ppt_controller.py     # PowerPoint COM 自动化 (win32com)
│   ├── settings_manager.py   # JSON 配置持久化 (%APPDATA%)
│   ├── fallback_viewer.py    # 降级方案：python-pptx 静态渲染
│   ├── file_associator.py    # .pptx 文件关联工具
│   └── resources/
│       └── styles/
│           └── overlay.qss   # Qt 样式表
├── build.spec                # PyInstaller 打包配置
├── requirements.txt
├── run.bat                   # Windows 快捷启动脚本
└── README.md
```

## 🔧 技术原理

### 透明穿透窗口

```
┌──────────────────┐
│  PPT 全屏窗口     │  ← 底层
│   ┌──────────┐   │
│   │ [◀ ≡ ▶] │   │  ← WS_EX_LAYERED 透明浮层
│   └──────────┘   │     (WA_TranslucentBackground)
└──────────────────┘
```

- **`WS_EX_LAYERED`** — 窗口支持透明
- **`WM_NCHITTEST`** — Qt `nativeEvent` 拦截消息，返回 `HTTRANSPARENT`（穿透到底层 PPT）或 `HTCLIENT`（按钮区域捕获点击）
- **`QWidget.grabMouse()`** — 拖拽时抓取鼠标，确保光标移出按钮后事件不丢失

### COM 自动化

- 使用 `win32com.client.gencache.EnsureDispatch` 获取早期绑定 COM 代理
- 单线程架构：COM 和 Qt 共存同一线程，避免跨线程信号投递失败
- 每 500ms 轮询 `SlideShowWindows(1)` 检测放映是否结束

### 降级方案

未检测到 PowerPoint 注册表项时，使用 `python-pptx` 解析 PPTX 文件，`Pillow` 渲染幻灯片为图片在 QLabel 中显示。

## 📦 打包为独立 EXE

```bash
# PyInstaller 打包（输出到 dist/）
pyinstaller build.spec

# 生成文件: dist/PPTTouchController.exe
```

打包后无需安装 Python 即可运行。`build.spec` 已排除 Qt 非必要模块（WebEngine、Multimedia、3D 等），压缩后约 55MB。

## ⚙️ 配置

配置文件位置：`%APPDATA%\PPTTouchController\config.json`

```json
{
  "button_size": 80,
  "button_opacity": 75,
  "button_color": "#0078D4",
  "hand_mode": "right",
  "overlay_position": { "x": 960, "y": 900 },
  "confirm_exit": false,
  "last_directory": "C:\\Users\\..."
}
```

| 字段 | 说明 | 默认值 |
|------|------|--------|
| `button_size` | 按钮直径 (px) | 80 |
| `button_opacity` | 透明度 (%) | 75 |
| `button_color` | 颜色主题 | `#0078D4` |
| `hand_mode` | 手模式 `right` / `left` | `right` |
| `overlay_position` | 上次保存的窗口坐标 | 自动居中 |

## ❓ 常见问题

**Q: PowerPoint 未安装？**
A: 程序会弹出对话框询问是否使用内置简易查看器。简易查看器不支持动画和转场，但可以浏览所有幻灯片。也可以从 Microsoft Store 安装免费 PowerPoint 移动版。

**Q: 按钮不显示？**
A: 确认 PowerPoint 已以全屏模式启动（非窗口模式）。检查任务栏是否有 Python 图标，确认程序未崩溃。

**Q: 点击按钮没反应？**
A: 确认点击的是 `<` `>` 按钮区域（非中间拖拽手柄）。尝试缩小 PPT 窗口后重试，确保浮层在 PPT 上方。

**Q: 如何移动按钮到屏幕其他位置？**
A: 长按中间的 `≡` 手柄约半秒，光标变成抓手后拖拽移动。位置会自动保存，下次启动时恢复。

## 📄 License

MIT
