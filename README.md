# pencil2pptx

将 [Pencil](https://pencil.evolveui.com/) `.pen` 文件转换为 PowerPoint `.pptx` 文件。

通过 Pencil MCP server 获取精确的布局数据（坐标、尺寸由 Pencil 引擎计算），渲染为原生 PPT 元素（文本框、形状、图片），不可还原的节点（图标、SVG 路径、图片填充）按需导出为 PNG 保真插入。

## 安装

```bash
pip install pencil2pptx
```

## 使用

前提：Pencil 桌面应用需要已安装（运行时会自动启动）。

```bash
# 基本用法（输出同名 .pptx）
pencil2pptx input.pen

# 指定输出路径
pencil2pptx input.pen -o output.pptx

# 导出指定页码
pencil2pptx input.pen --pages 1,3,5
pencil2pptx input.pen --pages 1-3,7

# 调整字体缩放系数（默认 0.73）
pencil2pptx input.pen --font-scale 0.70

# 指定 Pencil MCP server 路径
pencil2pptx input.pen --pencil-cmd "/path/to/mcp-server"
```

用 `uvx` 免安装运行：

```bash
uvx pencil2pptx input.pen
uvx pencil2pptx input.pen -o output.pptx
uvx pencil2pptx input.pen --font-scale 0.70
```

用 `python -m` 运行：

```bash
python -m pencil2pptx input.pen -o output.pptx
```

## 参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `input` | 输入 .pen 文件路径 | 必填 |
| `-o, --output` | 输出 .pptx 路径 | 同名 .pptx |
| `--pages` | 导出指定页码，如 `1,2` 或 `1-3,5`（从1开始） | 全部 |
| `--font-scale` | 字体缩放系数 (Pencil px → PPT pt) | 0.73 |
| `--pencil-cmd` | Pencil MCP server 可执行文件路径 | Windows: `%LOCALAPPDATA%\Programs\Pencil\...\mcp-server-windows-x64.exe` |
| `--pencil-app` | Pencil 桌面应用路径（用于自动启动） | Windows: `%LOCALAPPDATA%\Programs\Pencil\Pencil.exe` |

## 特性

### 自动启动 Pencil

运行时自动检测 Pencil 桌面应用是否在运行，未运行则自动启动并等待初始化完成，无需手动操作。

### 选择性页面导出

通过 `--pages` 参数指定要导出的页码，支持逗号分隔和范围语法：

```bash
pencil2pptx input.pen --pages 1,3,5
pencil2pptx input.pen --pages 1-3,7
```

### 按需图片导出

不再将整页导出为背景图，而是根据节点类型智能判断：

- **原生 PPT 元素**：`text`、`rectangle`、`ellipse`、`line`、`frame` 等节点直接用 PPT 原生形状渲染，保持可编辑
- **图片填充节点**：`fill` 为 `{"type": "image"}` 的节点，优先通过 `export_nodes` 导出，失败时从 `.pen` 文件的图片 URL 解析本地路径作为 fallback
- **SVG 路径节点**：`path` 类型节点（如 logo）导出为 4x 高清 PNG 插入
- **图标节点**：`icon_font` 类型节点导出为 4x 高清 PNG 插入
- **Context Image**：`context` 属性为 `image`/`img` 的节点整体导出为 PNG，适用于复杂表格、图表等

### 图片透明度与遮罩

图片填充节点的 `opacity` 属性通过 PPT 的 `alphaModFix` 机制正确还原。配合父级 frame 的纯色填充，可精确还原设计稿中的颜色遮罩效果（如蓝色底 + 半透明图片叠加）。

### 形状与文本透明度

所有形状（矩形、椭圆、线条等）和文本节点的 `opacity` 属性均正确还原，通过 PPT 原生 alpha 通道实现。

### 圆角矩形

自动识别节点的 `cornerRadius` 属性，根据圆角大小智能选择形状类型：
- 无圆角 → 标准矩形
- 圆角 ≥ 短边一半 → 椭圆/圆形
- 其他 → 圆角矩形，精确还原圆角半径

### 居中对齐优化

居中对齐的 frame 内单文本节点使用 PPT 原生对齐（水平居中 + 垂直居中），避免坐标偏移导致的对齐误差。

### 无阴影渲染

导出的 PPT 形状默认移除阴影效果，确保与 Pencil 设计稿一致。

### 页面顺序

导出的 PPT 页面顺序与 Pencil 画布上的视觉排列一致，按从上到下、从左到右排序。

## 工作原理

1. 通过 MCP 协议连接正在运行的 Pencil 桌面应用
2. 调用 `snapshot_layout` 获取 Pencil 引擎计算后的精确布局（坐标、尺寸）
3. 按画布位置（先 y 后 x）排序顶层帧，确保页面顺序与 Pencil 视觉顺序一致
4. 调用 `batch_get` 获取节点属性（类型、样式、文本内容等）
5. 按节点类型按需导出图片：
   - `icon_font` → 4x PNG
   - `path` → 4x PNG
   - 图片填充 → 2x PNG（或本地路径 fallback）
   - `context: image/img` → 2x PNG
6. 合并布局和属性数据，渲染为原生 PPT 元素，不可还原的节点以图片插入

## 依赖

- Python 3.10+
- [python-pptx](https://python-pptx.readthedocs.io/)
- [mcp](https://pypi.org/project/mcp/) (Model Context Protocol SDK)
- [Pencil](https://pencil.evolveui.com/) 桌面应用（需运行中）

## 更新日志

### v0.4.0 (2025-04-23)

**新功能**

- 新增 `--pages` 参数，支持选择性导出指定页码（如 `1,3` 或 `1-3,5`）
- 新增 `--pencil-app` 参数，支持自定义 Pencil 桌面应用路径
- 自动检测并启动 Pencil 桌面应用，无需手动打开
- 圆角矩形支持：根据 `cornerRadius` 智能选择矩形/圆角矩形/椭圆
- 居中对齐 frame 内单文本使用 PPT 原生对齐，提升居中精度
- 形状和文本节点的 `opacity` 透明度支持
- 8 位 RGBA 颜色解析，自动忽略完全透明的颜色
- 移除 PPT 形状的默认阴影效果
- 修复文本中 `\v`（垂直制表符）导致 PPT 显示乱码的问题

### v0.3.0 (2025-04-23)

**重构：按需导出替代整页背景图**

- 移除整页背景图导出策略，改为根据节点类型智能判断是否需要导出为图片
- 新增 `path` 类型节点（SVG 路径）的收集与导出支持
- 图片填充节点增加本地路径 fallback：当 `export_nodes` 失败时，从 `.pen` 文件的图片 URL 解析本地路径
- 新增 `_set_picture_opacity` 函数，通过 `alphaModFix` 正确设置图片透明度，还原设计稿中的颜色遮罩效果
- 移除 `skip_bg_shapes` 逻辑，所有节点统一按类型渲染
- 导出速度显著提升（无需等待整页截图）

### v0.2.2

- 初始发布
- 支持 text、rectangle、ellipse、line、frame、icon_font 节点渲染
- 支持 context image 整体导出
- 支持图片填充节点导出
- 支持整页背景图导出（已在 v0.3.0 中移除）

## License

MIT
