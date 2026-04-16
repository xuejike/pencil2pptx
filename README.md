# pencil2pptx

将 [Pencil](https://pencil.evolveui.com/) `.pen` 文件转换为 PowerPoint `.pptx` 文件。

通过 Pencil MCP server 获取精确的布局数据（坐标、尺寸由 Pencil 引擎计算），渲染为原生 PPT 元素（文本框、形状、图片），图标通过 PNG 导出保真插入。

## 安装

```bash
pip install pencil2pptx
```

## 使用

前提：Pencil 桌面应用需要正在运行。

```bash
# 基本用法（输出同名 .pptx）
pencil2pptx input.pen

# 指定输出路径
pencil2pptx input.pen -o output.pptx

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
| `--font-scale` | 字体缩放系数 (Pencil px → PPT pt) | 0.73 |
| `--pencil-cmd` | Pencil MCP server 可执行文件路径 | Windows: `%LOCALAPPDATA%\Programs\Pencil\...\mcp-server-windows-x64.exe` |

## 特性

### Context Image 整体导出

当 Pencil 中的节点设置了 `context` 属性为 `image` 或 `img` 时，该节点会被整体导出为 PNG 图片并直接插入 PPT，而不是逐子节点渲染。适用于复杂表格、图表等不适合拆解为原生 PPT 元素的内容。

在 Pencil 中设置方式：选中节点 → 设置 `context` 属性为 `image` 或 `img`。

### 页面顺序

导出的 PPT 页面顺序与 Pencil 画布上的视觉排列一致，按从上到下、从左到右排序。

## 工作原理

1. 通过 MCP 协议连接正在运行的 Pencil 桌面应用
2. 调用 `snapshot_layout` 获取 Pencil 引擎计算后的精确布局（坐标、尺寸）
3. 按画布位置（先 y 后 x）排序顶层帧，确保页面顺序与 Pencil 视觉顺序一致
4. 调用 `batch_get` 获取节点属性（类型、样式、文本内容等）
5. 调用 `export_nodes` 将 `icon_font` 节点导出为 4x 高清 PNG
6. 检测 `context` 为 `image`/`img` 的节点，整体导出为 2x PNG 图片
7. 合并布局和属性数据，渲染为原生 PPT 元素（context image 节点直接以图片插入）

## 依赖

- Python 3.10+
- [python-pptx](https://python-pptx.readthedocs.io/)
- [mcp](https://pypi.org/project/mcp/) (Model Context Protocol SDK)
- [Pencil](https://pencil.evolveui.com/) 桌面应用（需运行中）

## License

MIT
