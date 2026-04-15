#!/usr/bin/env python3
"""
pen2pptx — 将 Pencil .pen 文件转换为 PowerPoint .pptx

通过 Pencil MCP server 获取精确布局数据，渲染为原生 PPT 元素。
需要 Pencil 桌面应用正在运行。

用法:
    python pen2pptx.py input.pen [-o output.pptx] [--font-scale 0.73]
"""

from __future__ import annotations

import argparse
import asyncio
import json
import logging
import os
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

logger = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════
# 常量
# ═══════════════════════════════════════════════════════════════

PX_TO_EMU = 9525  # 1 像素 = 9525 EMU
DEFAULT_FONT_SCALE = 0.73  # Pencil px → PPT pt 的默认系数

# Pencil MCP server 路径（根据实际安装位置修改）
DEFAULT_PENCIL_CMD = str(
    Path(os.environ.get("LOCALAPPDATA", ""))
    / "Programs" / "Pencil" / "resources"
    / "app.asar.unpacked" / "out" / "mcp-server-windows-x64.exe"
)

# ═══════════════════════════════════════════════════════════════
# 数据模型
# ═══════════════════════════════════════════════════════════════


@dataclass
class LayoutNode:
    """合并了属性和布局信息的节点"""
    id: str = ""
    node_type: str = ""
    name: str = ""
    x: float = 0.0
    y: float = 0.0
    width: float = 0.0
    height: float = 0.0
    fill: str | None = None
    content: str = ""
    font_family: str | None = None
    font_size: float | None = None
    font_weight: str = "normal"
    font_style: str = "normal"
    text_align: str = "left"
    line_height: float = 1.2
    letter_spacing: float = 0.0
    corner_radius: float = 0.0
    opacity: float = 1.0
    stroke_color: str | None = None
    stroke_width: float = 0.0
    icon_font_name: str = ""
    icon_font_family: str = ""
    icon_image_path: str = ""
    children: list[LayoutNode] = field(default_factory=list)
    raw_props: dict = field(default_factory=dict)


@dataclass
class PageData:
    """一个页面（幻灯片）的数据"""
    id: str = ""
    name: str = ""
    width: float = 960.0
    height: float = 540.0
    nodes: list[LayoutNode] = field(default_factory=list)


# ═══════════════════════════════════════════════════════════════
# MCP 客户端 — 获取节点属性 + 计算后布局 + 导出图标
# ═══════════════════════════════════════════════════════════════

class PencilMcpClient:

    def __init__(self, pencil_cmd: str = DEFAULT_PENCIL_CMD):
        self.pencil_cmd = pencil_cmd

    async def fetch_pages(self, pen_file: str) -> list[PageData]:
        server_params = StdioServerParameters(
            command=self.pencil_cmd, args=["--app", "desktop"],
        )
        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                return await self._fetch_impl(session, pen_file)

    async def _fetch_impl(self, session: ClientSession, pen_file: str) -> list[PageData]:
        top_layout = await self._call(session, "snapshot_layout", {
            "filePath": pen_file, "maxDepth": 0,
        })

        pages: list[PageData] = []
        all_icons: list[LayoutNode] = []

        for frame in top_layout:
            pid = frame["id"]

            layout = await self._call(session, "snapshot_layout", {
                "filePath": pen_file, "parentId": pid, "maxDepth": 20,
            })
            props_list = await self._call(session, "batch_get", {
                "filePath": pen_file, "parentId": pid,
                "readDepth": 10, "resolveInstances": True,
            })
            props_map = self._build_props_map(props_list)

            page_info = await self._call(session, "batch_get", {
                "filePath": pen_file, "nodeIds": [pid], "readDepth": 1,
            })
            page_name = page_info[0].get("name", "") if page_info else ""

            nodes = [self._merge(c, props_map) for c in layout.get("children", []) if isinstance(c, dict)]

            pages.append(PageData(
                id=pid, name=page_name,
                width=float(frame.get("width", 960)),
                height=float(frame.get("height", 540)),
                nodes=nodes,
            ))
            self._collect_icons(nodes, all_icons)

        if all_icons:
            await self._export_icons(session, pen_file, all_icons)

        return pages

    # --- 图标导出 ---

    def _collect_icons(self, nodes: list[LayoutNode], out: list[LayoutNode]) -> None:
        for n in nodes:
            if n.node_type == "icon_font" and n.id:
                out.append(n)
            self._collect_icons(n.children, out)

    async def _export_icons(self, session: ClientSession, pen_file: str, icons: list[LayoutNode]) -> None:
        icon_dir = tempfile.mkdtemp(prefix="pencil_icons_")
        try:
            raw = await session.call_tool("export_nodes", {
                "filePath": pen_file, "nodeIds": [n.id for n in icons],
                "outputDir": icon_dir, "format": "png", "scale": 4,
            })
            # 解析返回
            for item in raw.content:
                if hasattr(item, "text"):
                    try:
                        json.loads(item.text)
                    except (json.JSONDecodeError, ValueError):
                        pass
                    break
            # 从目录中匹配文件
            id_map = {n.id: n for n in icons}
            for fname in os.listdir(icon_dir):
                base = os.path.splitext(fname)[0]
                if base in id_map:
                    id_map[base].icon_image_path = os.path.join(icon_dir, fname)
        except Exception as e:
            logger.warning("图标导出失败: %s", e)

    # --- 数据构建 ---

    def _build_props_map(self, nodes) -> dict[str, dict]:
        result: dict[str, dict] = {}
        self._collect_props(nodes if isinstance(nodes, list) else [nodes], result)
        return result

    def _collect_props(self, nodes: list, result: dict[str, dict]) -> None:
        for node in nodes:
            if not isinstance(node, dict):
                continue
            nid = node.get("id", "")
            if nid:
                result[nid] = node
            children = node.get("children", [])
            if isinstance(children, list):
                self._collect_props(children, result)

    def _merge(self, layout: dict, props_map: dict[str, dict]) -> LayoutNode:
        nid = layout.get("id", "")
        p = props_map.get(nid, {})

        stroke_color, stroke_width = None, 0.0
        sr = p.get("stroke")
        if isinstance(sr, dict):
            fv = sr.get("fill")
            if isinstance(fv, str):
                stroke_color = fv
            th = sr.get("thickness", 0)
            if isinstance(th, (int, float)):
                stroke_width = float(th)
        elif isinstance(sr, str):
            stroke_color = sr

        node = LayoutNode(
            id=nid, node_type=p.get("type", ""), name=p.get("name", ""),
            x=float(layout.get("x", 0)), y=float(layout.get("y", 0)),
            width=float(layout.get("width", 0)), height=float(layout.get("height", 0)),
            fill=p.get("fill"), content=p.get("content", ""),
            font_family=p.get("fontFamily"), font_size=_sf(p.get("fontSize")),
            font_weight=str(p.get("fontWeight", "normal")),
            font_style=str(p.get("fontStyle", "normal")),
            text_align=p.get("textAlign", "left"),
            line_height=_sf(p.get("lineHeight"), 1.2),
            letter_spacing=_sf(p.get("letterSpacing"), 0.0),
            corner_radius=_sf(p.get("cornerRadius"), 0.0),
            opacity=_sf(p.get("opacity"), 1.0),
            stroke_color=stroke_color, stroke_width=stroke_width,
            icon_font_name=p.get("iconFontName", ""),
            icon_font_family=p.get("iconFontFamily", ""),
            raw_props=p,
        )
        for cl in layout.get("children", []):
            if isinstance(cl, dict):
                node.children.append(self._merge(cl, props_map))
        return node

    async def _call(self, session: ClientSession, tool: str, args: dict) -> Any:
        result = await session.call_tool(tool, args)
        for item in result.content:
            if hasattr(item, "text"):
                return json.loads(item.text)
        return {}


# ═══════════════════════════════════════════════════════════════
# PPT 渲染器
# ═══════════════════════════════════════════════════════════════

def render_pages(pages: list[PageData], output: Path, font_scale: float) -> None:
    prs = Presentation()
    for page in pages:
        prs.slide_width = Emu(int(page.width * PX_TO_EMU))
        prs.slide_height = Emu(int(page.height * PX_TO_EMU))
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for node in page.nodes:
            _render(slide, node, 0.0, 0.0, font_scale)
    prs.save(str(output))


def _render(slide, n: LayoutNode, px: float, py: float, fs: float) -> None:
    ax, ay = px + n.x, py + n.y
    t = n.node_type
    if t == "text":
        _text(slide, n, ax, ay, fs)
    elif t == "rectangle":
        _rect(slide, n, ax, ay)
    elif t == "ellipse":
        _ellipse(slide, n, ax, ay)
    elif t == "line":
        _line(slide, n, ax, ay)
    elif t in ("frame", "group"):
        _frame(slide, n, ax, ay, fs)
    elif t == "icon_font":
        _icon(slide, n, ax, ay)


def _text(slide, n: LayoutNode, ax: float, ay: float, fs: float) -> None:
    if not n.content:
        return
    left = Emu(int(ax * PX_TO_EMU))
    top = Emu(int(ay * PX_TO_EMU))
    w = Emu(int(n.width * PX_TO_EMU)) if n.width > 0 else Emu(200 * PX_TO_EMU)
    h = Emu(int(n.height * PX_TO_EMU)) if n.height > 0 else Emu(30 * PX_TO_EMU)

    tf = slide.shapes.add_textbox(left, top, w, h).text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)

    align_map = {"center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}

    for i, line in enumerate(n.content.split("\n")):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align_map.get(n.text_align, PP_ALIGN.LEFT)
        run = para.add_run()
        run.text = line
        f = run.font
        if n.font_family:
            f.name = n.font_family
        if n.font_size and n.font_size > 0:
            f.size = Pt(n.font_size * fs)
        if n.fill:
            f.color.rgb = _rgb(n.fill)
        if n.font_weight in ("bold", "600", "700", "800", "900"):
            f.bold = True
        if n.font_style == "italic":
            f.italic = True
        if n.line_height and n.line_height != 1.0:
            para.line_spacing = n.line_height


def _rect(slide, n: LayoutNode, ax: float, ay: float) -> None:
    l, t, w, h = _box(ax, ay, n.width, n.height)
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if n.corner_radius > 0 else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(shape_type, l, t, w, h)
    _fill(s, n); _stroke(s, n)


def _ellipse(slide, n: LayoutNode, ax: float, ay: float) -> None:
    l, t, w, h = _box(ax, ay, n.width, n.height)
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, l, t, w, h)
    _fill(s, n); _stroke(s, n)


def _line(slide, n: LayoutNode, ax: float, ay: float) -> None:
    l, t, w, h = _box(ax, ay, n.width, max(n.height, 0.5))
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    _fill(s, n); s.line.fill.background()


def _frame(slide, n: LayoutNode, ax: float, ay: float, fs: float) -> None:
    if n.fill or n.stroke_color:
        l, t, w, h = _box(ax, ay, n.width, n.height)
        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if n.corner_radius > 0 else MSO_SHAPE.RECTANGLE
        s = slide.shapes.add_shape(shape_type, l, t, w, h)
        _fill(s, n); _stroke(s, n)
    for c in n.children:
        _render(slide, c, ax, ay, fs)


def _icon(slide, n: LayoutNode, ax: float, ay: float) -> None:
    if not n.icon_image_path or not os.path.exists(n.icon_image_path):
        return
    l, t, w, h = _box(ax, ay, n.width, n.height)
    slide.shapes.add_picture(n.icon_image_path, l, t, w, h)


# --- 工具函数 ---

def _box(ax, ay, w, h):
    return Emu(int(ax * PX_TO_EMU)), Emu(int(ay * PX_TO_EMU)), Emu(int(w * PX_TO_EMU)), Emu(int(h * PX_TO_EMU))

def _fill(s, n: LayoutNode):
    if n.fill:
        s.fill.solid(); s.fill.fore_color.rgb = _rgb(n.fill)
    else:
        s.fill.background()

def _stroke(s, n: LayoutNode):
    if n.stroke_color and n.stroke_width > 0:
        s.line.color.rgb = _rgb(n.stroke_color); s.line.width = Pt(n.stroke_width)
    else:
        s.line.fill.background()

def _rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)) if len(h) == 6 else RGBColor(0, 0, 0)

def _sf(v, d: float = 0.0) -> float:
    if v is None: return d
    try: return float(v)
    except (ValueError, TypeError): return d


# ═══════════════════════════════════════════════════════════════
# CLI 入口
# ═══════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(
        prog="pencil2pptx",
        description="将 Pencil .pen 文件转换为 PowerPoint .pptx（需要 Pencil 应用运行）",
    )
    ap.add_argument("input", help="输入 .pen 文件路径")
    ap.add_argument("-o", "--output", default=None, help="输出 .pptx 路径（默认同名 .pptx）")
    ap.add_argument("--font-scale", type=float, default=DEFAULT_FONT_SCALE,
                     help=f"字体缩放系数，Pencil px → PPT pt（默认 {DEFAULT_FONT_SCALE}）")
    ap.add_argument("--pencil-cmd", default=DEFAULT_PENCIL_CMD,
                     help="Pencil MCP server 可执行文件路径")
    parsed = ap.parse_args()

    inp = Path(parsed.input).resolve()
    if not inp.exists():
        print(f"错误: 文件不存在: {inp}", file=sys.stderr); sys.exit(1)

    out = Path(parsed.output) if parsed.output else inp.with_suffix(".pptx")

    try:
        print(f"连接 Pencil MCP server...")
        client = PencilMcpClient(parsed.pencil_cmd)
        pages = asyncio.run(client.fetch_pages(str(inp)))
        total_nodes = sum(_count(p.nodes) for p in pages)
        print(f"获取到 {len(pages)} 页, {total_nodes} 个节点")

        render_pages(pages, out, parsed.font_scale)
        print(f"完成: {out}")
    except Exception as e:
        print(f"失败: {e}", file=sys.stderr); sys.exit(1)


def _count(nodes) -> int:
    return sum(1 + _count(n.children) for n in nodes)


if __name__ == "__main__":
    main()
