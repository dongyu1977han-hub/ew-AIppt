"""Core SVG -> DrawingML dispatcher, group handling, and main entry point."""

from __future__ import annotations

import re
from copy import deepcopy
from pathlib import Path
from xml.etree import ElementTree as ET

from .drawingml_context import ConvertContext, ShapeResult
from .drawingml_utils import (
    SVG_NS, XLINK_NS,
    _extract_inheritable_styles, resolve_url_id, _f,
)
from .drawingml_styles import build_effect_xml
from .drawingml_elements import (
    convert_rect, convert_circle, convert_ellipse,
    convert_line, convert_path,
    convert_polygon, convert_polyline,
    convert_text, convert_image,
)


# ---------------------------------------------------------------------------
# Transform & layout helpers
# ---------------------------------------------------------------------------

import math as _math


def parse_transform(transform_str: str) -> tuple[float, float, float, float, float]:
    """Parse SVG transform string, extract translate, scale, and rotate.

    Supports: translate(), scale(), rotate(), matrix().
    Returns:
        (dx, dy, sx, sy, angle_deg) tuple.
    """
    if not transform_str:
        return 0.0, 0.0, 1.0, 1.0, 0.0

    dx, dy = 0.0, 0.0
    sx, sy = 1.0, 1.0
    angle_deg = 0.0

    # matrix(a, b, c, d, e, f) — most general, check first
    m = re.search(r'matrix\(\s*([-\d.eE+]+)[\s,]+([-\d.eE+]+)[\s,]+([-\d.eE+]+)[\s,]+([-\d.eE+]+)[\s,]+([-\d.eE+]+)[\s,]+([-\d.eE+]+)\s*\)', transform_str)
    if m:
        a, b, c, d = float(m.group(1)), float(m.group(2)), float(m.group(3)), float(m.group(4))
        dx, dy = float(m.group(5)), float(m.group(6))
        sx = _math.sqrt(a * a + b * b) if (a * a + b * b) > 1e-10 else 1.0
        sy = _math.sqrt(c * c + d * d) if (c * c + d * d) > 1e-10 else 1.0
        angle_deg = _math.degrees(_math.atan2(b, a))
        return dx, dy, sx, sy, angle_deg

    m = re.search(r'translate\(\s*([-\d.eE+]+)[\s,]+([-\d.eE+]+)\s*\)', transform_str)
    if m:
        dx = float(m.group(1))
        dy = float(m.group(2))

    m = re.search(r'scale\(\s*([-\d.eE+]+)(?:[\s,]+([-\d.eE+]+))?\s*\)', transform_str)
    if m:
        sx = float(m.group(1))
        sy = float(m.group(2)) if m.group(2) else sx

    m = re.search(r'rotate\(\s*([-\d.eE+]+)', transform_str)
    if m:
        angle_deg = float(m.group(1))

    return dx, dy, sx, sy, angle_deg


# ---------------------------------------------------------------------------
# Group handling
# ---------------------------------------------------------------------------

def convert_g(elem: ET.Element, ctx: ConvertContext) -> ShapeResult | None:
    """Convert SVG <g> to DrawingML group shape <p:grpSp>.

    Preserves group structure so elements can be selected and moved together
    in PowerPoint. Single-child groups are flattened to avoid unnecessary nesting.

    Uses identity coordinate mapping (chOff/chExt == off/ext) so child shapes
    keep their absolute slide coordinates unchanged.
    """
    transform = elem.get('transform', '')
    dx, dy, sx, sy, angle_deg = parse_transform(transform)

    filter_id = resolve_url_id(elem.get('filter', ''))
    style_overrides = _extract_inheritable_styles(elem)
    child_ctx = ctx.child(dx, dy, sx, sy, filter_id, style_overrides)

    child_results: list[ShapeResult] = []
    for child in elem:
        result = convert_element(child, child_ctx)
        if result:
            child_results.append(result)

    ctx.sync_from_child(child_ctx)

    if not child_results:
        return None

    # Single child: flatten only if the group has no effects/styles
    if len(child_results) == 1 and not filter_id and not style_overrides:
        return child_results[0]

    # Multiple children: wrap in <p:grpSp>
    min_x = min_y = float('inf')
    max_x = max_y = float('-inf')

    for child_result in child_results:
        bounds = child_result.bounds_emu
        if bounds is None:
            continue
        min_x = min(min_x, bounds[0])
        min_y = min(min_y, bounds[1])
        max_x = max(max_x, bounds[2])
        max_y = max(max_y, bounds[3])

    if min_x == float('inf'):
        return ShapeResult(xml='\n'.join(result.xml for result in child_results))

    group_x = int(min_x)
    group_y = int(min_y)
    group_w = max(int(max_x - min_x), 1)
    group_h = max(int(max_y - min_y), 1)

    shapes_xml = '\n'.join(result.xml for result in child_results)
    group_id = ctx.next_id()

    group_effect = ''
    if filter_id and filter_id in ctx.defs:
        group_effect = build_effect_xml(ctx.defs[filter_id])

    rot_emu = int(angle_deg * 60000)
    rot_attr = f' rot="{rot_emu}"' if rot_emu else ''

    return ShapeResult(xml=f'''<p:grpSp>
<p:nvGrpSpPr>
<p:cNvPr id="{group_id}" name="Group {group_id}"/>
<p:cNvGrpSpPr/>
<p:nvPr/>
</p:nvGrpSpPr>
<p:grpSpPr>
<a:xfrm{rot_attr}>
<a:off x="{group_x}" y="{group_y}"/>
<a:ext cx="{group_w}" cy="{group_h}"/>
<a:chOff x="{group_x}" y="{group_y}"/>
<a:chExt cx="{group_w}" cy="{group_h}"/>
</a:xfrm>
{group_effect}
</p:grpSpPr>
{shapes_xml}
</p:grpSp>''', bounds_emu=(group_x, group_y, group_x + group_w, group_y + group_h))


# ---------------------------------------------------------------------------
# <use> element handling
# ---------------------------------------------------------------------------

def convert_use(elem: ET.Element, ctx: ConvertContext) -> ShapeResult | None:
    """Convert SVG <use> by resolving the referenced element and converting it.

    Looks up href/xlink:href in ctx.defs, clones the element with x/y offset,
    and delegates to the appropriate converter.
    """
    href = elem.get('href') or elem.get(f'{{{XLINK_NS}}}href') or ''
    ref_id = href.lstrip('#')
    if not ref_id or ref_id not in ctx.defs:
        return None

    # Clone the referenced element
    ref_elem = ctx.defs[ref_id]
    clone = deepcopy(ref_elem)

    # Apply use element's x/y as translate offset
    use_x = _f(elem.get('x'))
    use_y = _f(elem.get('y'))
    if use_x or use_y:
        existing_transform = clone.get('transform', '')
        clone.set('transform', f'translate({use_x}, {use_y}) {existing_transform}'.strip())

    # Apply use element's dimensions (width/height) if the referenced element supports them
    use_w = elem.get('width')
    use_h = elem.get('height')
    if use_w and clone.tag.endswith('}') and 'image' in clone.tag:
        clone.set('width', use_w)
    if use_h and clone.tag.endswith('}') and 'image' in clone.tag:
        clone.set('height', use_h)

    return convert_element(clone, ctx)


# ---------------------------------------------------------------------------
# Defs collection & element dispatch
# ---------------------------------------------------------------------------

_NON_VISUAL_TAGS = frozenset(('defs', 'title', 'desc', 'metadata', 'style'))

_CONVERTERS = {
    'rect': convert_rect,
    'circle': convert_circle,
    'ellipse': convert_ellipse,
    'line': convert_line,
    'path': convert_path,
    'polygon': convert_polygon,
    'polyline': convert_polyline,
    'text': convert_text,
    'image': convert_image,
    'g': convert_g,
    'use': convert_use,
}


def collect_defs(root: ET.Element) -> dict[str, ET.Element]:
    """Collect all <defs> children into an {id: element} dictionary."""
    defs: dict[str, ET.Element] = {}
    seen_ids: set[str] = set()
    for defs_elem in root.iter(f'{{{SVG_NS}}}defs'):
        for child in defs_elem:
            elem_id = child.get('id')
            if elem_id and elem_id not in seen_ids:
                defs[elem_id] = child
                seen_ids.add(elem_id)
    # Also check for defs without namespace (avoid duplicates via seen_ids)
    for defs_elem in root.iter('defs'):
        for child in defs_elem:
            elem_id = child.get('id')
            if elem_id and elem_id not in seen_ids:
                defs[elem_id] = child
                seen_ids.add(elem_id)
    return defs


def convert_element(elem: ET.Element, ctx: ConvertContext) -> ShapeResult | None:
    """Dispatch an SVG element to the appropriate converter."""
    tag = elem.tag.replace(f'{{{SVG_NS}}}', '')

    converter = _CONVERTERS.get(tag)
    if converter:
        try:
            return converter(elem, ctx)
        except Exception as e:
            print(f'  Warning: Failed to convert <{tag}>: {e}')
            return None

    if tag in _NON_VISUAL_TAGS:
        return None

    return None


def convert_svg_to_slide_shapes(
    svg_path: Path,
    slide_num: int = 1,
    verbose: bool = False,
) -> tuple[str, dict[str, bytes], list[dict[str, str]]]:
    """Convert an SVG file to a complete DrawingML slide XML.

    Args:
        svg_path: Path to the SVG file.
        slide_num: Slide number (for naming).
        verbose: Print progress info.

    Returns:
        (slide_xml, media_files, rel_entries) where:
        - slide_xml: Complete slide XML string.
        - media_files: Dict of {filename: bytes} for media to write.
        - rel_entries: List of relationship entries to add.
    """
    tree = ET.parse(str(svg_path))
    root = tree.getroot()

    defs = collect_defs(root)
    ctx = ConvertContext(defs=defs, slide_num=slide_num, svg_dir=Path(svg_path).parent)

    shapes: list[str] = []
    converted = 0
    skipped = 0

    for child in root:
        tag = child.tag.replace(f'{{{SVG_NS}}}', '')
        if tag == 'defs':
            continue
        result = convert_element(child, ctx)
        if result:
            shapes.append(result.xml)
            converted += 1
        else:
            if tag not in _NON_VISUAL_TAGS:
                skipped += 1

    if verbose:
        print(f'  Converted {converted} elements, skipped {skipped}')

    shapes_xml = '\n'.join(shapes)

    slide_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:spTree>
<p:nvGrpSpPr>
<p:cNvPr id="1" name=""/>
<p:cNvGrpSpPr/><p:nvPr/>
</p:nvGrpSpPr>
<p:grpSpPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
</p:grpSpPr>
{shapes_xml}
</p:spTree>
</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''

    return slide_xml, ctx.media_files, ctx.rel_entries
