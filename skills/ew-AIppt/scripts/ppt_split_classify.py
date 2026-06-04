#!/usr/bin/env python3
"""
ppt_split_classify.py — PPT 拆页 + 四维标签自动分类

将多页 PPTX 拆分为单页模板文件，并对每页进行四维标签智能识别：
  维度一：逻辑拓扑标签 (logic)
  维度二：内容领域标签 (domain)
  维度三：数量规模标签 (count)
  维度四：图形类型标签 (chart-type)

用法:
    python ppt_split_classify.py <input.pptx> [--output-dir <dir>] [--name <prefix>]

输出:
    <output_dir>/
      slide_001.pptx
      slide_002.pptx
      ...
      manifest.json
"""

import argparse
import io
import json
import math
import os
import re
import sys
from collections import Counter
from copy import deepcopy
from pathlib import Path

# Windows 控制台 UTF-8 输出
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ─────────────────────────────────────────────────────────
# 配置
# ─────────────────────────────────────────────────────────

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent.parent.parent  # ew-AIppt root
TAG_SCHEMA_PATH = PROJECT_ROOT / "skills" / "ew-AIppt" / "templates" / "_tag_schema.json"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "template-library"

# ─────────────────────────────────────────────────────────
# 标签 Schema 加载
# ─────────────────────────────────────────────────────────

def load_tag_schema(path: Path = TAG_SCHEMA_PATH) -> dict:
    """加载 _tag_schema.json，返回 {dim_id: [tag_id, ...]} 映射"""
    with open(path, "r", encoding="utf-8") as f:
        schema = json.load(f)
    result = {}
    for dim in schema["dimensions"]:
        result[dim["id"]] = [t["id"] for t in dim["tags"]]
    return result


# ─────────────────────────────────────────────────────────
# 工具函数
# ─────────────────────────────────────────────────────────

def get_shape_bbox(shape):
    """返回 (left, top, width, height, cx, cy) 单位 EMU"""
    left = shape.left if shape.left else 0
    top = shape.top if shape.top else 0
    width = shape.width if shape.width else 0
    height = shape.height if shape.height else 0
    cx = left + width // 2
    cy = top + height // 2
    return left, top, width, height, cx, cy


def extract_text(shape_or_slide) -> str:
    """递归提取 shape 或 slide 中的所有文本"""
    texts = []
    # 如果传入的是 Slide 对象，遍历其 shapes
    if hasattr(shape_or_slide, "shapes") and not hasattr(shape_or_slide, "has_text_frame"):
        for s in shape_or_slide.shapes:
            texts.append(extract_text(s))
        return "\n".join(t for t in texts if t)
    shape = shape_or_slide
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            t = para.text.strip()
            if t:
                texts.append(t)
    if hasattr(shape, "table"):
        for row in shape.table.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t:
                    texts.append(t)
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            texts.extend(extract_text(child).split("\n"))
    return "\n".join(texts)


def is_title_shape(shape, slide_height: int) -> bool:
    """判断是否为标题/页码等非内容 shape"""
    _, top, _, height, _, _ = get_shape_bbox(shape)
    # 页面顶部 15% 区域的文本框通常是标题
    if top < slide_height * 0.15 and height < slide_height * 0.12:
        return True
    # 页面底部 8% 区域通常是页码/脚注
    if top > slide_height * 0.92:
        return True
    # 纯装饰性形状（无文本、非图片）
    if not shape.has_text_frame and shape.shape_type not in (
        MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.PLACEHOLDER
    ):
        # 保留表格、图表、连接线
        if hasattr(shape, "table"):
            return False
        if hasattr(shape, "has_chart") and shape.has_chart:
            return False
        if hasattr(shape, "connector") and shape.connector:
            return False
        return True
    return False


def get_content_shapes(slide, slide_width: int, slide_height: int) -> list:
    """获取页面中真正承载内容的 shape 列表（排除标题、页码、装饰）"""
    content_shapes = []
    for shape in slide.shapes:
        if is_title_shape(shape, slide_height):
            continue
        text = extract_text(shape)
        has_image = shape.shape_type == MSO_SHAPE_TYPE.PICTURE
        has_table = hasattr(shape, "table")
        has_chart = shape.has_chart if hasattr(shape, "has_chart") else False
        # 有文本、图片、表格、图表的内容 shape
        if text or has_image or has_table or has_chart:
            content_shapes.append(shape)
    return content_shapes


# ─────────────────────────────────────────────────────────
# 维度一：逻辑拓扑标签 (logic)
# ─────────────────────────────────────────────────────────

LOGIC_KEYWORDS = {
    "sequential": ["→", "➡", ">>", "->", "步骤", "step", "phase", "阶段"],
    "cycle": ["循环", "闭环", "PDCA", "cycle", "loop", "迭代"],
    "pyramid": ["棱锥", "金字塔", "pyramid", "层级递进"],
    "matrix": ["矩阵", "matrix", "象限", "quadrant"],
    "contrast": ["对比", "vs", "versus", "比较", "差异", "优劣"],
    "star": ["中心", "辐射", "hub", "核心", "围绕"],
    "hierarchy": ["架构", "层级", "树", "tree", "组织"],
}

LOGIC_TAG_ORDER = ["parallel", "sequential", "star", "contrast", "hierarchy",
                   "matrix", "cycle", "pyramid", "other"]


def classify_logic(slide, content_shapes: list, slide_width: int, slide_height: int) -> str:
    """基于空间布局和文本关键词判断逻辑拓扑类型"""
    n = len(content_shapes)
    if n == 0:
        return "other"
    if n == 1:
        return "parallel"

    # ── 提取所有文本用于关键词匹配 ──
    all_text = extract_text(slide).lower()

    # ── 基于关键词的快速判断 ──
    for tag_id, keywords in LOGIC_KEYWORDS.items():
        for kw in keywords:
            if kw.lower() in all_text:
                return tag_id

    # ── 基于空间布局的几何分析 ──
    bboxes = [get_shape_bbox(s) for s in content_shapes]
    cxs = [b[4] for b in bboxes]
    cys = [b[5] for b in bboxes]

    # 检测 cycle（循环）：形状呈环形分布
    if _is_cycle_layout(cxs, cys, slide_width, slide_height):
        return "cycle"

    # 检测 pyramid（棱锥）：从上到下每行形状数量递增
    if _is_pyramid_layout(bboxes, slide_height):
        return "pyramid"

    # 检测 matrix（矩阵）：均匀网格分布
    if _is_matrix_layout(bboxes, n):
        return "matrix"

    # 检测 contrast（对比）：左右或上下两大组
    if _is_contrast_layout(bboxes, slide_width, slide_height):
        return "contrast"

    # 检测 star（星型）：一个中心 + 周围环绕
    if _is_star_layout(bboxes, slide_width, slide_height):
        return "star"

    # 检测 hierarchy（层级）：树状结构，上少下多
    if _is_hierarchy_layout(bboxes, slide_height):
        return "hierarchy"

    # 检测 sequential（递进）：水平或垂直的线性排列 + 箭头
    if _is_sequential_layout(bboxes, all_text):
        return "sequential"

    # 默认：并列
    return "parallel"


def _is_cycle_layout(cxs, cys, sw, sh) -> bool:
    """形状中心点是否呈环形分布"""
    n = len(cxs)
    if n < 3:
        return False
    # 计算中心点的几何中心
    avg_cx = sum(cxs) / n
    avg_cy = sum(cys) / n
    # 计算各点到中心的距离
    dists = [math.sqrt((cx - avg_cx) ** 2 + (cy - avg_cy) ** 2) for cx, cy in zip(cxs, cys)]
    if not dists:
        return False
    avg_dist = sum(dists) / len(dists)
    if avg_dist == 0:
        return False
    # 如果所有点到中心的距离标准差 < 平均距离的 35%，认为是环形
    variance = sum((d - avg_dist) ** 2 for d in dists) / len(dists)
    std = math.sqrt(variance)
    return std / avg_dist < 0.35


def _is_pyramid_layout(bboxes, sh) -> bool:
    """从上到下每行形状数量递增"""
    n = len(bboxes)
    if n < 3:
        return False
    # 按 y 坐标分层（容差 = 页高 * 0.08）
    layer_threshold = sh * 0.08
    sorted_bboxes = sorted(bboxes, key=lambda b: b[5])  # 按 cy 排序
    layers = []
    current_layer = [sorted_bboxes[0]]
    for bb in sorted_bboxes[1:]:
        if bb[5] - current_layer[-1][5] < layer_threshold:
            current_layer.append(bb)
        else:
            layers.append(current_layer)
            current_layer = [bb]
    layers.append(current_layer)
    if len(layers) < 3:
        return False
    counts = [len(l) for l in layers]
    # 检查是否单调递增
    return all(counts[i] < counts[i + 1] for i in range(len(counts) - 1))


def _is_matrix_layout(bboxes, n) -> bool:
    """形状是否呈网格均匀分布"""
    if n < 4:
        return False
    cxs = sorted(set(round(b[4] / 50000) * 50000 for b in bboxes))
    cys = sorted(set(round(b[5] / 50000) * 50000 for b in bboxes))
    # 至少 2 行 × 2 列
    if len(cxs) >= 2 and len(cys) >= 2:
        # 检查是否每个交叉点都有形状
        expected = len(cxs) * len(cys)
        if abs(n - expected) <= 1:
            return True
    return False


def _is_contrast_layout(bboxes, sw, sh) -> bool:
    """形状是否分为左右或上下两大组"""
    n = len(bboxes)
    if n < 2:
        return False
    mid_x = sw // 2
    left_count = sum(1 for b in bboxes if b[4] < mid_x)
    right_count = n - left_count
    # 左右各半（允许 1 个误差）
    if abs(left_count - right_count) <= 1 and min(left_count, right_count) >= 1:
        # 检查左侧形状和右侧形状是否有明显间距
        left_shapes = [b for b in bboxes if b[4] < mid_x]
        right_shapes = [b for b in bboxes if b[4] >= mid_x]
        left_max_right = max(b[0] + b[2] for b in left_shapes)
        right_min_left = min(b[0] for b in right_shapes)
        gap = right_min_left - left_max_right
        if gap > sw * 0.05:
            return True
    return False


def _is_star_layout(bboxes, sw, sh) -> bool:
    """是否一个中心 + 周围环绕"""
    n = len(bboxes)
    if n < 4:
        return False
    # 找到离页面中心最近的形状作为"中心"
    page_cx, page_cy = sw // 2, sh // 2
    dists_to_center = [
        (i, math.sqrt((b[4] - page_cx) ** 2 + (b[5] - page_cy) ** 2))
        for i, b in enumerate(bboxes)
    ]
    dists_to_center.sort(key=lambda x: x[1])
    center_idx = dists_to_center[0][0]
    center_bbox = bboxes[center_idx]
    # 其余形状到中心形状的距离
    other_dists = [
        math.sqrt((b[4] - center_bbox[4]) ** 2 + (b[5] - center_bbox[5]) ** 2)
        for i, b in enumerate(bboxes) if i != center_idx
    ]
    if not other_dists:
        return False
    avg_dist = sum(other_dists) / len(other_dists)
    if avg_dist == 0:
        return False
    variance = sum((d - avg_dist) ** 2 for d in other_dists) / len(other_dists)
    std = math.sqrt(variance)
    # 距离标准差小（均匀辐射）且中心形状相对较大
    center_area = center_bbox[2] * center_bbox[3]
    avg_other_area = sum(bboxes[i][2] * bboxes[i][3] for i in range(n) if i != center_idx) / (n - 1)
    return std / avg_dist < 0.4 and center_area > avg_other_area * 0.8


def _is_hierarchy_layout(bboxes, sh) -> bool:
    """是否树状层级结构（上少下多）"""
    n = len(bboxes)
    if n < 3:
        return False
    # 按 y 分层
    layer_threshold = sh * 0.1
    sorted_bboxes = sorted(bboxes, key=lambda b: b[5])
    layers = []
    current_layer = [sorted_bboxes[0]]
    for bb in sorted_bboxes[1:]:
        if bb[5] - current_layer[-1][5] < layer_threshold:
            current_layer.append(bb)
        else:
            layers.append(current_layer)
            current_layer = [bb]
    layers.append(current_layer)
    if len(layers) < 2:
        return False
    counts = [len(l) for l in layers]
    # 第一层最少，后面递增（但不要求严格单调）
    if counts[0] == 1 and counts[-1] > 1:
        return True
    return False


def _is_sequential_layout(bboxes, all_text: str) -> bool:
    """是否线性递进排列"""
    n = len(bboxes)
    if n < 2:
        return False
    # 检查文本中的箭头或编号
    has_arrow = any(arrow in all_text for arrow in ["→", "➡", ">>", "->"])
    has_numbering = bool(re.search(r'[\[（(]?\s*[1-9]\s*[\]）).]', all_text))
    if not (has_arrow or has_numbering):
        return False
    # 检查空间排列是否基本线性
    cxs = [b[4] for b in bboxes]
    cys = [b[5] for b in bboxes]
    x_range = max(cxs) - min(cxs) if cxs else 0
    y_range = max(cys) - min(cys) if cys else 0
    # 水平排列或垂直排列
    return x_range > y_range * 1.5 or y_range > x_range * 1.5


# ─────────────────────────────────────────────────────────
# 维度二：内容领域标签 (domain)
# ─────────────────────────────────────────────────────────

DOMAIN_KEYWORDS = {
    "strategy": [
        "战略", "规划", "目标", "愿景", "使命", "方向", "布局", "蓝图",
        "strategy", "vision", "mission", "goal", "roadmap", "plan",
        "五年", "三年", "年度计划", "顶层设计"
    ],
    "current-state": [
        "现状", "问题", "挑战", "痛点", "分析", "诊断", "评估", "差距",
        "current", "gap", "challenge", "issue", "problem", "analysis",
        "瓶颈", "困境", "不足", "短板", "swot"
    ],
    "architecture": [
        "架构", "系统", "平台", "框架", "体系", "整体", "全景",
        "architecture", "system", "platform", "framework",
        "顶层设计", "总体架构", "技术架构", "应用架构", "数据架构"
    ],
    "technology": [
        "技术", "开发", "数据", "算法", "模型", "AI", "人工智能",
        "technology", "development", "data", "algorithm", "model",
        "数字化", "智能化", "云计算", "大数据", "物联网", "区块链",
        "微服务", "中台", "API", "数据库"
    ],
    "organization": [
        "组织", "团队", "人员", "部门", "岗位", "职责", "人才",
        "organization", "team", "personnel", "department", "talent",
        "人力资源", "HR", "管理层", "治理结构", "委员会"
    ],
    "operation": [
        "运营", "流程", "管理", "绩效", "效率", "质量", "成本",
        "operation", "process", "management", "performance", "efficiency",
        "KPI", "OKR", "PDCA", "精益", "改善", "优化", "管控"
    ],
}

DOMAIN_TAG_ORDER = ["strategy", "current-state", "architecture", "technology",
                    "organization", "operation", "other"]


def classify_domain(slide) -> str:
    """基于文本关键词判断内容领域"""
    all_text = extract_text(slide).lower()
    if not all_text.strip():
        return "other"

    scores = {}
    for tag_id, keywords in DOMAIN_KEYWORDS.items():
        score = 0
        for kw in keywords:
            if kw.lower() in all_text:
                score += 1
        if score > 0:
            scores[tag_id] = score

    if not scores:
        return "other"

    # 返回得分最高的
    return max(scores, key=scores.get)


# ─────────────────────────────────────────────────────────
# 维度三：数量规模标签 (count)
# ─────────────────────────────────────────────────────────

COUNT_TAG_MAP = {
    1: "count-1", 2: "count-2", 3: "count-3", 4: "count-4",
    5: "count-5", 6: "count-6",
}
COUNT_TAG_ORDER = ["count-1", "count-2", "count-3", "count-4",
                   "count-5", "count-6", "count-7-plus"]


def classify_count(content_shapes: list) -> str:
    """基于内容 shape 数量判断规模"""
    n = len(content_shapes)
    if n == 0:
        return "count-1"
    if n >= 7:
        return "count-7-plus"
    return COUNT_TAG_MAP.get(n, "count-7-plus")


# ─────────────────────────────────────────────────────────
# 维度四：图形类型标签 (chart-type)
# ─────────────────────────────────────────────────────────

CHART_TYPE_TAG_ORDER = [
    "timeline", "house-chart", "logic-steps", "value-tree",
    "business-arch", "app-arch", "data-arch", "tech-arch",
    "process-value-chain", "scenario-map", "org-structure",
    "gantt", "kpi-system", "implementation-phase", "ecosystem",
    "table", "chart", "other"
]

CHART_TYPE_KEYWORDS = {
    "timeline": ["时间", "timeline", "里程碑", "milestone", "历程", "沿革", "年表"],
    "gantt": ["甘特", "gantt", "进度", "排期", "schedule", "wbs"],
    "org-structure": ["组织", "org", "部门", "department", "汇报", "架构图"],
    "kpi-system": ["指标", "kpi", "绩效", "metric", "dashboard", "仪表盘", "考核"],
    "implementation-phase": ["实施", "落地", "阶段", "phase", "步骤一", "步骤二"],
    "business-arch": ["业务架构", "business architecture"],
    "app-arch": ["应用架构", "application architecture", "系统架构"],
    "data-arch": ["数据架构", "data architecture", "数据模型"],
    "tech-arch": ["技术架构", "technical architecture", "技术栈", "tech stack"],
    "process-value-chain": ["价值链", "value chain", "流程", "process chain"],
    "house-chart": ["屋型", "house", "质量屋", "qfd"],
    "logic-steps": ["步骤", "step", "方法", "路径", "路径图"],
    "value-tree": ["价值树", "value tree", "分解", "拆解", "wbs"],
    "scenario-map": ["场景", "scenario", "用例", "use case"],
    "ecosystem": ["生态", "ecosystem", "联盟", "合作"],
    "table": ["表格", "table", "明细", "清单", "列表", "对照表", "一览表", "统计表"],
    "chart": ["图表", "chart", "柱状图", "折线图", "饼图", "bar chart", "line chart", "pie chart", "趋势图", "数据图"],
}


def classify_chart_type(slide, content_shapes: list, slide_width: int, slide_height: int) -> str:
    """基于文本关键词和视觉模式判断图形类型"""
    all_text = extract_text(slide).lower()

    # ── shape 类型直接检测（最可靠） ──
    has_table = any(hasattr(s, "table") for s in content_shapes)
    has_chart = any(hasattr(s, "has_chart") and s.has_chart for s in content_shapes)
    if has_table and has_chart:
        # 同时有表格和图表时，看哪个面积更大
        table_area = sum(s.width * s.height for s in content_shapes if hasattr(s, "table"))
        chart_area = sum(s.width * s.height for s in content_shapes if hasattr(s, "has_chart") and s.has_chart)
        return "table" if table_area >= chart_area else "chart"
    if has_table:
        return "table"
    if has_chart:
        return "chart"

    # ── 关键词匹配 ──
    for tag_id, keywords in CHART_TYPE_KEYWORDS.items():
        for kw in keywords:
            if kw.lower() in all_text:
                return tag_id

    # ── 视觉模式检测 ──
    n = len(content_shapes)
    if n == 0:
        return "other"

    bboxes = [get_shape_bbox(s) for s in content_shapes]

    # timeline：水平排列 + 箭头/连接线
    if _detect_timeline(content_shapes, bboxes, all_text):
        return "timeline"

    # org-structure：树状 + 有连线
    if _detect_org_structure(content_shapes, bboxes):
        return "org-structure"

    # logic-steps：编号步骤
    if _detect_logic_steps(all_text, n):
        return "logic-steps"

    # gantt：长条状水平矩形
    if _detect_gantt(bboxes, slide_width):
        return "gantt"

    # kpi-system：数字密集 + 大字号
    if _detect_kpi(content_shapes, all_text):
        return "kpi-system"

    # implementation-phase：分阶段区块
    if _detect_phases(bboxes, slide_width, all_text):
        return "implementation-phase"

    return "other"


def _detect_timeline(shapes, bboxes, text: str) -> bool:
    """检测时间轴布局：水平排列的节点 + 连接线"""
    n = len(bboxes)
    if n < 3:
        return False
    # 检查是否有连接线形状
    has_connectors = any(hasattr(s, "connector") and s.connector for s in shapes)
    # 检查是否水平排列
    cys = [b[5] for b in bboxes]
    if cys:
        y_variance = sum((cy - sum(cys) / len(cys)) ** 2 for cy in cys) / len(cys)
        y_std = math.sqrt(y_variance)
        avg_height = sum(b[3] for b in bboxes) / n
        if avg_height > 0 and y_std / avg_height < 0.5:
            if has_connectors or "→" in text or "—" in text:
                return True
    return False


def _detect_org_structure(shapes, bboxes) -> bool:
    """检测组织架构图：树状连线"""
    has_connectors = sum(1 for s in shapes if hasattr(s, "connector") and s.connector)
    if has_connectors >= 2:
        # 检查是否有层级分布
        cys = sorted(set(round(b[5] / 80000) * 80000 for b in bboxes))
        if len(cys) >= 2:
            return True
    return False


def _detect_logic_steps(text: str, n: int) -> bool:
    """检测逻辑步骤：编号序列"""
    # 匹配 [1] [2] [3] 或 步骤1 步骤2 或 1. 2. 3.
    numbered = re.findall(r'(?:步骤|step|阶段|phase)\s*[1-9]', text, re.IGNORECASE)
    bracketed = re.findall(r'[\[（(]\s*[1-9]\s*[\]）)]', text)
    dotted = re.findall(r'(?:^|\n)\s*[1-9]\s*[.、．]', text)
    return len(numbered) >= 2 or len(bracketed) >= 2 or len(dotted) >= 2


def _detect_gantt(bboxes, sw: int) -> bool:
    """检测甘特图：多个水平长条矩形"""
    n = len(bboxes)
    if n < 3:
        return False
    # 检查形状是否普遍较宽（宽度 > 高度 * 3）
    wide_count = sum(1 for b in bboxes if b[2] > b[3] * 3)
    if wide_count >= n * 0.6:
        # 检查是否垂直排列
        cys = [b[5] for b in bboxes]
        y_range = max(cys) - min(cys) if cys else 0
        if y_range > 0:
            return True
    return False


def _detect_kpi(shapes, text: str) -> bool:
    """检测指标体系：多个大字号数字"""
    # 统计文本中的数字密度
    digits = re.findall(r'\d+[%％万亿]?', text)
    if len(digits) >= 3:
        # 检查是否有大字号文本
        large_font_count = 0
        for shape in shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.size and run.font.size > Pt(24):
                            large_font_count += 1
        if large_font_count >= 2:
            return True
    return False


def _detect_phases(bboxes, sw: int, text: str) -> bool:
    """检测实施阶段：分阶段区块"""
    # 检查是否有"阶段"、"phase"等关键词
    phase_matches = re.findall(r'(?:阶段|phase)\s*[一二三四五六1-6]', text, re.IGNORECASE)
    if len(phase_matches) >= 2:
        return True
    # 检查是否有箭头连接的分段区块
    if "→" in text and len(bboxes) >= 3:
        return True
    return False


# ─────────────────────────────────────────────────────────
# PPT 拆页
# ─────────────────────────────────────────────────────────

def split_pptx(input_path: str, output_dir: str, name_prefix: str = "") -> list:
    """
    将多页 PPTX 拆分为单页 PPTX 文件。

    返回: [{"file": "slide_001.pptx", "path": "/full/path/..."}, ...]
    """
    prs = Presentation(input_path)
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    n_slides = len(prs.slides)

    if n_slides == 0:
        print("⚠ 输入文件没有幻灯片页")
        return []

    os.makedirs(output_dir, exist_ok=True)

    if not name_prefix:
        name_prefix = Path(input_path).stem

    results = []
    for i, slide in enumerate(prs.slides):
        # 创建新的演示文稿
        new_prs = Presentation()
        new_prs.slide_width = slide_width
        new_prs.slide_height = slide_height

        # 复制版式（如果可能）
        slide_layout = new_prs.slide_layouts[6]  # 空白版式
        new_slide = new_prs.slides.add_slide(slide_layout)

        # 复制所有形状
        for shape in slide.shapes:
            _copy_shape(shape, new_slide, slide_width, slide_height)

        # 保存
        filename = f"{name_prefix}_slide_{i + 1:03d}.pptx"
        filepath = os.path.join(output_dir, filename)
        new_prs.save(filepath)

        results.append({
            "slide_index": i + 1,
            "file": filename,
            "path": os.path.abspath(filepath),
        })
        print(f"  ✅ 拆分: slide {i + 1}/{n_slides} → {filename}")

    return results


def _copy_shape(source_shape, target_slide, slide_width, slide_height):
    """将源形状复制到目标幻灯片"""
    try:
        # 获取形状的 XML 元素并深拷贝
        from lxml import etree
        sp_xml = deepcopy(source_shape._element)
        target_slide.shapes._spTree.append(sp_xml)
    except Exception:
        # 降级：只复制基本属性
        pass


# ─────────────────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────────────────

def classify_slide(pptx_path: str, slide_index: int, slide_width: int, slide_height: int) -> dict:
    """对单个 PPTX 文件的第 slide_index 页进行四维分类"""
    prs = Presentation(pptx_path)
    slide = prs.slides[slide_index]

    sw = slide_width
    sh = slide_height

    content_shapes = get_content_shapes(slide, sw, sh)

    logic = classify_logic(slide, content_shapes, sw, sh)
    domain = classify_domain(slide)
    count = classify_count(content_shapes)
    chart_type = classify_chart_type(slide, content_shapes, sw, sh)

    return {
        "logic": logic,
        "domain": domain,
        "count": count,
        "chart-type": chart_type,
    }


def run(input_path: str, output_dir: str = None, name_prefix: str = ""):
    """主入口：拆页 + 分类 + 生成清单"""
    input_path = str(Path(input_path).resolve())

    if not os.path.exists(input_path):
        print(f"❌ 文件不存在: {input_path}")
        sys.exit(1)

    if not input_path.lower().endswith(".pptx"):
        print(f"❌ 不是 .pptx 文件: {input_path}")
        sys.exit(1)

    # 确定输出目录
    if not output_dir:
        stem = Path(input_path).stem
        output_dir = str(DEFAULT_OUTPUT_DIR / stem)
    output_dir = str(Path(output_dir).resolve())

    if not name_prefix:
        name_prefix = Path(input_path).stem

    print(f"📂 输入: {input_path}")
    print(f"📁 输出: {output_dir}")
    print()

    # 读取原始文件获取尺寸信息
    orig_prs = Presentation(input_path)
    slide_width = orig_prs.slide_width
    slide_height = orig_prs.slide_height
    n_slides = len(orig_prs.slides)
    print(f"📊 共 {n_slides} 页幻灯片 ({slide_width}x{slide_height} EMU)")
    print()

    # ── Step 1: 拆页 ──
    print("━━━ Step 1: 拆分幻灯片 ━━━")
    split_results = split_pptx(input_path, output_dir, name_prefix)
    print()

    # ── Step 2: 逐页分类 ──
    print("━━━ Step 2: 四维标签分类 ━━━")
    manifest_entries = []

    for entry in split_results:
        idx = entry["slide_index"]
        pptx_file = entry["path"]

        tags = classify_slide(pptx_file, 0, slide_width, slide_height)

        thumb_name = Path(entry["file"]).stem + ".png"
        manifest_entry = {
            "slide_index": idx,
            "file_name": entry["file"],
            "file_path": entry["path"],
            "thumbnail": os.path.join(output_dir, "thumbnails", thumb_name),
            "tags": tags,
            "confidence": None,  # AI 校验后填充
        }
        manifest_entries.append(manifest_entry)

        print(f"  slide {idx:3d}: "
              f"logic={tags['logic']:12s} "
              f"domain={tags['domain']:14s} "
              f"count={tags['count']:10s} "
              f"chart-type={tags['chart-type']}")

    print()

    # ── Step 3: 生成清单 ──
    print("━━━ Step 3: 生成模板清单 ━━━")
    manifest = {
        "meta": {
            "source": os.path.basename(input_path),
            "source_path": input_path,
            "total_slides": n_slides,
            "slide_width": slide_width,
            "slide_height": slide_height,
            "tag_version": "2.0",
            "output_dir": output_dir,
        },
        "templates": manifest_entries,
    }

    manifest_path = os.path.join(output_dir, "manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(f"  ✅ 清单已保存: {manifest_path}")
    print()

    # ── 汇总 ──
    print("━━━ 分类汇总 ━━━")
    for dim_id, dim_name in [
        ("logic", "逻辑拓扑"), ("domain", "内容领域"),
        ("count", "数量规模"), ("chart-type", "图形类型")
    ]:
        counter = Counter(e["tags"][dim_id] for e in manifest_entries)
        print(f"  {dim_name}:")
        for tag, cnt in counter.most_common():
            print(f"    {tag}: {cnt}")
    print()
    print(f"✅ 完成! 共 {n_slides} 个模板文件已输出到: {output_dir}")

    return manifest


# ─────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="PPT 拆页 + 四维标签自动分类工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", help="输入的 .pptx 文件路径")
    parser.add_argument("--output-dir", "-o", default=None,
                        help=f"输出目录 (默认: {DEFAULT_OUTPUT_DIR}/<文件名>/)")
    parser.add_argument("--name", "-n", default="",
                        help="输出文件名前缀 (默认: 使用输入文件名)")

    args = parser.parse_args()
    run(args.input, args.output_dir, args.name)


if __name__ == "__main__":
    main()
