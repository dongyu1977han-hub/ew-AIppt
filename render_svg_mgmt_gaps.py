"""
管理赋能（职能中台）CBM 三级业务能力大图 SVG 渲染器（带差距与痛点评估高亮版）
L0: 管理赋能（职能中台）
L1x6: 经营管理 / 财务资金 / 综合管理 / 法务风控 / 数智技术 / 审计监察
"""
import os

font = "Microsoft YaHei"

# ========================================================
# 核心痛点与差距分析数据字典 (精炼提炼，适配 PPT 卡片空间)
# ========================================================
GAPS = {
    "经营预算协同管理": "业财数据脱节，一体化协同不深入",
    "经营数据分析 Jun 诊断": "数据割裂，经营分析严重滞后",  # 兼容防错
    "经营数据分析与诊断": "数据割裂，经营分析严重滞后",
    "战略执行跟踪": "端到端可视化薄弱，缺少系统跟踪",
    "组织发展与长期人才激励": "干部输出库与长效激励尚在初期",
    "资金计划": "依赖线下人工，缺乏全局动态预测",
    "业财融合": "数据孤岛，业财税数据口径不一致",
    "集团化财务管理策略": "缺少跨区域复制的财务管控底座",
    "人才储备培养策略": "继任机制已启动，长效机制未成熟",
    "人才评价（认证、画像）": "尚未建立数字化人才全景画像库",
    "绩效管理": "绩效高度依赖手工，未完全适配业务",
    "合规管理": "涉外与ESG合规前置防线亟待建立",
    "法律支持及服务": "法务资源单薄，复杂交易设计弱",
    "内控管理": "内控多为事后，事前事中融合不足",
    "数智化战略与应用蓝图": "信息化刚起步，缺乏全局转型规划",
    "流程梳理与标准化": "缺乏战略牵引，未下沉至一线SOP",
    "数据治理与数据中台建设": "业财数据极度孤岛，口径不一未打通",
    "数智化产品应用与推广": "系统推广难，缺乏数字运营人才",
    "审计制度与标准管理": "大纲制度编制中，内控标准亟待细化",
    "日常监督与巡察": "缺少独立举报途径，日常监督难"
}

data = {
    "经营管理": {
        "D": [
            ("经营分析与决策支持", [
                "经营计划与KPI指标管理",
                "经营预算协同管理",
            ]),
            ("战略规划与管理", [
                "行业政策研究",
                "战略规划制定",
            ]),
            ("变革管理", [
                "变革项目统筹与协同",
            ]),
        ],
        "C": [
            ("经营分析与决策支持", [
                "经营数据分析与诊断",
                "监控预警与协同改进",
            ]),
            ("战略规划与管理", [
                "战略执行跟踪",
                "绩效评估考核",
            ]),
            ("变革管理", [
                "变革绩效跟踪",
                "企业文化管理",
            ]),
            ("董事会事务管理", [
                "董事会合规运作",
                "投资者关系管理",
                "公共关系管理",
            ]),
        ],
        "E": [
            ("变革管理", [
                "变革项目宣贯与推广",
            ]),
            ("经营分析与决策支持", [
                "组织发展与长期人才激励",
            ]),
        ],
    },
    "财务资金": {
        "D": [
            ("财务预算与规划", [
                "预算编制",
            ]),
            ("资金管理", [
                "资金计划",
            ]),
            ("税务管理", [
                "税务筹划",
            ]),
        ],
        "C": [
            ("财务预算与规划", [
                "预算执行与调整",
                "预算考核与分析",
                "成本控制与经营绩效评价",
            ]),
            ("资金管理", [
                "票据管理",
                "账户管理",
            ]),
        ],
        "E": [
            ("财务预算与规划", [
                "集团化财务管理策略",
            ]),
            ("财务核算", [
                "多法人主体财务核算",
                "业财融合",
                "财务共享服务",
                "合并报表与报告编制",
            ]),
            ("资金管理", [
                "资金往来与收付款",
                "资金对账",
            ]),
            ("税务管理", [
                "税务申报及缴纳",
            ]),
        ],
    },
    "综合管理": {
        "D": [
            ("人力资源规划", [
                "人力资源发展规划",
                "人才储备培养策略",
                "人力资源政策知识",
            ]),
            ("组织发展", [
                "组织机构管理",
                "岗位职级体系管理",
            ]),
        ],
        "C": [
            ("人才管理", [
                "薪酬社保管理",
                "绩效管理",
            ]),
            ("人力资源服务", [
                "员工关系管理",
            ]),
            ("综合管理", [
                "后勤物业",
                "工会党群行政",
                "车队管理",
                "食堂管理",
            ]),
        ],
        "E": [
            ("人才管理", [
                "人才（干部）选聘",
                "人才开发（培训、培养）",
                "人才评价（认证、画像）",
                "人才退出",
            ]),
            ("人力资源服务", [
                "人力资源共享服务",
                "员工自助服务",
            ]),
        ],
    },
    "法务风控": {
        "D": [
            ("法务管理", [
                "公司规章制度体系",
            ]),
            ("风控管理", [
                "风控体系建设",
            ]),
        ],
        "C": [
            ("法务管理", [
                "法律审核",
                "合规管理",
                "知识产权",
                "法律纠纷",
            ]),
            ("风控管理", [
                "内控管理",
                "全面风险管理",
            ]),
        ],
        "E": [
            ("法务管理", [
                "法律支持及服务",
            ]),
        ],
    },
    "数智技术": {
        "D": [
            ("数智化规划", [
                "数智化战略与应用蓝图",
                "数字化投资预算管理",
                "信息化制度与服务标准管理",
            ]),
        ],
        "C": [
            ("业务流程管理", [
                "流程绩效监控",
            ]),
            ("数智化开发管理", [
                "数智化项目管理",
                "架构设计与管理",
                "矿业知识库建设管理",
                "数据治理与数据中台建设",
            ]),
            ("数智化应用", [
                "AI应用与智能体管理",
                "对外服务数智化产品",
            ]),
            ("数字化运维管理", [
                "运维计划与资源协调管理",
                "网络与信息安全管理",
                "系统监控与故障处理",
            ]),
        ],
        "E": [
            ("业务流程管理", [
                "流程梳理与标准化",
                "端到端流程资产与版本管理",
                "流程验证优化与推广",
            ]),
            ("数智化开发管理", [
                "数智化应用建设",
            ]),
            ("数智化应用", [
                "数智化产品应用与推广",
            ]),
            ("数字化运维管理", [
                "IT基础设施建设与运行",
            ]),
        ],
    },
    "审计监察": {
        "D": [],
        "C": [
            ("审计管理", [
                "审计制度与标准管理",
                "审计项目执行管理",
                "审计监督与整改",
                "审计报告与档案管理",
            ]),
            ("监察管理", [
                "监察制度与标准管理",
                "日常监督与巡察",
                "廉洁教育与风险预警",
            ]),
        ],
        "E": [
            ("监察管理", [
                "案件处理与整改问责",
            ]),
        ],
    },
}

L1_ORDER = ["经营管理", "财务资金", "综合管理", "法务风控", "数智技术", "审计监察"]

# ========================================================
# 渲染参数
# ========================================================
svg_w, svg_h = 1280, 720

# 配色
c_l0 = "#1565C0"       # L0 深蓝
c_l1 = "#1E88E5"       # L1 蓝
c_l2 = "#D3D8E0"       # L2 灰蓝
c_l3_fill = "#FFFFFF"  # L3 白
c_l3_stroke = "#D1D5DB"
c_text_l1 = "#FFFFFF"
c_text_l2 = "#1F2D3D"
c_text_l3 = "#2B3D4F"
c_line = "#AEAEAE"

# 版面参数
margin_left = 35        # 左侧边距（放 D策略/C管理/E执行 标签）
margin_right = 8
margin_top = 90         # 标题 + 序言空间
l0_bar_h = 16           # L0 条高
l1_bar_h = 14           # L1 条高
content_top = margin_top + l0_bar_h + l1_bar_h + 2  # 内容区起始 Y

# 三层区域 Y 划分 (D/C/E)
avail_h = svg_h - 8 - content_top  # 可用高度
# 按比例分配：D=20%, C=45%, E=35%
d_h = int(avail_h * 0.20)
c_h = int(avail_h * 0.45)
e_h = avail_h - d_h - c_h

y_d = content_top
y_c = y_d + d_h
y_e = y_c + c_h

# 列参数
label_w = 24            # 左侧 D/C/E 标签宽度
content_left = margin_left + label_w + 3  # 内容区起始 X
content_w = svg_w - content_left - margin_right  # 内容区总宽
n_cols = 6
col_gap = 4
col_w = (content_w - (n_cols - 1) * col_gap) / n_cols

# 卡片参数
gap_y = 1.5             # 卡片纵向间距
fs = 7.5                # 字体大小
fs_l2 = 8               # L2 字体略大


def col_x(i):
    """第 i 列的起始 X"""
    return content_left + i * (col_w + col_gap)


def render_box(svg, x, y, w, h, fill, text, text_fill, bold=False, stroke=None):
    """渲染一个文本矩形"""
    stroke_attr = f' stroke="{stroke}" stroke-width="0.5"' if stroke else ''
    svg.append(f'  <rect x="{x}" y="{y}" width="{w}" height="{h}" fill="{fill}"{stroke_attr} rx="1" ry="1"/>')
    fw = 'bold' if bold else 'normal'
    fsize = fs_l2 if bold else fs
    svg.append(f'  <text x="{x + w/2}" y="{y + h/2 + 3}" font-family="{font}" font-size="{fsize}" font-weight="{fw}" fill="{text_fill}" text-anchor="middle">{text}</text>')


def render_column_tier(svg, col_idx, tier_data, y_start, tier_h):
    """
    在某一列的某一层（D/C/E）中自适应渲染 L2+L3 组。
    """
    if not tier_data:
        return
    
    x = col_x(col_idx)
    w = col_w
    
    # 收集所有的卡片，以及它们的基本高度
    cards = []
    for l2_name, l3_list in tier_data:
        # L2 卡片
        cards.append({"type": "l2", "name": l2_name, "base_h": 15})
        for l3_name in l3_list:
            gap_desc = GAPS.get(l3_name, None)
            if gap_desc:
                cards.append({"type": "l3_gap", "name": l3_name, "gap": gap_desc, "base_h": 25})
            else:
                cards.append({"type": "l3", "name": l3_name, "base_h": 15})
                
    # 计算所有卡片的基本总高度（包含 gap_y 间距）
    total_base_h = sum(c["base_h"] for c in cards) + (len(cards) - 1) * gap_y
    
    # 自适应缩放因子
    scale = 1.0
    if total_base_h > tier_h - 4:
        # 超过了可用空间，需要按比例压缩高度
        scale = (tier_h - 4) / total_base_h
        scale = max(0.68, scale)  # 设立保护阈值以保证基本可读性
        
    # 计算实际卡片大小并居中对齐
    actual_cards = []
    for c in cards:
        h = c["base_h"] * scale
        actual_cards.append({**c, "h": h})
        
    actual_needed_h = sum(c["h"] for c in actual_cards) + (len(cards) - 1) * gap_y
    cy = y_start + max(2, (tier_h - actual_needed_h) / 2)
    
    for c in actual_cards:
        ch = c["h"]
        if c["type"] == "l2":
            render_box(svg, x, cy, w, ch, c_l2, c["name"], c_text_l2, bold=True)
        elif c["type"] == "l3_gap":
            # 渲染带差距评估高亮的多行卡片：象牙淡黄背景 (#FFF9E6)，警示橘黄边框 (#FFB000)，高级暗红痛点字
            svg.append(f'  <rect x="{x}" y="{cy}" width="{w}" height="{ch}" fill="#FFF9E6" stroke="#FFA000" stroke-width="0.8" rx="2" ry="2"/>')
            
            fs_name = max(6.0, 7.2 * scale)
            fs_gap = max(4.8, 5.6 * scale)
            
            # 卡片内部分行 Y 位置计算
            y_name = cy + ch * 0.42
            y_gap = cy + ch * 0.82
            
            # 能力名称（左对齐，左侧留出 12 像素给小图标 ⚠）
            svg.append(f'  <text x="{x + 13}" y="{y_name}" font-family="{font}" font-size="{fs_name}" font-weight="bold" fill="#3E2723" text-anchor="start">{c["name"]}</text>')
            # 亮橘红色警告小图标 ⚠
            svg.append(f'  <text x="{x + 4}" y="{y_name}" font-family="{font}" font-size="{fs_name}" font-weight="bold" fill="#E65100" text-anchor="start">⚠</text>')
            # 第二行：痛点说明（暗红橘字 #BF360C）
            svg.append(f'  <text x="{x + 4}" y="{y_gap}" font-family="{font}" font-size="{fs_gap}" fill="#BF360C" text-anchor="start">核心痛点: {c["gap"]}</text>')
        else:
            render_box(svg, x, cy, w, ch, c_l3_fill, c["name"], c_text_l3, stroke=c_l3_stroke)
            
        cy += ch + gap_y


def generate_svg():
    svg = []
    svg.append(f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {svg_w} {svg_h}" width="{svg_w}" height="{svg_h}">')
    
    # 背景
    svg.append(f'  <rect width="{svg_w}" height="{svg_h}" fill="#FFFFFF"/>')
    
    # ---- 标题 ----
    svg.append(f'  <text x="35" y="30" font-family="{font}" font-size="22" font-weight="bold" fill="#1A2D3D">5.6 管理赋能业务能力差距评估</text>')
    
    # ---- 序言 ----
    line1 = "管理赋能（职能中台）共计包含6个L1业务能力，23个L2业务能力和88个L3业务能力。根据最新评估，共有19个L3业务能力存在差距，"
    line2 = "主要集中在业财一体化协同、数据孤岛治理、长效人才机制、前置合规防线及数智化应用推广等方面，亟待针对核心痛点进行补强和提升。"
    svg.append(f'  <text x="35" y="48" font-family="{font}" font-size="10" fill="#5F6F65">{line1}</text>')
    svg.append(f'  <text x="35" y="60" font-family="{font}" font-size="10" fill="#5F6F65">{line2}</text>')
    
    # ---- 颜色图例说明 (Legend) ----
    legend_x = svg_w - 240
    svg.append(f'  <rect x="{legend_x}" y="18" width="8" height="8" fill="#FFF9E6" stroke="#FFA000" stroke-width="0.5"/>')
    svg.append(f'  <text x="{legend_x + 12}" y="25" font-family="{font}" font-size="8.5" fill="#2B3D4F">能力存在差距 / 亟待优化 (共 19 项)</text>')
    
    # ---- 左侧 D/C/E 标签 + 纵向分隔线 ----
    svg.append(f'  <line x1="{margin_left + label_w}" y1="{content_top}" x2="{margin_left + label_w}" y2="{svg_h - 10}" stroke="{c_line}" stroke-width="1"/>')
    
    # 横向虚线
    svg.append(f'  <line x1="40" y1="{y_c}" x2="{svg_w - 10}" y2="{y_c}" stroke="{c_line}" stroke-width="0.8" stroke-dasharray="4,4"/>')
    svg.append(f'  <line x1="40" y1="{y_e}" x2="{svg_w - 10}" y2="{y_e}" stroke="{c_line}" stroke-width="0.8" stroke-dasharray="4,4"/>')
    
    # D 策略
    lx = margin_left + label_w / 2
    d_cy = y_d + d_h / 2
    svg.append(f'  <text x="{lx}" y="{d_cy - 8}" font-family="{font}" font-size="12" font-weight="bold" fill="#1A2D3D" text-anchor="middle">D</text>')
    svg.append(f'  <text x="{lx}" y="{d_cy + 4}" font-family="{font}" font-size="10" fill="#1A2D3D" text-anchor="middle">策</text>')
    svg.append(f'  <text x="{lx}" y="{d_cy + 16}" font-family="{font}" font-size="10" fill="#1A2D3D" text-anchor="middle">略</text>')
    
    # C 管理
    c_cy = y_c + c_h / 2
    svg.append(f'  <text x="{lx}" y="{c_cy - 8}" font-family="{font}" font-size="12" font-weight="bold" fill="#1A2D3D" text-anchor="middle">C</text>')
    svg.append(f'  <text x="{lx}" y="{c_cy + 4}" font-family="{font}" font-size="10" fill="#1A2D3D" text-anchor="middle">管</text>')
    svg.append(f'  <text x="{lx}" y="{c_cy + 16}" font-family="{font}" font-size="10" fill="#1A2D3D" text-anchor="middle">理</text>')
    
    # E 执行
    e_cy = y_e + e_h / 2
    svg.append(f'  <text x="{lx}" y="{e_cy - 8}" font-family="{font}" font-size="12" font-weight="bold" fill="#1A2D3D" text-anchor="middle">E</text>')
    svg.append(f'  <text x="{lx}" y="{e_cy + 4}" font-family="{font}" font-size="10" fill="#1A2D3D" text-anchor="middle">执</text>')
    svg.append(f'  <text x="{lx}" y="{e_cy + 16}" font-family="{font}" font-size="10" fill="#1A2D3D" text-anchor="middle">行</text>')
    
    # ---- L0 顶条 ----
    l0_x = content_left
    l0_w = n_cols * col_w + (n_cols - 1) * col_gap
    svg.append(f'  <rect x="{l0_x}" y="{margin_top}" width="{l0_w}" height="{l0_bar_h}" fill="{c_l0}" rx="1" ry="1"/>')
    svg.append(f'  <text x="{l0_x + l0_w/2}" y="{margin_top + l0_bar_h/2 + 3.5}" font-family="{font}" font-size="10" font-weight="bold" fill="#FFFFFF" text-anchor="middle">管理赋能（职能中台）</text>')
    
    # ---- L1 列头 ----
    for i, l1_name in enumerate(L1_ORDER):
        x = col_x(i)
        y = margin_top + l0_bar_h
        svg.append(f'  <rect x="{x}" y="{y}" width="{col_w}" height="{l1_bar_h}" fill="{c_l1}" rx="1" ry="1"/>')
        svg.append(f'  <text x="{x + col_w/2}" y="{y + l1_bar_h/2 + 3}" font-family="{font}" font-size="9" font-weight="bold" fill="{c_text_l1}" text-anchor="middle">{l1_name}</text>')
    
    # ---- 渲染每列的 D/C/E 数据 ----
    for i, l1_name in enumerate(L1_ORDER):
        col_data = data[l1_name]
        render_column_tier(svg, i, col_data.get("D", []), y_d, d_h)
        render_column_tier(svg, i, col_data.get("C", []), y_c, c_h)
        render_column_tier(svg, i, col_data.get("E", []), y_e, e_h)
    
    svg.append('</svg>')
    
    output_dir = r"projects\mingxin_capability_map_ppt169_20260525\svg_output"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "02_mgmt_platform.svg")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(svg))
    print(f"SVG 生成完毕: {output_path}")


if __name__ == "__main__":
    generate_svg()
