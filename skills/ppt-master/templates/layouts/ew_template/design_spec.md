# Design Specification: 亿汇智标准模版 (蓝灰 2026)

## I. Identity
- **ID**: `ew_template`
- **Name**: 亿汇智标准模版 (蓝灰 2026 V1.0)
- **Category**: `brand`
- **Canvas**: PPT 16:9 (1280×720)
- **Style**: Professional Corporate, 央国企正式商务汇报

---

## II. Color Palette (严格遵循设计规范)

| Token | Hex | Usage |
|---|---|---|
| Primary Blue | `#4067A1` | 标题、区块标题、核心视觉锚点、页脚主色块 |
| Accent Blue 1 | `#55A4DC` | 次级标题、图标配色、卡片边框、目录当前项高亮 |
| Accent Blue 2 | `#8EB4E3` | 浅蓝辅助 |
| Accent Red | `#AF1E23` | 正文关键词高亮、核心数据强调 |
| Accent Orange | `#E46C0A` | 渐变分隔线橙色端、章节色块 |
| Accent Gray | `#6D6C68` | 第四辅色、说明文字 |
| Neutral Light | `#D3D3D3` | 第三辅色、灰色装饰 |
| BG White | `#FFFFFF` | 页面主背景 |
| BG LightBlue | `#EAF3FB` | 内容区块浅蓝底色 |
| Text Dark | `#0D0D0D` | 正文主要文字 |
| Text Mid | `#494949` | 正文次要文字 |
| Text Light | `#666666` | 备注、页码 |
| Footer BG | `#428AC8` | 页脚条带背景色 |

---

## III. Typography

| Element | Font | Size (1280×720) | Weight | Color |
|---|---|---|---|---|
| 封面主标题 | Microsoft YaHei | 48px | Bold | `#FFFFFF` |
| 封面副标题 | Microsoft YaHei | 28px | Regular | `#FFFFFF` |
| 封面公司名 | Microsoft YaHei | 18px | Regular | `#FFFFFF` |
| 页面标题 | Microsoft YaHei | 37px | Bold | `#4067A1` |
| 副标题/区块标题 | Microsoft YaHei | 21px | Bold | `#4067A1` |
| 引言正文 | Microsoft YaHei | 19px | Light | `#0D0D0D` |
| 定制内容 | Microsoft YaHei | 16–24px | Regular | `#0D0D0D` |
| 备注 | Microsoft YaHei | 11px | Regular | `#666666` |
| 页码 | Microsoft YaHei | 11px | Regular | `#FFFFFF` |
| 目录项 | Microsoft YaHei | 24px | Bold | `#FFFFFF` / `#D3D3D3` |

---

## IV. Layout Structure (核心布局规范)

### 全局元素（每页固定）
- **页面背景**: 纯白 `#FFFFFF`，无纹理
- **页脚色条**: 全宽，高度约14px（y=706），颜色 `#428AC8`（蓝色）
- **安全边距**: 左右各 64px（约5%页宽）

### 标题区（内容页）
- 标题位置：`x=64, y=36`，微软雅黑 Bold 37px，颜色 `#4067A1`
- 分隔线：标题下方约36px处，全宽（64到1216），高度3px
  - 渐变：`#4067A1`（左）→ `#55A4DC`（中）→ `#D3D3D3`（右）
- 副标题（若有）：分隔线下方24px，微软雅黑 Bold 21px，`#4067A1`

### 内容区
- 起始Y（无副标题）：约120px
- 起始Y（有副标题）：约155px
- 结束Y：约695px（页脚上方）

---

## V. Page Types

| 文件 | 类型 | 说明 |
|---|---|---|
| `01_cover.svg` | 封面 | 全屏背景图+蓝色叠加，标题左下角 |
| `02_toc.svg` | 目录页 | 左侧装饰+右侧目录列表，支持6-8条目 |
| `03_content_with_subtitle.svg` | 内容页（有副标题） | 标题+分隔线+副标题+内容区 |
| `04_content_no_subtitle.svg` | 内容页（无副标题） | 标题+分隔线+内容区，更大内容空间 |
| `05_ending.svg` | 结尾页 | 蓝色背景+感谢语，左侧斜切三角构图 |

---

## VI. Placeholder Mapping

### 封面 (`01_cover.svg`)
- `{{TITLE}}` — 主标题（项目名称）
- `{{SUBTITLE}}` — 副标题（例：方案建议书）
- `{{CLIENT}}` — 委托单位（例：中国****集团）
- `{{DATE}}` — 日期（例：2026年4月）

### 目录页 (`02_toc.svg`)
- `{{TOC_ITEM_1}}` ~ `{{TOC_ITEM_N}}` — 目录项（格式：第X篇 标题）
- `{{CURRENT_CHAPTER}}` — 当前高亮章节（可选，用于章节过渡页）

### 内容页 (`03_content_with_subtitle.svg` / `04_content_no_subtitle.svg`)
- **固定元素**: 右上角固定展示亿汇智官方 Logo (`inline_46d3b2d7f02f.png`)
- `{{PAGE_TITLE}}` — 页面标题（微软雅黑 Bold 37px 蓝色）
- `{{SUBTITLE}}` — 副标题（仅03用）
- `{{LEAD_TEXT}}` — 引言正文（14号以上）
- `{{CONTENT_AREA}}` — 定制内容区（表格/图表/分栏等）
- `{{REMARK}}` — 备注（页脚上方，8号灰色，可选）
- `{{PAGE_NUM}}` — 页码（白色，页脚内）

### 结尾页 (`05_ending.svg`)
- `{{CLOSING_TITLE}}` — 结束语（例：感谢聆听，敬请指正！）
- `{{CLOSING_BODY}}` — 补充内容（核心观点摘要，可选）

---

## VII. Visual Signature Elements

1. **斜切三角形装饰** (目录页/结尾页): 从左侧约1/4处到右下角的蓝色边框斜切三角形，`stroke: #3C64A0`
2. **全宽渐变分隔线** (内容页): 蓝→浅蓝→灰的水平渐变线，是系列识别标志
3. **彩色页脚条带** (所有页): `#428AC8` 纯蓝色横条，高度14px，紧贴底部
4. **右侧背景图案** (目录页): 斜切剪切区域内的灰蓝色装饰图
5. **内容区浅蓝底色块** (`#EAF3FB`): 用于区分信息模块
