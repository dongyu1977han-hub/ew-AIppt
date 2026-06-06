#!/usr/bin/env python3
"""
ppt_classify_ai.py — AI 校验四维标签分类

读取 manifest.json（规则引擎输出）+ 缩略图 PNG，
通过多模态 AI 校验并修正每页的四维标签。

用法:
    python ppt_classify_ai.py <manifest.json> [--api-key <key>] [--base-url <url>]

环境变量:
    ANTHROPIC_API_KEY  — Anthropic API 密钥
    OPENAI_API_KEY     — OpenAI 兼容 API 密钥
    AI_BASE_URL        — 自定义 API 地址
"""

import argparse
import base64
import io
import json
import os
import sys
import time
from pathlib import Path

# Windows 控制台 UTF-8
if sys.platform == "win32" and hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import requests

# ─────────────────────────────────────────────────────────
# 标签 Schema
# ─────────────────────────────────────────────────────────

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent.parent.parent
TAG_SCHEMA_PATH = PROJECT_ROOT / "skills" / "ew-AIppt" / "templates" / "_tag_schema.json"

VALID_TAGS = {
    "logic": ["parallel", "sequential", "star", "contrast", "hierarchy",
              "matrix", "cycle", "pyramid", "other"],
    "domain": ["strategy", "current-state", "architecture", "technology",
               "organization", "operation", "other"],
    "count": ["count-1", "count-2", "count-3", "count-4", "count-5",
              "count-6", "count-7-plus"],
    "chart-type": ["timeline", "house-chart", "logic-steps", "value-tree",
                   "business-arch", "app-arch", "data-arch", "tech-arch",
                   "process-value-chain", "scenario-map", "org-structure",
                   "gantt", "kpi-system", "implementation-phase", "ecosystem", "other"],
}

# 中文标签名映射（用于 prompt 中给 AI 参考）
TAG_NAMES = {
    "logic": {
        "parallel": "并列", "sequential": "递进", "star": "星型",
        "contrast": "对比", "hierarchy": "层级", "matrix": "矩阵",
        "cycle": "循环", "pyramid": "棱锥", "other": "其它",
    },
    "domain": {
        "strategy": "战略", "current-state": "现状", "architecture": "架构",
        "technology": "技术", "organization": "组织", "operation": "运营",
        "other": "其它",
    },
    "count": {
        "count-1": "一项", "count-2": "两项", "count-3": "三项",
        "count-4": "四项", "count-5": "五项", "count-6": "六项",
        "count-7-plus": "七项及以上",
    },
    "chart-type": {
        "timeline": "时间轴", "house-chart": "屋型图", "logic-steps": "逻辑步骤",
        "value-tree": "价值树", "business-arch": "业务架构", "app-arch": "应用架构",
        "data-arch": "数据架构", "tech-arch": "技术架构",
        "process-value-chain": "流程价值链", "scenario-map": "场景图",
        "org-structure": "组织架构", "gantt": "甘特图", "kpi-system": "指标体系",
        "implementation-phase": "实施阶段", "ecosystem": "生态图", "other": "其它",
    },
}

DIM_NAMES = {
    "logic": "逻辑拓扑标签",
    "domain": "内容领域标签",
    "count": "数量规模标签",
    "chart-type": "图形类型标签",
}


# ─────────────────────────────────────────────────────────
# Prompt 构建
# ─────────────────────────────────────────────────────────

def build_system_prompt() -> str:
    """构建系统 prompt"""
    schema_text = ""
    for dim_id in ["logic", "domain", "count", "chart-type"]:
        tags = VALID_TAGS[dim_id]
        names = TAG_NAMES[dim_id]
        tag_list = ", ".join(f"{t}({names[t]})" for t in tags)
        schema_text += f"  {DIM_NAMES[dim_id]}: {tag_list}\n"

    return f"""你是一个 PPT 模板标签分类专家。你的任务是审核一张 PPT 幻灯片的四维标签分类结果。

## 四维标签体系

{schema_text}

## 你的任务

你会收到：
1. 一张 PPT 幻灯片的缩略图
2. 规则引擎自动分类的结果（可能有误）

你需要：
1. 仔细观察缩略图中的布局、内容、图形类型
2. 判断规则引擎的分类是否正确
3. 如果有误，修正为正确的标签值
4. 每个维度只能选择一个值

## 判断要点

- **逻辑拓扑**: 观察页面内容块的空间排列关系（并列=横向/纵向均匀排列, 递进=有方向性的顺序, 星型=中心辐射, 对比=左右/上下两组, 层级=树状上少下多, 矩阵=网格, 循环=环形, 棱锥=三角分层）
- **内容领域**: 根据文本内容判断所属领域
- **数量规模**: 数页面中独立的内容区块数量（不是 shape 总数，是真正承载内容的区块）
- **图形类型**: 判断页面使用的主要可视化图形类型

## 输出格式

只输出 JSON，不要其他文字：
{{"logic": "...", "domain": "...", "count": "...", "chart-type": "..."}}"""


def build_user_prompt(rule_tags: dict, slide_index: int) -> str:
    """构建用户 prompt（含规则引擎结果）"""
    rule_text = json.dumps(rule_tags, ensure_ascii=False)
    return f"""以下是规则引擎对第 {slide_index} 页的自动分类结果：

{rule_text}

请观察缩略图，审核并修正这个分类结果。只输出修正后的 JSON。"""


# ─────────────────────────────────────────────────────────
# API 调用
# ─────────────────────────────────────────────────────────

def encode_image_base64(image_path: str) -> str:
    """将图片编码为 base64"""
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def call_claude_api(system_prompt: str, user_prompt: str, image_path: str,
                    api_key: str, base_url: str = "https://api.anthropic.com",
                    model: str = "claude-sonnet-4-20250514") -> str:
    """调用 Anthropic Claude API"""
    image_b64 = encode_image_base64(image_path)

    # 检测图片类型
    ext = Path(image_path).suffix.lower()
    media_type = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                  "gif": "image/gif", "webp": "image/webp"}.get(ext.lstrip("."), "image/png")

    payload = {
        "model": model,
        "max_tokens": 256,
        "system": system_prompt,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": image_b64,
                        },
                    },
                    {
                        "type": "text",
                        "text": user_prompt,
                    },
                ],
            }
        ],
    }

    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }

    resp = requests.post(f"{base_url}/v1/messages", json=payload, headers=headers, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    return data["content"][0]["text"]


def call_openai_api(system_prompt: str, user_prompt: str, image_path: str,
                    api_key: str, base_url: str = "https://api.openai.com",
                    model: str = "gpt-4o") -> str:
    """调用 OpenAI 兼容 API"""
    image_b64 = encode_image_base64(image_path)
    ext = Path(image_path).suffix.lower()
    media_type = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg"}.get(
        ext.lstrip("."), "image/png")

    payload = {
        "model": model,
        "max_tokens": 256,
        "messages": [
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {
                        "url": f"data:{media_type};base64,{image_b64}"
                    }},
                    {"type": "text", "text": user_prompt},
                ],
            },
        ],
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    resp = requests.post(f"{base_url}/v1/chat/completions", json=payload, headers=headers, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"]


def parse_ai_response(response_text: str) -> dict | None:
    """解析 AI 返回的 JSON"""
    text = response_text.strip()
    # 尝试直接解析
    try:
        result = json.loads(text)
        return _validate_tags(result)
    except json.JSONDecodeError:
        pass
    # 尝试提取 JSON 块
    import re
    match = re.search(r'\{[^}]+\}', text)
    if match:
        try:
            result = json.loads(match.group())
            return _validate_tags(result)
        except json.JSONDecodeError:
            pass
    return None


def _validate_tags(tags: dict) -> dict | None:
    """验证标签值是否在合法范围内"""
    result = {}
    for dim_id in ["logic", "domain", "count", "chart-type"]:
        val = tags.get(dim_id)
        if val and val in VALID_TAGS[dim_id]:
            result[dim_id] = val
        else:
            return None  # 非法值，整体作废
    return result


# ─────────────────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────────────────

def verify_manifest(manifest_path: str, api_key: str,
                    api_type: str = "claude",
                    base_url: str = None,
                    model: str = None) -> dict:
    """读取 manifest，逐页 AI 校验，返回更新后的 manifest"""
    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    templates = manifest["templates"]
    output_dir = Path(manifest["meta"]["output_dir"])

    # API 配置
    if api_type == "claude":
        base_url = base_url or os.environ.get("AI_BASE_URL", "https://api.anthropic.com")
        model = model or "claude-sonnet-4-20250514"
        call_fn = lambda sys_p, user_p, img: call_claude_api(
            sys_p, user_p, img, api_key, base_url, model)
    else:
        base_url = base_url or os.environ.get("AI_BASE_URL", "https://api.openai.com")
        model = model or "gpt-4o"
        call_fn = lambda sys_p, user_p, img: call_openai_api(
            sys_p, user_p, img, api_key, base_url, model)

    system_prompt = build_system_prompt()

    print(f"📂 Manifest: {manifest_path}")
    print(f"🤖 API: {api_type} / {model}")
    print(f"📊 共 {len(templates)} 页待校验")
    print()

    corrected = 0
    errors = 0

    for entry in templates:
        idx = entry["slide_index"]
        rule_tags = entry["tags"]

        # 查找缩略图
        thumb_path = entry.get("thumbnail_path")
        if not thumb_path or not Path(thumb_path).exists():
            # 尝试根据文件名推断
            thumb_name = Path(entry["file_name"]).stem + ".png"
            thumb_path = str(output_dir / thumb_name)
        if not Path(thumb_path).exists():
            print(f"  slide {idx:3d}: ⚠ 无缩略图，跳过")
            continue

        user_prompt = build_user_prompt(rule_tags, idx)

        try:
            response_text = call_fn(system_prompt, user_prompt, thumb_path)
            ai_tags = parse_ai_response(response_text)

            if ai_tags is None:
                print(f"  slide {idx:3d}: ⚠ AI 返回无法解析: {response_text[:80]}")
                errors += 1
                continue

            # 对比差异
            changes = []
            for dim_id in ["logic", "domain", "count", "chart-type"]:
                if ai_tags[dim_id] != rule_tags[dim_id]:
                    changes.append(f"{dim_id}: {rule_tags[dim_id]} → {ai_tags[dim_id]}")

            if changes:
                corrected += 1
                entry["tags"] = ai_tags
                entry["tags_corrected"] = True
                entry["tags_rule"] = rule_tags  # 保留规则结果作为参考
                print(f"  slide {idx:3d}: ✏ 修正: {', '.join(changes)}")
            else:
                entry["tags_corrected"] = False
                print(f"  slide {idx:3d}: ✓ 一致")

        except requests.exceptions.RequestException as e:
            print(f"  slide {idx:3d}: ❌ API 错误: {e}")
            errors += 1
        except Exception as e:
            print(f"  slide {idx:3d}: ❌ 错误: {e}")
            errors += 1

        # 限流：每次请求间隔
        time.sleep(0.5)

    print()
    print(f"✅ 校验完成: {corrected} 页修正, {errors} 页错误")

    # 更新 meta
    manifest["meta"]["ai_verified"] = True
    manifest["meta"]["ai_model"] = model
    manifest["meta"]["ai_corrected_count"] = corrected
    manifest["meta"]["ai_error_count"] = errors

    return manifest


def main():
    parser = argparse.ArgumentParser(
        description="AI 校验四维标签分类结果",
    )
    parser.add_argument("manifest", help="manifest.json 文件路径")
    parser.add_argument("--api-key", "-k", default=None,
                        help="API 密钥 (或设置环境变量 ANTHROPIC_API_KEY / OPENAI_API_KEY)")
    parser.add_argument("--api-type", "-t", default="claude", choices=["claude", "openai"],
                        help="API 类型 (默认: claude)")
    parser.add_argument("--base-url", "-u", default=None,
                        help="自定义 API 地址")
    parser.add_argument("--model", "-m", default=None,
                        help="模型名称 (默认: claude-sonnet-4-20250514 / gpt-4o)")
    parser.add_argument("--output", "-o", default=None,
                        help="输出 manifest 路径 (默认: 覆盖原文件)")

    args = parser.parse_args()

    # 获取 API key
    api_key = args.api_key
    if not api_key:
        if args.api_type == "claude":
            api_key = os.environ.get("ANTHROPIC_API_KEY")
        else:
            api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("❌ 未提供 API 密钥，请通过 --api-key 或环境变量设置")
        sys.exit(1)

    manifest = verify_manifest(args.manifest, api_key, args.api_type,
                               args.base_url, args.model)

    # 保存
    output_path = args.output or args.manifest
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print(f"💾 已保存: {output_path}")


if __name__ == "__main__":
    main()
