#!/usr/bin/env python3
"""将研究方案 JSON 转换为 Markdown 文档。支持 v2 的 stages 结构和 v1 的 methods 结构。"""

import json
import sys
import os
from datetime import datetime


def _format_stage_header(stage, stage_idx):
    collection = stage.get("collection_method", "")
    techniques = stage.get("embedded_techniques", [])
    name = stage.get("name", "")
    display_name = name or collection

    parts = [f"阶段{stage_idx}：{display_name}"]
    if techniques:
        parts.append(f"（含嵌入技术：{'、'.join(techniques)}）")
    return "".join(parts)


def _format_content_outline(content_outline):
    lines = []
    if not content_outline:
        return lines

    lines.append("#### 研究内容概览")
    lines.append("")

    # 主采集方式
    if content_outline.get("collection_method"):
        lines.append(f"**主采集方式**：{content_outline['collection_method']}")
    if content_outline.get("embedded_techniques"):
        lines.append(f"**嵌入技术**：{'、'.join(content_outline['embedded_techniques'])}")
    lines.append("")

    # 模块/话题列表
    sections = content_outline.get("sections", [])
    if sections:
        for i, sec in enumerate(sections, 1):
            title = sec.get("title", f"模块{i}")
            lines.append(f"**{title}**{('（' + sec.get('duration', '') + '）') if sec.get('duration') else ''}")
            if sec.get("description"):
                lines.append(f"- {sec['description']}")
            probes = sec.get("probes", [])
            for probe in probes:
                lines.append(f"  - 探针：{probe}")
            lines.append("")
        return lines

    # 问卷表格形式（从 module_table 读取）
    modules = content_outline.get("modules", [])
    if modules:
        lines.append("| 模块 | 内容 | 题型 | 题量 |")
        lines.append("|------|------|------|------|")
        for mod in modules:
            name = mod.get("name", "")
            content = mod.get("content", "")
            qtype = mod.get("type", "")
            count = mod.get("count", "")
            lines.append(f"| {name} | {content} | {qtype} | {count} |")
        lines.append("")
        return lines

    # 任务表格形式
    tasks = content_outline.get("tasks", [])
    if tasks:
        lines.append("| 任务编号 | 任务场景 | 关键观察点 |")
        lines.append("|---------|---------|-----------|")
        for task in tasks:
            tid = task.get("id", "")
            scene = task.get("scene", "")
            observe = task.get("observation", "")
            lines.append(f"| {tid} | {scene} | {observe} |")
        lines.append("")
        return lines

    # 自由文本描述
    text = content_outline.get("description", "")
    if text:
        lines.append(text)
        lines.append("")

    return lines


def _format_embed_task(embed_task):
    lines = []
    if not embed_task:
        return lines
    lines.append(f"**嵌入任务——{embed_task.get('technique', '')}**：")
    lines.append(embed_task.get("description", ""))
    lines.append("")
    return lines


def _format_content_outline_v1(method):
    """Handle v1 flat content outline from embedded techniques."""
    lines = []
    outline = method.get("content_outline")

    if isinstance(outline, dict):
        # Try new format first
        formatted = _format_content_outline(outline)
        if formatted:
            lines.extend(formatted)
        # Also try embed_task
        embed = outline.get("embed_task")
        if embed:
            lines.extend(_format_embed_task(embed))
    elif isinstance(outline, str):
        lines.append("#### 研究内容概览")
        lines.append("")
        lines.append(outline)
        lines.append("")

    return lines


def json_to_markdown(plan: dict) -> str:
    lines = []

    # 标题
    title = plan.get("plan_title", "研究方案")
    lines.append(f"# {title}")
    lines.append("")

    # 一、研究背景与目标
    lines.append("## 一、研究背景与目标")
    lines.append("")
    if plan.get("research_background"):
        lines.append("### 1.1 研究背景")
        lines.append("")
        lines.append(plan["research_background"])
        lines.append("")

    if plan.get("research_objectives"):
        lines.append("### 1.2 研究目标")
        lines.append("")
        for i, obj in enumerate(plan["research_objectives"], 1):
            lines.append(f"{i}. {obj}")
        lines.append("")

    if plan.get("research_type"):
        lines.append("### 1.3 研究类型")
        lines.append("")
        lines.append(plan["research_type"])
        lines.append("")

    # 二、目标用户群体
    if plan.get("target_users"):
        lines.append("## 二、目标用户群体")
        lines.append("")
        tu = plan["target_users"]
        if isinstance(tu, dict):
            if tu.get("description"):
                lines.append("### 2.1 用户群体概述")
                lines.append("")
                lines.append(tu["description"])
                lines.append("")
            if tu.get("segments"):
                lines.append("### 2.2 用户分层")
                lines.append("")
                for i, seg in enumerate(tu["segments"], 1):
                    lines.append(f"- **群体{i}**：{seg}")
                lines.append("")
        elif isinstance(tu, str):
            lines.append(tu)
            lines.append("")

    # 获取研究阶段列表：优先用 stages，兼容 methods
    stages = plan.get("stages") or plan.get("methods") or []

    if stages:
        lines.append("## 三、研究方法设计")
        lines.append("")

        for idx, stage in enumerate(stages, 1):
            # 判断是 v2 stages 还是 v1 methods
            is_v2 = "collection_method" in stage

            if is_v2:
                section_title = _format_stage_header(stage, idx)
            else:
                name = stage.get("name", "未命名方法")
                stage_num = stage.get("stage", idx)
                section_title = f"### 3.{stage_num} 阶段{stage_num}：{name}"

            lines.append(section_title)
            lines.append("")

            # 研究目的
            if stage.get("purpose"):
                lines.append("#### 研究目的")
                lines.append("")
                lines.append(stage["purpose"])
                lines.append("")

            # 研究内容概览 - v2 format
            if is_v2 and stage.get("content_outline"):
                lines.extend(_format_content_outline(stage["content_outline"]))
                # Also check for embed_task within content_outline
                embed = stage["content_outline"].get("embed_task")
                if embed:
                    lines.extend(_format_embed_task(embed))

            # 研究内容概览 - v1 format (methods style)
            if not is_v2 and stage.get("content_outline"):
                lines.extend(_format_content_outline_v1(stage))

            # 样本量设计
            if stage.get("sample_size"):
                lines.append("#### 样本量设计")
                lines.append("")
                ss = stage["sample_size"]
                if isinstance(ss, dict):
                    if ss.get("recommended"):
                        lines.append(f"- **推荐样本量**：{ss['recommended']}")
                    if ss.get("rationale"):
                        lines.append(f"- **依据**：{ss['rationale']}")
                    if ss.get("breakdown"):
                        lines.append(f"- **分层说明**：{ss['breakdown']}")
                    if ss.get("calculation"):
                        lines.append(f"- **计算过程**：{ss['calculation']}")
                elif isinstance(ss, str):
                    lines.append(ss)
                lines.append("")

            # 样本特征要求
            if stage.get("sample_characteristics"):
                lines.append("#### 样本特征要求")
                lines.append("")
                sc = stage["sample_characteristics"]
                if isinstance(sc, dict):
                    cats = [
                        ("人口学特征", "demographic"),
                        ("行为特征", "behavioral"),
                        ("态度特征", "attitudinal"),
                        ("排除条件", "exclusion"),
                    ]
                    for label, key in cats:
                        if sc.get(key):
                            lines.append(f"**{label}**：")
                            for item in sc[key]:
                                lines.append(f"- {item}")
                            lines.append("")
                elif isinstance(sc, str):
                    lines.append(sc)
                    lines.append("")

            # 样本配额要求
            if stage.get("quota_requirements"):
                lines.append("#### 样本配额要求")
                lines.append("")
                qr = stage["quota_requirements"]
                if isinstance(qr, dict):
                    if qr.get("total_target"):
                        lines.append(f"- **总目标样本量**：{qr['total_target']}")
                    if qr.get("dimensions"):
                        lines.append("")
                        lines.append("| 配额维度 | 分层 | 比例 |")
                        lines.append("|---------|------|------|")
                        for dim in qr["dimensions"]:
                            dname = dim.get("name", "")
                            levels = "、".join(dim.get("levels", []))
                            ratio = dim.get("ratio", "")
                            lines.append(f"| {dname} | {levels} | {ratio} |")
                    lines.append("")
                elif isinstance(qr, str):
                    lines.append(qr)
                    lines.append("")

            # 实施要点
            if stage.get("procedure_summary"):
                lines.append("#### 实施要点")
                lines.append("")
                for i, step in enumerate(stage["procedure_summary"], 1):
                    lines.append(f"{i}. {step}")
                lines.append("")

            # 预期产出
            if stage.get("expected_outputs"):
                lines.append("#### 预期产出")
                lines.append("")
                for output in stage["expected_outputs"]:
                    lines.append(f"- {output}")
                lines.append("")

    # 四、整体研究设计说明
    if plan.get("overall_rationale"):
        lines.append("## 四、整体研究设计说明")
        lines.append("")
        lines.append(plan["overall_rationale"])
        lines.append("")

    # 五、局限性说明与注意事项
    if plan.get("limitations"):
        lines.append("## 五、局限性说明与注意事项")
        lines.append("")
        for lim in plan["limitations"]:
            lines.append(f"- {lim}")
        lines.append("")

    # 附录
    if plan.get("next_steps"):
        lines.append("---")
        lines.append("")
        lines.append("## 附录一：后续建议")
        lines.append("")
        for i, step in enumerate(plan["next_steps"], 1):
            lines.append(f"{i}. {step}")
        lines.append("")

    generated_at = plan.get("generated_at", datetime.now().strftime("%Y-%m-%d"))
    lines.append("")
    lines.append("---")
    lines.append(f"*方案生成日期：{generated_at}*")

    return "\n".join(lines)


def main():
    if len(sys.argv) < 2:
        print("用法: python generate_plan.py <plan.json> [output.md]", file=sys.stderr)
        print("      python generate_plan.py -  (从 stdin 读取 JSON)", file=sys.stderr)
        sys.exit(1)

    input_path = sys.argv[1]

    if input_path == "-":
        data = json.load(sys.stdin)
    else:
        if not os.path.exists(input_path):
            print(f"文件不存在: {input_path}", file=sys.stderr)
            sys.exit(1)
        with open(input_path, "r", encoding="utf-8") as f:
            data = json.load(f)

    markdown = json_to_markdown(data)

    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(markdown)
        print(f"已输出至: {output_path}")
    else:
        print(markdown)


if __name__ == "__main__":
    main()
