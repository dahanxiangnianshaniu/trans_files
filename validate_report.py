#!/usr/bin/env python3
"""
隐私声明/个人数据说明合规检测报告 — 自动校验脚本
用法：python validate_report.py <报告.md路径>

校验项：
1. 章节结构完整性（对照模板）
2. 统计数字加和校验
3. 3.3节8组全覆盖检查
4. 审计概述表4字段完整性
5. 降级模式标注检查
6. △轻重标注检查
"""

import sys
import re
import argparse

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MODE1_REQUIRED_SECTIONS = [
    "# 隐私声明合规检测报告",
    "## 审计概述",
    "## 审计结论",
    "## 一、文档发现",
    "## 二、文档结构识别与数据提取",
    "### 2.1 数据来源结构",
    "### 2.2 空值检测",
    "## 三、个人数据类型四要素检查",
    "### 3.1 四要素矩阵",
    "### 3.2 缺失项详细说明",
    "#### 3.2.1 ✗缺失项",
    "#### 3.2.2 △模糊项",
    "### 3.3 非个人数据识别",
    "## 四、第三方SDK/共享信息检查",
    "## 五、附录",
    "### 审计范围",
    "### 审计说明",
    "### 文档质量异常",
]

MODE2_EXTRA_SECTIONS = [
    "## 四、格式合规性检查",
    "### 4.1 分类结构",
    "### 4.2 边界清晰",
    "### 4.3 数据来源标注",
    "### 4.4 版本与修订记录",
    "## 五、第三方SDK/共享信息检查",
    "## 六、附录",
]

STANDARD_8_GROUPS = [
    "用户身份类",
    "生物特征类",
    "网络标识类",
    "设备标识类",
    "系统配置类",
    "日志记录类",
    "业务数据类",
    "运维配置类",
]

OVERVIEW_FIELDS = ["审计条目", "检查模式", "审计结论", "审计日期"]


def validate_sections(content, mode):
    """校验章节结构完整性"""
    issues = []
    required = MODE1_REQUIRED_SECTIONS[:]
    if mode == 2:
        # 模式二有额外的格式合规检查章节，附录编号变为六
        required = [s for s in required if not s.startswith("## 四") and not s.startswith("## 五")]
        required += MODE2_EXTRA_SECTIONS

    for section in required:
        if section not in content:
            issues.append(f"[章节缺失] 未找到章节: {section}")
    return issues


def validate_overview(content):
    """校验审计概述表4字段"""
    issues = []
    # 提取审计概述到下一个##之间的内容
    overview_match = re.search(r"## 审计概述\s*\n(.*?)(?=\n## |\Z)", content, re.DOTALL)
    if not overview_match:
        issues.append("[审计概述] 未找到审计概述章节")
        return issues

    overview_text = overview_match.group(1)

    # 检查4个必需字段是否出现在表格中
    for field in OVERVIEW_FIELDS:
        if field not in overview_text:
            issues.append(f"[审计概述] 缺少字段: {field}")

    # 检查是否有多余的字段行（表格中"项目"列不应有超出4个的行）
    # 提取所有"项目"列的内容
    rows = re.findall(r"\|\s*([^|\n]+?)\s*\|\s*([^|\n]+?)\s*\|", overview_text)
    field_names = []
    for key, val in rows:
        key = key.strip()
        if key and key != "项目" and key != "------" and key != "---":
            field_names.append(key)

    if len(field_names) > 4:
        extra = [f for f in field_names if f not in OVERVIEW_FIELDS]
        if extra:
            issues.append(f"[审计概述] 存在非标准字段: {extra[:3]}，模板要求仅4字段: {OVERVIEW_FIELDS}")
    return issues


def validate_statistics(content):
    """校验统计数字加和"""
    issues = []
    # 查找四要素状态统计表
    stat_match = re.search(
        r"\*\*四要素状态统计\*\*.*?\|.*?\n\s*\|.*?\n\s*\|\s*✓\s*合规.*?\n\s*\|\s*△\s*模糊.*?\n\s*\|\s*✗\s*不合规.*?\n\s*\|\s*\*?\*?合计\*?\*?.*?\n",
        content, re.DOTALL
    )
    if not stat_match:
        issues.append("[统计] 未找到四要素状态统计表")
        return issues

    stat_text = stat_match.group()
    # 提取数字
    numbers = re.findall(r"\|\s*(\d+)\s*项?\s*\|", stat_text)
    if len(numbers) >= 3:
        try:
            pass_count = int(numbers[0])
            warn_count = int(numbers[1])
            fail_count = int(numbers[2])
            total = int(numbers[3]) if len(numbers) > 3 else pass_count + warn_count + fail_count
            if pass_count + warn_count + fail_count != total:
                issues.append(f"[统计] 加和不一致: {pass_count}+{warn_count}+{fail_count}={pass_count+warn_count+fail_count} ≠ 合计{total}")
        except ValueError:
            issues.append("[统计] 无法解析统计数字")
    return issues


def validate_33_groups(content):
    """校验3.3节8组全覆盖"""
    issues = []
    # 提取3.3节内容
    section_33_match = re.search(r"### 3\.3 非个人数据识别.*?(?=\n## |\Z)", content, re.DOTALL)
    if not section_33_match:
        issues.append("[3.3] 未找到3.3节")
        return issues

    section_33 = section_33_match.group()

    for group in STANDARD_8_GROUPS:
        if group not in section_33:
            issues.append(f"[3.3] 缺少标准组: {group}（即使不涉及也必须显式声明）")

    # 检查是否有"不涉及"声明（至少应有几个不涉及的组）
    not_involved_count = len(re.findall(r"不涉及", section_33))
    # 对于大多数产品文档，8组中至少1-3组不涉及是正常的
    # 但如果没有一个"不涉及"，可能遗漏了显式声明
    if not_involved_count == 0 and "可能非个人数据" not in section_33:
        issues.append("[3.3] 未发现任何'不涉及'声明——8组标准组中通常至少有1组不涉及，可能遗漏了显式声明")

    # 检查边界组是否有3问题讨论
    boundary_groups = ["网络标识类", "设备标识类", "系统配置类", "日志记录类", "运维配置类"]
    has_3questions = any(q in section_33 for q in ["直接识别", "间接识别", "用户行为"])
    if not has_3questions:
        issues.append("[3.3] 边界组缺少3问题讨论（①能否直接识别？②能否间接识别？③是否关联用户行为？）")

    return issues


def validate_degradation(content):
    """校验降级模式标注"""
    issues = []
    audit_note_match = re.search(r"### 审计说明.*?(?=\n### |\n## |\Z)", content, re.DOTALL)
    if not audit_note_match:
        issues.append("[附录] 未找到审计说明章节")
        return issues

    audit_note = audit_note_match.group()
    if "降级模式" not in audit_note and "降级" not in audit_note:
        issues.append("[附录] 审计说明未标注降级模式——无论是否使用MCP工具，都必须标注")
    return issues


def validate_delta_weight(content):
    """校验△轻重标注"""
    issues = []
    # 查找3.2.2节
    section_322_match = re.search(r"#### 3\.2\.2.*?(?=\n### |\n## |\Z)", content, re.DOTALL)
    if not section_322_match:
        return issues

    section_322 = section_322_match.group()
    # 检查是否有△（重）和△（轻）标注
    has_heavy = "△（重）" in section_322 or "△(重)" in section_322
    has_light = "△（轻）" in section_322 or "△(轻)" in section_322

    delta_count = section_322.count("△")
    if delta_count > 0:
        if not has_heavy and not has_light:
            issues.append(f"[3.2.2] △模糊项未标注轻重（共{delta_count}项△，应标注△（重）或△（轻））")
        elif has_heavy and not has_light:
            # 检查△（重）是否排在△（轻）前面
            heavy_pos = section_322.find("△（重）") if "△（重）" in section_322 else section_322.find("△(重)")
            light_pos = section_322.find("△（轻）") if "△（轻）" in section_322 else section_322.find("△(轻)")
            # 如果两者都存在，重应在轻前面
    return issues


def validate_conclusion_consistency(content):
    """校验审计结论与统计数字一致性"""
    issues = []
    # 提取审计结论
    conclusion_match = re.search(r"\|\s*审计结论\s*\|\s*\*?\*?(PASS|FAIL|WARNING)\*?\*?\s*\|", content)
    if not conclusion_match:
        issues.append("[审计结论] 未找到审计结论字段")
        return issues

    conclusion = conclusion_match.group(1)

    # 提取统计数字
    stat_match = re.search(
        r"\*\*四要素状态统计\*\*.*?\|.*?\n\s*\|.*?\n\s*\|\s*✓\s*合规.*?\n\s*\|\s*△\s*模糊.*?\n\s*\|\s*✗\s*不合规.*?\n",
        content, re.DOTALL
    )
    if not stat_match:
        return issues

    numbers = re.findall(r"\|\s*(\d+)\s*项?\s*\|", stat_match.group())
    if len(numbers) >= 3:
        try:
            fail_count = int(numbers[2])
            warn_count = int(numbers[1])
            if conclusion == "PASS" and (fail_count > 0 or warn_count > 0):
                issues.append(f"[结论不一致] 审计结论为PASS但存在{fail_count}项✗和{warn_count}项△")
            elif conclusion == "FAIL" and fail_count == 0:
                issues.append(f"[结论不一致] 审计结论为FAIL但✗项为0（应为WARNING）")
            elif conclusion == "WARNING" and fail_count > 0:
                issues.append(f"[结论不一致] 审计结论为WARNING但存在{fail_count}项✗（应为FAIL）")
        except ValueError:
            pass
    return issues


def main():
    parser = argparse.ArgumentParser(description="隐私声明合规检测报告自动校验")
    parser.add_argument("report_path", help="报告markdown文件路径")
    parser.add_argument("--mode", type=int, choices=[1, 2], default=1, help="检查模式: 1=数据控制者, 2=设备供应者")
    args = parser.parse_args()

    try:
        with open(args.report_path, "r", encoding="utf-8") as f:
            content = f.read()
    except Exception as e:
        print(f"错误: 无法读取文件 {args.report_path}: {e}")
        sys.exit(1)

    print(f"校验报告: {args.report_path}")
    print(f"检查模式: {'数据控制者' if args.mode == 1 else '设备供应者'}")
    print("=" * 60)

    all_issues = []

    all_issues.extend(validate_sections(content, args.mode))
    all_issues.extend(validate_overview(content))
    all_issues.extend(validate_statistics(content))
    all_issues.extend(validate_33_groups(content))
    all_issues.extend(validate_degradation(content))
    all_issues.extend(validate_delta_weight(content))
    all_issues.extend(validate_conclusion_consistency(content))

    if all_issues:
        print(f"\n发现 {len(all_issues)} 个问题：\n")
        for i, issue in enumerate(all_issues, 1):
            print(f"  {i}. {issue}")
        print(f"\n校验结果: FAIL ({len(all_issues)} 个问题)")
        sys.exit(1)
    else:
        print("\n所有校验项通过！")
        print("校验结果: PASS")
        sys.exit(0)


if __name__ == "__main__":
    main()
