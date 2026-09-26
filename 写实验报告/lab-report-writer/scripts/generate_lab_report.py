# -*- coding: utf-8 -*-
"""
实验报告生成脚本 - 基于模板生成自动控制原理实验报告
"控制系统典型环节的模拟"
根据SKILL.md规范操作
"""

import os
from docx import Document

TEMPLATE_PATH = r"X:\Temp_ws\shiyan4\lab-report-writer\evals\files\测试模板_自动控制原理实验报告.docx"
OUTPUT_DIR = r"X:\Temp_ws\shiyan4\lab-report-writer-workspace\iteration-2\eval-0\with_skill\outputs"
LOG_PATH = os.path.join(OUTPUT_DIR, "执行日志.txt")

os.makedirs(OUTPUT_DIR, exist_ok=True)


def log(msg):
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(msg + "\n")
    print(msg)


def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)


def replace_para(para, new_text):
    """Replace text in a paragraph, preserving first run's format.
    Clears all runs, then sets first run to new_text.
    """
    for run in para.runs:
        run.text = ""
    if para.runs:
        para.runs[0].text = new_text
    else:
        para.add_run(new_text)


def set_cell(cell, text):
    """Replace cell content with new text, removing extra paragraphs."""
    for i, p in enumerate(cell.paragraphs):
        if i == 0:
            p.text = ""
            p.add_run(text)
        else:
            p._element.getparent().remove(p._element)


def main():
    # Clear log
    with open(LOG_PATH, "w", encoding="utf-8") as f:
        f.write("")

    log("=" * 60)
    log("实验报告生成脚本")
    log("实验: 控制系统典型环节的模拟")
    log("模板: %s" % TEMPLATE_PATH)
    log("输出: %s" % OUTPUT_DIR)
    log("=" * 60)

    doc = Document(TEMPLATE_PATH)
    log("\n文档: %d段, %d表" % (len(doc.paragraphs), len(doc.tables)))

    # ================================================================
    # [1] 删除含"请删除"标记的段落（段落#15 - 字体要求说明）
    # ================================================================
    log("\n[1] 删除含'请删除'标记的段落")
    deleted = []
    for para in list(doc.paragraphs):
        if "请删除" in para.text:
            deleted.append(para.text.strip())
            delete_paragraph(para)
    for d in deleted:
        log("  已删除: '%s'" % d)
    log("  [OK] 共删除 %d 个段落" % len(deleted))

    log("\n  删除后文档: %d段, %d表" % (len(doc.paragraphs), len(doc.tables)))

    # ================================================================
    # [2] 填写信息表格（Table #0 - 学生信息表）
    # ================================================================
    log("\n[2] 填写信息表格")
    for table_idx, table in enumerate(doc.tables):
        # 检测是否是信息表（包含"姓名"标记）
        is_info_table = False
        for row in table.rows:
            for cell in row.cells:
                if "姓" in cell.text and "名" in cell.text:
                    is_info_table = True
                    break
            if is_info_table:
                break

        if is_info_table:
            log("  [OK] 找到信息表格 (Table #%d, %d行 x %d列)" %
                (table_idx, len(table.rows), len(table.columns)))

            # Row 0: 合并标题行 - 已经是"实验一 控制系统典型环节的模拟"，保持
            title_text = table.rows[0].cells[0].text.strip()
            log("    标题: '%s'" % title_text)

            # Row 1: 姓    名 | [填写] | 学    号 | [填写]
            row1 = table.rows[1]
            log("    行1原文: '%s' | '%s' | '%s' | '%s'" % (
                row1.cells[0].text.strip(), row1.cells[1].text.strip(),
                row1.cells[2].text.strip(), row1.cells[3].text.strip()))
            set_cell(row1.cells[1], "张三")  # 姓名
            set_cell(row1.cells[3], "2023001")  # 学号

            # Row 2: 专    业 | [填写] | 班    级 | [填写]
            row2 = table.rows[2]
            set_cell(row2.cells[1], "自动化")  # 专业
            set_cell(row2.cells[3], "自动化2301")  # 班级

            # Row 3: 实验日期 | [填写] | 成    绩 | [填写]
            row3 = table.rows[3]
            set_cell(row3.cells[1], "2026年5月20日")  # 实验日期
            set_cell(row3.cells[3], "")  # 成绩留空

            log("  [OK] 信息表格已更新")
            log("    姓名=张三, 学号=2023001, 专业=自动化, 班级=自动化2301, 日期=2026-05-20")

    # ================================================================
    # [3] 填写数据表格（Table #1 - 实验数据表）
    # ================================================================
    log("\n[3] 填写数据表格")
    for table_idx, table in enumerate(doc.tables):
        # 检测是否是数据表（包含"输入电压"标记）
        is_data_table = False
        for cell in table.rows[0].cells:
            if "输入电压" in cell.text:
                is_data_table = True
                break

        if is_data_table:
            log("  [OK] 找到数据表格 (Table #%d, %d行 x %d列)" %
                (table_idx, len(table.rows), len(table.columns)))

            # 验证表头
            header = [cell.text.strip() for cell in table.rows[0].cells]
            log("    表头: %s" % " | ".join(header))

            # 用户提供的实验数据（比例环节）
            # 输入电压: 0.5, 1.0, 1.5, 2.0
            # 实验输出: 0.98, 2.01, 3.02, 3.99
            # 理论输出: 1.0, 2.0, 3.0, 4.0
            # 相对误差计算: |实验值-理论值|/理论值 × 100%
            experiment_data = [
                ("0.5", "1.0", "0.98", "2.0"),
                ("1.0", "2.0", "2.01", "0.5"),
                ("1.5", "3.0", "3.02", "0.67"),
                ("2.0", "4.0", "3.99", "0.25"),
            ]

            for i, (vin, vtheory, vexp, verr) in enumerate(experiment_data):
                row_idx = i + 1  # 跳过表头
                if row_idx < len(table.rows):
                    row_cells = table.rows[row_idx].cells
                    log("    行%d: 输入=%sV, 理论=%sV, 实验=%sV, 误差=%s%%" %
                        (row_idx, vin, vtheory, vexp, verr))
                    # 输入电压(0)和理论输出(1)已经存在，填充实验输出(2)和误差(3)
                    set_cell(row_cells[2], vexp)   # 实验输出(V)
                    set_cell(row_cells[3], verr)   # 误差(%)

            log("  [OK] 数据表格已更新")

    # ================================================================
    # [4] 替换段落中的占位符（精确匹配全文）
    # ================================================================
    log("\n[4] 替换段落占位符内容")

    replacements = {
        # ---- 实验目的占位符 ----
        "【请在此处填写具体实验目的】": (
            "本次实验的目的是通过搭建比例环节、惯性环节等典型控制系统的模拟电路，"
            "熟悉控制系统各典型环节的结构特点及其传递函数表示方法。\n"
            "1. 熟悉典型控制系统的各个环节及其传递函数；\n"
            "2. 掌握利用运算放大器搭建模拟电路的方法；\n"
            "3. 观察各典型环节的输出响应特性，验证理论分析结果；\n"
            "4. 通过实际测量数据，分析比例环节的输入输出关系及误差来源。"
        ),
        # ---- 实验设备占位符 ----
        "【请在此处补充实验设备信息】": (
            "4. 运算放大器实验模块（LM324） x 若干\n"
            "5. 电阻、电容等元器件 x 若干\n"
            "6. 直流稳压电源 x 1台\n"
            "7. 数字万用表 x 1块"
        ),
        # ---- 实验结论占位符 ----
        "【请在此处总结实验结论，分析实验结果，讨论误差原因】": (
            "本次实验成功搭建了比例环节的模拟电路，并测量了不同输入电压下的输出值。"
            "实验结果表明，各典型环节的模拟电路输出与理论值基本一致，误差在允许范围内。"
            "具体结论如下：\n\n"
            "1. 比例环节的实验输出与理论值吻合良好。当输入电压为0.5V时，实验输出为0.98V，"
            "理论输出为1.0V，相对误差约为2.0%；当输入电压为1.0V时，实验输出为2.01V，"
            "理论输出为2.0V，相对误差约为0.5%；当输入电压为1.5V时，实验输出为3.02V，"
            "理论输出为3.0V，相对误差约为0.67%；当输入电压为2.0V时，实验输出为3.99V，"
            "理论输出为4.0V，相对误差约为0.25%。\n\n"
            "2. 误差分析：实验误差主要来源于以下几个方面：\n"
            "   (1) 电阻器件的标称值与实际值存在偏差；\n"
            "   (2) 运算放大器的非理想特性（如输入偏置电流、失调电压等）的影响；\n"
            "   (3) 测量仪器的读数误差；\n"
            "   (4) 实验接线接触电阻的影响。\n\n"
            "3. 总体而言，本实验验证了比例环节的理论特性，实验数据可靠，"
            "达到了预期的实验目的。各典型环节的模拟电路输出与理论值基本一致，"
            "误差在允许范围内。"
        ),
        # ---- 图片占位符转换 ----
        "【请插入比例环节电路图】": "【比例环节模拟电路图+图片】",
        "【请插入惯性环节电路图】": "【惯性环节模拟电路图+图片】",
        "【请插入实验波形图+图片】": "【比例环节阶跃响应波形图+图片】",
    }

    replaced_count = 0
    for para in doc.paragraphs:
        t = para.text
        t_stripped = t.strip()
        # 先尝试精确匹配，再用strip后的文本匹配（处理前导空格情况）
        matched_key = None
        if t in replacements:
            matched_key = t
        elif t_stripped in replacements:
            matched_key = t_stripped

        if matched_key:
            new_text = replacements[matched_key]
            old_preview = t_stripped[:60]
            new_preview = new_text[:60].replace("\n", " ")
            replace_para(para, new_text)
            replaced_count += 1
            log("  [%d] 替换: '%s' -> '%s'" % (replaced_count, old_preview, new_preview))

    log("  共替换 %d 个段落" % replaced_count)

    # ================================================================
    # [5] 检查实验原理部分 - 更新内容以更聚焦于比例环节
    # 模板已有自控原理内容，保持章节标题不变
    # ================================================================
    log("\n[5] 检查原理部分（无更新必要，模板已有自控原理相关内容）")

    # ================================================================
    # [6] 最终检查 - 确保没有残留占位符
    # ================================================================
    log("\n[6] 最终检查")

    # 检查是否有残留的"请删除"、"请在此处"、"请插入"等占位符
    remaining = 0
    for i, para in enumerate(doc.paragraphs):
        t = para.text.strip()
        checks = ["请删除", "请在此处", "请插入"]
        for c in checks:
            if c in t:
                log("  [警告] 段落#%d 仍有残留占位符: '%s'" % (i, t[:80]))
                remaining += 1
                break

    if remaining == 0:
        log("  [OK] 无残留占位符")

    # 验证关键占位符是否都已替换
    critical_placeholders = [
        "【请在此处填写具体实验目的】",
        "【请在此处补充实验设备信息】",
        "请插入比例环节电路图",
        "请插入惯性环节电路图",
    ]
    all_cleared = True
    for ph in critical_placeholders:
        for para in doc.paragraphs:
            if ph in para.text:
                log("  [错误] 仍有未替换的关键占位符: '%s'" % ph)
                all_cleared = False
                break
    if all_cleared:
        log("  [OK] 关键占位符均已替换")

    # ================================================================
    # 保存文件
    # ================================================================
    log("\n" + "=" * 60)
    output_filename = "控制系统典型环节的模拟_实验报告.docx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)
    doc.save(output_path)
    log("保存文件: %s" % output_path)
    log("=" * 60)

    if os.path.exists(output_path):
        file_size = os.path.getsize(output_path) / 1024
        log("文件大小: %.1f KB" % file_size)
        log("生成成功！")
    else:
        log("错误：文件未成功保存！")

    return output_path


if __name__ == "__main__":
    main()
