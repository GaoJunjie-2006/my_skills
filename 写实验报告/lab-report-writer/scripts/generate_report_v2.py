# -*- coding: utf-8 -*-
"""
实验报告生成脚本 v2 - 基于模板生成数电实验报告
使用精确文本匹配替换模板内容
"""

import os
import sys
from docx import Document

TEMPLATE_PATH = r"X:\Temp_ws\shiyan4\lab-report-writer\evals\files\测试模板_自动控制原理实验报告.docx"
OUTPUT_DIR = r"X:\Temp_ws\shiyan4\lab-report-writer-workspace\iteration-2\eval-1\with_skill\outputs"
LOG_PATH = os.path.join(OUTPUT_DIR, "生成日志.txt")

# Clear log
with open(LOG_PATH, "w", encoding="utf-8") as f:
    f.write("")


def log(msg):
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(msg + "\n")
    # Also try stdout (may fail for Chinese but that's ok)
    try:
        print(msg)
    except:
        pass


def delete_paragraph(paragraph):
    """Delete a paragraph element from the document."""
    p = paragraph._element
    p.getparent().remove(p)


def replace_para(para, new_text):
    """Replace all text in a paragraph while preserving formatting of first run."""
    for run in para.runs:
        run.text = ""
    if para.runs:
        para.runs[0].text = new_text
    else:
        para.add_run(new_text)


def set_cell(cell, text):
    """Replace all text in a table cell."""
    # First, try to clear by setting cell.text (handles merged cells better)
    # Then set first paragraph text directly via runs
    p = cell.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(text)
    # Remove extra paragraphs
    for pi in range(1, len(cell.paragraphs)):
        cell.paragraphs[pi]._element.getparent().remove(cell.paragraphs[pi]._element)


def clear_para(para):
    """Clear all text from a paragraph."""
    for run in para.runs:
        run.text = ""


def main():
    log("=" * 60)
    log("实验报告生成脚本 v2")
    log("模板: %s" % TEMPLATE_PATH)
    log("输出: %s" % OUTPUT_DIR)
    log("=" * 60)

    doc = Document(TEMPLATE_PATH)
    log("\n文档: %d段, %d表" % (len(doc.paragraphs), len(doc.tables)))

    # ======================== [0] Check exact text of all paragraphs ========================
    # We need to know exact text for matching, let's dump and read back
    log("\n[0] 段落内容检查 (repr)")
    for i, para in enumerate(doc.paragraphs):
        t = para.text
        if t.strip() or para.runs:
            log("  P%d: len=%d repr=%r" % (i, len(t), t[:80]))

    # ======================== [1] 删除段落P15（字体说明，含"请删除"标记）=======================
    log("\n[1] 删除含'请删除'标记的段落")

    # Find and delete paragraphs containing "请删除"
    paras_to_delete = []
    for i, para in enumerate(doc.paragraphs):
        if '请删除' in para.text:
            paras_to_delete.append((i, para))
            log("  找到P%d: '%s'" % (i, para.text.strip()[:60]))

    for idx, (i, para) in enumerate(paras_to_delete):
        delete_paragraph(para)
        log("  [OK] 已删除P%d (%d/%d)" % (i, idx+1, len(paras_to_delete)))

    # ======================== [2] 主标题 ========================
    log("\n[2] 修改主标题")
    # The first non-empty paragraph is the title
    for para in doc.paragraphs:
        if para.text.strip():
            old = para.text.strip()
            replace_para(para, "集成触发器逻辑功能测试实验报告")
            log("  [OK] '%s' -> '集成触发器逻辑功能测试实验报告'" % old[:40])
            break

    # ======================== [3] 处理信息表格 (Table 0) ========================
    log("\n[3] 填写信息表格")
    info_table = doc.tables[0]  # First table is the info table

    # Row0: 实验名称 (标题行，4列gridSpan共享)
    log("  Row0 - 设置实验名称")
    info_table.rows[0].cells[0].text = "集成触发器逻辑功能测试"
    # NOTE: 不要修改cells[1-3]，它们与cell[0]共享gridSpan=4内容

    # Row1: 姓名、学号
    log("  Row1 - 设置姓名学号")
    set_cell(info_table.rows[1].cells[0], "姓    名")
    set_cell(info_table.rows[1].cells[1], "李四")
    set_cell(info_table.rows[1].cells[2], "学    号")
    set_cell(info_table.rows[1].cells[3], "2023002")

    # Row2: 专业、班级
    log("  Row2 - 设置专业班级")
    set_cell(info_table.rows[2].cells[0], "专    业")
    set_cell(info_table.rows[2].cells[1], "电子信息工程")
    set_cell(info_table.rows[2].cells[2], "班    级")
    set_cell(info_table.rows[2].cells[3], "电子2301")

    # Row3: 实验日期、成绩
    log("  Row3 - 设置实验日期")
    set_cell(info_table.rows[3].cells[0], "实验日期")
    set_cell(info_table.rows[3].cells[1], "2026年5月22日")
    set_cell(info_table.rows[3].cells[2], "成    绩")
    set_cell(info_table.rows[3].cells[3], "")
    log("  [OK] 信息表格已更新")

    # ======================== [4] 替换数据表格 (Table 1) ========================
    log("\n[4] 替换数据表格")
    data_table = doc.tables[1]

    # Update headers for flip-flop testing data
    set_cell(data_table.rows[0].cells[0], "触发器类型")
    set_cell(data_table.rows[0].cells[1], "输入条件")
    set_cell(data_table.rows[0].cells[2], "输出Q")
    set_cell(data_table.rows[0].cells[3], "功能说明")

    # JK Flip-flop data
    set_cell(data_table.rows[1].cells[0], "JK触发器")
    set_cell(data_table.rows[1].cells[1], "J=1, K=0")
    set_cell(data_table.rows[1].cells[2], "Q=1")
    set_cell(data_table.rows[1].cells[3], "置位")

    set_cell(data_table.rows[2].cells[0], "JK触发器")
    set_cell(data_table.rows[2].cells[1], "J=0, K=1")
    set_cell(data_table.rows[2].cells[2], "Q=0")
    set_cell(data_table.rows[2].cells[3], "复位")

    set_cell(data_table.rows[3].cells[0], "JK触发器")
    set_cell(data_table.rows[3].cells[1], "J=1, K=1")
    set_cell(data_table.rows[3].cells[2], "Q翻转")
    set_cell(data_table.rows[3].cells[3], "翻转")

    set_cell(data_table.rows[4].cells[0], "D触发器")
    set_cell(data_table.rows[4].cells[1], "D=1")
    set_cell(data_table.rows[4].cells[2], "Q=1")
    set_cell(data_table.rows[4].cells[3], "跟随输入")

    log("  [OK] 数据表格已更新")

    # ======================== [5] 替换段落内容 ========================
    log("\n[5] 替换段落内容")

    # Define replacements as a dict: old_text -> new_text
    # We need to match the exact text from template. Since terminal shows garbled,
    # we match by using known strings.

    REPLACEMENTS = {
        # 实验目的条目 (P3-P5)
        "1. 熟悉典型控制系统的各个环节及其传递函数。":
            "1. 掌握JK触发器的逻辑功能测试方法及其真值表验证。",
        "2. 掌握利用运算放大器搭建模拟电路的方法。":
            "2. 掌握D触发器的逻辑功能测试方法及其真值表验证。",
        "3. 观察各典型环节的输出响应特性。":
            "3. 理解触发器的置位、复位和翻转功能，掌握时序电路的基本测试方法。",

        # 实验原理内容 (P8)
        "控制系统由各种典型环节组成，常见的典型环节包括比例环节、惯性环节、积分环节、微分环节等。":
            "触发器是数字电路中的基本时序逻辑单元，具有记忆功能，能够存储一位二进制信息。"
            "本次实验主要研究JK触发器和D触发器两种常见的触发器类型。",

        "各环节的传递函数如下：":
            "JK触发器和D触发器的逻辑功能如下：",

        # 公式行
        "（1）比例环节：G(s)=K":
            "（1）JK触发器：当J=1,K=0时，触发器置位（Q=1）；当J=0,K=1时，触发器复位（Q=0）；"
            "当J=1,K=1时，触发器状态翻转；当J=0,K=0时，触发器保持原状态。",

        "（2）惯性环节：G(s)=K/(Ts+1)":
            "（2）JK触发器特性方程：Q^(n+1) = J*Q' + K'*Q",

        "（3）积分环节：G(s)=K/s":
            "（3）D触发器：当D=1时，输出Q=1；当D=0时，输出Q=0。输出跟随输入变化。",

        "（4）微分环节：G(s)=K·s":
            "（4）D触发器特性方程：Q^(n+1) = D",

        # 设备条目
        "1. 计算机 x 1台":
            "1. 数字电路实验箱 x 1台",
        "2. MATLAB/Simulink 仿真软件 x 1套":
            "2. 双踪示波器 x 1台",
        "3. 数据采集卡 x 1块":
            "3. 直流稳压电源 x 1台",

        # 实验步骤
        "1. 搭建比例环节的模拟电路":
            "1. 将JK触发器芯片（如74LS76）插入实验箱，连接电源和地线",
        "2. 搭建惯性环节的模拟电路":
            "2. 将D触发器芯片（如74LS74）插入实验箱，连接电源和地线",
        "3. 分别输入阶跃信号，观察并记录输出波形":
            "3. 分别设置J、K、D、CLK等输入信号，观察并记录输出Q的状态",

        # 子标题
        "步骤一：比例环节实验":
            "步骤一：JK触发器逻辑功能测试",
        "步骤二：惯性环节实验":
            "步骤二：D触发器逻辑功能测试",

        # 图片引导句
        "  连接电路图如下所示：":
            "  实验电路连接图如下所示：",

        # 图片占位符
        "  【请插入比例环节电路图】":
            "【JK触发器测试电路连接图+图片】",
        "  【请插入惯性环节电路图】":
            "【D触发器测试电路连接图+图片】",
        "【请插入实验波形图+图片】":
            "【触发器输出波形图+图片】",

        # 图标题
        "图1 比例环节阶跃响应曲线":
            "图1 JK触发器输出波形图",

        # 数据区域
        "1. 比例环节实验结果":
            "1. JK触发器测试结果",
        "表1 比例环节实验数据":
            "表1 JK触发器和D触发器功能测试数据",

        # 结束语
        "— 实验报告结束 —":
            "— 实验报告结束 —",
    }

    # 占位符内容（需要替换为实际内容）
    PLACEHOLDER_CONTENT = {
        "【请在此处填写具体实验目的】": (
            "本次实验的目的是通过测试JK触发器和D触发器的逻辑功能，"
            "掌握触发器的基本工作原理和测试方法。具体包括：\n"
            "1. 掌握JK触发器在J=1,K=0（置位）、J=0,K=1（复位）、J=1,K=1（翻转）"
            "等输入条件下的逻辑功能；\n"
            "2. 掌握D触发器在D=1和D=0条件下的逻辑功能；\n"
            "3. 验证触发器的真值表，理解时序逻辑电路的基本特性。"
        ),
        "【请在此处补充实验设备信息】": (
            "4. JK触发器芯片（74LS76） x 1片\n"
            "5. D触发器芯片（74LS74） x 1片\n"
            "6. 逻辑电平开关 x 若干\n"
            "7. 逻辑电平指示灯 x 若干\n"
            "8. 连接导线 x 若干"
        ),
        "【请在此处总结实验结论，分析实验结果，讨论误差原因】": (
            "本次实验完成了JK触发器和D触发器的逻辑功能测试，实验结果与理论真值表一致。\n\n"
            "对于JK触发器：\n"
            "（1）当J=1、K=0时，输出Q=1，触发器处于置位状态，与理论一致。\n"
            "（2）当J=0、K=1时，输出Q=0，触发器处于复位状态，与理论一致。\n"
            "（3）当J=1、K=1时，输出Q发生翻转，触发器处于翻转状态，与理论一致。\n\n"
            "对于D触发器：\n"
            "（1）当D=1时，输出Q=1，输出跟随输入。\n"
            "（2）当D=0时，输出Q=0，输出跟随输入。\n\n"
            "结论：实验结果与真值表一致，JK触发器和D触发器的逻辑功能符合预期。"
            "通过本次实验，掌握了触发器的基本测试方法，加深了对时序逻辑电路的理解。"
        ),
    }

    # Merge all replacements
    all_replacements = {}
    all_replacements.update(REPLACEMENTS)
    all_replacements.update(PLACEHOLDER_CONTENT)

    # Since terminal encoding may cause issues with literal string matching,
    # let's match by finding text that CONTAINS the key rather than exact match
    # Actually let's try exact match first with the literal text from the doc

    # Execute replacement on paragraphs
    replaced_count = 0
    for para in doc.paragraphs:
        t = para.text
        if t in all_replacements:
            old_short = t[:60]
            new_text = all_replacements[t]
            replace_para(para, new_text)
            replaced_count += 1
            log("  [%d] 已替换: '%s' -> '%s'" % (replaced_count, old_short, new_text[:60]))
        else:
            # Try matching trimmed version
            tt = t.strip()
            if tt in all_replacements and tt != t:
                new_text = all_replacements[tt]
                replace_para(para, new_text)
                replaced_count += 1
                log("  [%d] 已替换(trim): '%s' -> '%s'" % (replaced_count, tt[:60], new_text[:60]))

    log("  共替换 %d 段" % replaced_count)

    # ======================== [6] 处理表格中的实验名称 ========================
    log("\n[6] 处理表格中的实验项目标题")
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if '实验' in cell.text and ('控制系统' in cell.text or '典型环节' in cell.text):
                    set_cell(cell, "集成触发器逻辑功能测试")
                    log("  [OK] 替换表格中实验项目标题")

    # ======================== [7] 替换所有"请在此处"风格的占位符（备选方案）=======================
    log("\n[7] 检查并替换残留占位符")
    for para in doc.paragraphs:
        t = para.text.strip()
        if '请在此处填写具体实验目的' in t:
            replace_para(para, PLACEHOLDER_CONTENT["【请在此处填写具体实验目的】"])
            log("  [补充] 替换实验目的占位符")
        elif '请在此处补充实验设备信息' in t:
            replace_para(para, PLACEHOLDER_CONTENT["【请在此处补充实验设备信息】"])
            log("  [补充] 替换实验设备占位符")
        elif '请在此处总结实验结论' in t:
            replace_para(para, PLACEHOLDER_CONTENT["【请在此处总结实验结论，分析实验结果，讨论误差原因】"])
            log("  [补充] 替换实验结论占位符")

    # ======================== [8] 更新实验原理章节标题及其内容 ========================
    log("\n[8] 检查并更新章节标题")
    # Already handled by exact text matching above

    # ======================== [9] 最终检查 ========================
    log("\n[9] 最终检查")
    remaining = 0
    for i, para in enumerate(doc.paragraphs):
        t = para.text.strip()
        if '请在此处' in t or '请插入' in t or '请删除' in t:
            log("  [注意] P%d 可能还有占位符: '%s'" % (i, t[:60]))
            remaining += 1
    if remaining == 0:
        log("  [OK] 无残留占位符")

    # Also check for empty image placeholders that need attention
    for i, para in enumerate(doc.paragraphs):
        t = para.text.strip()
        if '【' in t and '图片' in t:
            log("  [图片占位符] P%d: '%s'" % (i, t[:60]))

    # ======================== 保存 ========================
    log("\n" + "=" * 60)
    output_filename = "集成触发器逻辑功能测试实验报告.docx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)
    doc.save(output_path)
    log("文件: %s" % output_path)
    log("=" * 60)
    if os.path.exists(output_path):
        log("大小: %.1f KB" % (os.path.getsize(output_path) / 1024))
        log("生成成功！")
    else:
        log("错误：文件未成功保存！")

    return output_path


if __name__ == "__main__":
    main()
