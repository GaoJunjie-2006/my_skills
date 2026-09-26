# -*- coding: utf-8 -*-
"""Debug: write exact paragraph texts to a UTF-8 file for inspection"""
from docx import Document

doc = Document(r"X:\Temp_ws\shiyan4\lab-report-writer\evals\files\测试模板_自动控制原理实验报告.docx")

# Write all paragraph texts to a UTF-8 file
with open(r"X:\Temp_ws\shiyan4\lab-report-writer-workspace\iteration-1\eval-2\with_skill\outputs\paragraph_texts.txt", "w", encoding="utf-8") as f:
    for i, para in enumerate(doc.paragraphs):
        t = para.text
        f.write("P%d: '%s'\n" % (i, t))
        # Also write the repr (with unicode escapes)
        f.write("  repr: %s\n" % repr(t))

    f.write("\n--- TABLE CELLS ---\n")
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                f.write("T%d[%d,%d]: '%s'\n" % (ti, ri, ci, cell.text))

    f.write("\n--- SEARCH FOR KEY TEXTS ---\n")
    targets = [
        "观察各典型环节",
        "阶跃响应",
        "控制系统由各典型环节组成",
        "示波器",
        "实验一",
        "实验报告到此结束",
        "比例环节实验数据",
        "1. 比例环节实验数据",
    ]
    for target in targets:
        for i, para in enumerate(doc.paragraphs):
            if target in para.text:
                f.write("FOUND '%s' in P%d\n" % (target, i))
                break
        else:
            f.write("NOT FOUND '%s' in any paragraph\n" % target)

        # Also check in tables
        for ti, table in enumerate(doc.tables):
            for ri, row in enumerate(table.rows):
                for ci, cell in enumerate(row.cells):
                    if target in cell.text:
                        f.write("FOUND '%s' in T%d[%d,%d]\n" % (target, ti, ri, ci))

print("Written to paragraph_texts.txt")
