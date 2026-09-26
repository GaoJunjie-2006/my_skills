"""
实验报告模板分析工具
用法: python analyze_template.py <模板文件路径>

输出模板的完整结构分析，包括:
- 文档概览（段落数、表格数、节数）
- 每段内容、样式、字体格式
- 表格结构（行列、合并单元格、内容）
- 图片位置
- 特殊标注（如"请删除"等提示）
"""

import sys
import os
from docx import Document
from docx.oxml.ns import qn


def analyze_font(run):
    """分析一个run的字体信息"""
    info = {}
    if run.font.name:
        info['字体'] = run.font.name
        rPr = run._element.find(qn('w:rPr'))
        if rPr is not None:
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is not None:
                ea = rFonts.get(qn('w:eastAsia'))
                if ea:
                    info['东亚字体'] = ea
    if run.font.size:
        size_pt = run.font.size.pt
        cname = ''
        mapping = {36: '小初', 26: '一号', 24: '小一', 22: '二号',
                   18: '小二', 16: '三号', 15: '小三', 14: '四号',
                   12: '小四', 10.5: '五号', 9: '小五'}
        cname = mapping.get(size_pt, '')
        info['字号'] = f'{size_pt}pt ({cname})' if cname else f'{size_pt}pt'
    if run.font.bold:
        info['加粗'] = '是'
    if run.font.italic:
        info['斜体'] = '是'
    if run.font.underline:
        info['下划线'] = '是'
    return info


def analyze_paragraph(para, index):
    """分析一个段落"""
    info = {
        '索引': index,
        '样式': para.style.name if para.style else '无',
        '文本': para.text,
        '对齐': str(para.alignment) if para.alignment else '默认',
    }

    runs_info = []
    for run in para.runs:
        run_data = {
            '文本': run.text,
            '字体信息': analyze_font(run)
        }
        runs_info.append(run_data)

    if runs_info:
        info['Run详情'] = runs_info

    images = para._element.findall('.//' + qn('w:drawing'))
    if images:
        info['包含图片'] = True

    text = para.text
    markers = []
    if '请删除' in text:
        markers.append('[注意] 包含[请删除]标记')
    if '删除' in text and ('说明' in text or '提示' in text):
        markers.append('[注意] 疑似模板说明')
    if '请在此' in text or '请输入' in text:
        markers.append('[提示] 填写提示占位')
    if '字体' in text and ('要求' in text or '格式' in text or '说明' in text):
        if '请删除' in text:
            markers.append('[注意] 字体要求说明（含删除标记）')
    if markers:
        info['标记'] = markers

    return info


def analyze_table(table, table_index):
    """分析一个表格"""
    info = {
        '表格索引': table_index,
        '行数': len(table.rows),
        '列数': len(table.columns),
    }

    header_row = table.rows[0]
    headers = []
    for cell in header_row.cells:
        headers.append(cell.text.strip())
    info['表头'] = headers

    rows_data = []
    for i, row in enumerate(table.rows):
        row_data = []
        for j, cell in enumerate(row.cells):
            cell_text = cell.text.strip()
            cell_info = {
                '内容': cell_text,
                '段落数': len(cell.paragraphs),
            }
            cell_images = cell._element.findall('.//' + qn('w:drawing'))
            if cell_images:
                cell_info['包含图片'] = True
            nested_tables = cell._element.findall('.//' + qn('w:tbl'))
            if nested_tables:
                cell_info['嵌套表格'] = len(nested_tables)
            row_data.append(cell_info)
        rows_data.append({'行号': i, '单元格': row_data})

    info['行数据'] = rows_data
    return info


def main():
    if len(sys.argv) < 2:
        print("用法: python analyze_template.py <模板文件路径>")
        sys.exit(1)

    filepath = sys.argv[1]
    if not os.path.exists(filepath):
        print(f"错误: 文件不存在: {filepath}")
        sys.exit(1)

    doc = Document(filepath)

    print("=" * 70)
    print(" 模板分析报告")
    print(f" 文件: {os.path.basename(filepath)}")
    print(f" 路径: {os.path.abspath(filepath)}")
    print("=" * 70)

    total_paragraphs = len(doc.paragraphs)
    total_tables = len(doc.tables)
    total_sections = len(doc.sections)

    print(f"\n[文档概览]")
    print(f"  段落数: {total_paragraphs}")
    print(f"  表格数: {total_tables}")
    print(f"  节数:   {total_sections}")

    # 文档结构摘要（仅非空段落）
    print(f"\n[文档结构摘要]")
    print("-" * 70)
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        style = para.style.name if para.style else 'Normal'
        if text or 'Heading' in style:
            marker = ""
            if '请删除' in text:
                marker = "  <<< [含删除标记]"
            print(f"  [{style}] {text[:80]}{marker}")

    # 详细段落分析
    print(f"\n[详细段落分析]")
    print("-" * 70)
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text and not para.runs:
            continue

        para_info = analyze_paragraph(para, i)
        print(f"\n--- 段落 #{i} ---")
        print(f"  样式: {para_info['样式']}")
        if text:
            print(f"  内容: {text[:120]}")
        else:
            print(f"  内容: (空段落, 包含格式化的run)")

        if '标记' in para_info:
            for m in para_info['标记']:
                print(f"  {m}")

        if '包含图片' in para_info:
            print(f"  [图片] 此段落包含图片")

        if para_info.get('Run详情'):
            has_font_info = any(run['字体信息'] for run in para_info['Run详情'] if run['字体信息'])
            if has_font_info:
                print(f"  [字体信息]:")
                for run in para_info['Run详情']:
                    if run['字体信息']:
                        font_str = ", ".join(f"{k}={v}" for k, v in run['字体信息'].items())
                        print(f"    [{font_str}] {run['文本'][:50]}")

    # 表格分析
    if total_tables > 0:
        print(f"\n[表格分析]")
        print("-" * 70)
        for i, table in enumerate(doc.tables):
            print(f"\n--- 表格 #{i} ({len(table.rows)}行 x {len(table.columns)}列) ---")
            for j, row in enumerate(table.rows):
                for k, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()
                    if cell_text:
                        print(f"  [{j},{k}] {cell_text[:100]}")
                    cell_images = cell._element.findall('.//' + qn('w:drawing'))
                    if cell_images:
                        print(f"  [{j},{k}] [图片] 包含图片")
    else:
        print(f"\n[表格分析] 无表格")

    # 节/页面设置分析
    print(f"\n[页面设置]")
    print("-" * 70)
    for i, section in enumerate(doc.sections):
        page_width = section.page_width
        page_height = section.page_height
        if page_width and page_height:
            w_in = page_width / 914400
            h_in = page_height / 914400
            print(f"  节 #{i}: 页面 {w_in:.1f}in x {h_in:.1f}in")
        margins = []
        if section.top_margin:
            margins.append(f"上={section.top_margin/914400:.1f}in")
        if section.bottom_margin:
            margins.append(f"下={section.bottom_margin/914400:.1f}in")
        if section.left_margin:
            margins.append(f"左={section.left_margin/914400:.1f}in")
        if section.right_margin:
            margins.append(f"右={section.right_margin/914400:.1f}in")
        if margins:
            print(f"  页边距: {', '.join(margins)}")

    # 检测可删除的内容
    print(f"\n[可删除内容检测]")
    print("-" * 70)
    found_deletable = False
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if '请删除' in text:
            found_deletable = True
            print(f"  段落 #{i}: \"{text[:100]}\"")
    if not found_deletable:
        print('  (未检测到含"请删除"标记的内容)')

    # 模板占位符检测
    print(f"\n[模板占位符/填写提示检测]")
    print("-" * 70)
    found_placeholder = False
    placeholders = ['请在此', '请输入', '请填写', '在此处', '______', '……', 'XXX', 'xxx']
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        for ph in placeholders:
            if ph in text:
                found_placeholder = True
                print(f"  段落 #{i}: \"{text[:120]}\"")
                break
    if not found_placeholder:
        print("  (未检测到明显占位符)")

    # 图片位置汇总
    print(f"\n[图片位置汇总]")
    print("-" * 70)
    image_count = 0
    for i, para in enumerate(doc.paragraphs):
        images = para._element.findall('.//' + qn('w:drawing'))
        if images:
            image_count += len(images)
            print(f"  段落 #{i}: 包含 {len(images)} 张图片 - 上下文: \"{para.text[:60]}\"")
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                images = cell._element.findall('.//' + qn('w:drawing'))
                if images:
                    image_count += len(images)
                    context = cell.text.strip()[:40]
                    print(f"  表格 #{t_idx}[{r_idx},{c_idx}]: 包含 {len(images)} 张图片 - 上下文: \"{context}\"")
    if image_count == 0:
        print("  (模板中未检测到图片)")
    else:
        print(f"  => 共检测到 {image_count} 张图片位置，将使用【描述+图片】占位符替代")

    print("\n" + "=" * 70)
    print("分析完成")


if __name__ == '__main__':
    main()
