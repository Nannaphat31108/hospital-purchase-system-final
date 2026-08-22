from docx.shared import Cm

def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    # 1. Update meta table widths
    old_meta = """    meta = doc.add_table(rows=1, cols=2)
    meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta.autofit = True"""
    new_meta = """    meta = doc.add_table(rows=1, cols=2)
    meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta.autofit = False
    meta.columns[0].width = Cm(9.5)
    meta.columns[1].width = Cm(8.0)
    for row in meta.rows:
        row.cells[0].width = Cm(9.5)
        row.cells[1].width = Cm(8.0)"""
    code = code.replace(old_meta, new_meta)

    # 2. Update main table widths
    old_table_widths = """    for i, value in enumerate(headers):
        _set_cell_text(table.rows[0].cells[i], value, True, WD_ALIGN_PARAGRAPH.CENTER)
    for line in purchase.lines:"""
    new_table_widths = """    table.autofit = False
    widths = [Cm(1.3), Cm(4.5), Cm(2.9), Cm(2.9), Cm(3.1), Cm(2.9)]
    for i in range(6):
        table.columns[i].width = widths[i]
    for i, value in enumerate(headers):
        _set_cell_text(table.rows[0].cells[i], value, True, WD_ALIGN_PARAGRAPH.CENTER)
        table.rows[0].cells[i].width = widths[i]
    for line in purchase.lines:"""
    code = code.replace(old_table_widths, new_table_widths)
    
    # Update cell widths in the loop
    old_cell_loop = """        for i, value in enumerate(values):
            align = WD_ALIGN_PARAGRAPH.LEFT if i == 1 else (WD_ALIGN_PARAGRAPH.RIGHT if i >= 4 else WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text(cells[i], value, align=align)"""
    new_cell_loop = """        for i, value in enumerate(values):
            cells[i].width = widths[i]
            align = WD_ALIGN_PARAGRAPH.LEFT if i == 1 else (WD_ALIGN_PARAGRAPH.RIGHT if i >= 4 else WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text(cells[i], value, align=align)"""
    code = code.replace(old_cell_loop, new_cell_loop)

    # 3. Update remarks tabs
    old_remarks = """    _paragraph(doc, "๑. การติดอากรแสตมป์ให้เป็นไปตามประมวลกฎหมายรัษฎากร หากต้องการให้ใบสั่งซื้อมีผลตามกฎหมาย")
    _paragraph(doc, f"๒. ใบสั่งซื้อนี้อ้างอิงตามเลขที่โครงการ {purchase.project_number or '........................'} ซื้อพัสดุจำนวน {len(purchase.lines)} รายการ เป็นเงิน {purchase.total_amount:,.2f} บาท ({baht_text(purchase.total_amount)}) โดยวิธีเฉพาะเจาะจง")"""
    new_remarks = """    _paragraph(doc, "\t๑. การติดอากรแสตมป์ให้เป็นไปตามประมวลกฎหมายรัษฎากร หากต้องการให้ใบสั่งซื้อมีผลตามกฎหมาย")
    _paragraph(doc, f"\t๒. ใบสั่งซื้อนี้อ้างอิงตามเลขที่โครงการ {purchase.project_number or '........................'} ซื้อพัสดุจำนวน {len(purchase.lines)} รายการ เป็นเงิน {purchase.total_amount:,.2f} บาท ({baht_text(purchase.total_amount)}) โดยวิธีเฉพาะเจาะจง")"""
    code = code.replace(old_remarks, new_remarks)

    # 4. Update signature block
    old_sig = """    _paragraph(doc, "ลงชื่อ ................................................ ผู้สั่งซื้อ\\n(................................................)\\nหัวหน้าเจ้าหน้าที่\\nวันที่ ................................................", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "ลงชื่อ ................................................ ผู้รับใบสั่งซื้อ\\n(................................................)\\nวันที่ ................................................", WD_ALIGN_PARAGRAPH.CENTER)"""
    new_sig = """    _paragraph(doc, "\\t\\t\\t\\tลงชื่อ ................................................ ผู้สั่งซื้อ\\n\\t\\t\\t\\t(................................................)\\n\\t\\t\\t\\tหัวหน้าเจ้าหน้าที่\\n\\t\\t\\t\\tวันที่ ................................................")
    _paragraph(doc, "\\t\\t\\t\\t\\tลงชื่อ ................................................ ผู้รับใบสั่งซื้อ\\n\\t\\t\\t\\t\\t(................................................)\\n\\t\\t\\t\\t\\tวันที่ ................................................")"""
    code = code.replace(old_sig, new_sig)

    with open("app/routes.py", "w") as f:
        f.write(code)
    print("Patched!")

patch()
