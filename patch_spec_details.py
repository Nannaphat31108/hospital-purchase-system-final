import re

def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    spec_details = """
def _add_spec_details(doc, purchase):
    # Add Garuda Image
    image_path = Path("/Users/gbru/Downloads/project/image.png")
    if image_path.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run()
        run.add_picture(str(image_path), width=Cm(1.27), height=Cm(1.4))

    # Add header text
    _paragraph(doc, "รายละเอียดขอบเขตและคุณลักษณะเวชภัณฑ์มิใช่ยาที่จะซื้อหรือจ้าง", WD_ALIGN_PARAGRAPH.CENTER, True, 16)
    _paragraph(doc, "เวชภัณฑ์มิใช่ยา", WD_ALIGN_PARAGRAPH.CENTER, True, 16)
    _paragraph(doc, "โรงพยาบาลสิงห์บุรี", WD_ALIGN_PARAGRAPH.CENTER, True, 16)

    # Add table
    table = doc.add_table(rows=1, cols=6)
    _apply_table_grid(table)
    table.autofit = False
    
    # Define widths (approximate from excel)
    widths = [Cm(1.2), Cm(10.0), Cm(2.0), Cm(1.5), Cm(2.0), Cm(2.3)]
    for i in range(6):
        table.columns[i].width = widths[i]
        table.rows[0].cells[i].width = widths[i]

    # Header Row
    headers = ["ลำดับ", "รายการ", "จำนวน", "", "หน่วยละ", "เป็นเงิน"]
    for i, value in enumerate(headers):
        _set_cell_text(table.rows[0].cells[i], value, True, WD_ALIGN_PARAGRAPH.CENTER, size=16)

    # Merge cell 2 and 3 for 'จำนวน' header
    table.rows[0].cells[2].merge(table.rows[0].cells[3])

    # Items
    for line in purchase.lines:
        cells = table.add_row().cells
        values = [
            line.line_no, 
            line.description, 
            f"{line.quantity:g}", 
            line.unit.name, 
            f"{line.unit_price:,.2f}", 
            f"{line.amount:,.2f}"
        ]
        alignments = [
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER
        ]
        for i, value in enumerate(values):
            cells[i].width = widths[i]
            _set_cell_text(cells[i], value, False, alignments[i], size=16)

    # Empty row for spacing
    empty_cells = table.add_row().cells
    for i in range(6):
        empty_cells[i].width = widths[i]
        _set_cell_text(empty_cells[i], "", False, WD_ALIGN_PARAGRAPH.LEFT, size=16)

    # Bullet points
    bullets = [
        " -ใช้เพื่อประกอบการรักษาพยาบาลผู้ป่วย",
        " -เวชภัณฑ์มิใช่ยาที่ใช้ในการรักษาพยาบาลผู้ป่วยที่ขึ้นทะเบียนกับ อย.ประเทศไทย",
        " -ผลิตภัณฑ์ต้องเป็นของใหม่และไม่เคยใช้งานมาก่อน",
        " -ผลิตภัณฑ์ต้องมีอายุการใช้งานอย่างน้อย 1 ปี นับตั้งแต่วันส่งมอบ"
    ]
    for bullet in bullets:
        row = table.add_row()
        for i in range(6):
            row.cells[i].width = widths[i]
        # Merge columns 1 to 5
        row.cells[1].merge(row.cells[5])
        _set_cell_text(row.cells[1], bullet, False, WD_ALIGN_PARAGRAPH.LEFT, size=16)

    # Empty rows for padding (4 rows)
    for _ in range(4):
        empty_cells = table.add_row().cells
        for i in range(6):
            empty_cells[i].width = widths[i]
            _set_cell_text(empty_cells[i], "", False, WD_ALIGN_PARAGRAPH.LEFT, size=16)

    # Total Row
    total_cells = table.add_row().cells
    for i in range(6):
        total_cells[i].width = widths[i]
    total_cells[1].merge(total_cells[2])
    _set_cell_text(total_cells[1], f"({baht_text(purchase.total_amount)})", True, WD_ALIGN_PARAGRAPH.CENTER, size=16)
    
    total_cells[3].merge(total_cells[4])
    _set_cell_text(total_cells[3], "รวมเป็นเงินทั้งสิ้น", True, WD_ALIGN_PARAGRAPH.CENTER, size=16)
    
    _set_cell_text(total_cells[5], f"{purchase.total_amount:,.2f}", True, WD_ALIGN_PARAGRAPH.CENTER, size=16)

    # Footer text
    _paragraph(doc, "")
    _paragraph(doc, "โดยวิธีซื้อครั้งหลังสุดภายในระยะเวลา ๒ ปีงบประมาณ พิจารณาคัดเลือกข้อเสนอโดยใช้หลักเกณฑ์ราคา", WD_ALIGN_PARAGRAPH.LEFT, False, 16)
    _paragraph(doc, "")
    _paragraph(doc, "")
    
    # Signature
    _paragraph(doc, "                         (ลงชื่อ)...........................................................ผู้กำหนดรายละเอียดคุณลักษณะเฉพาะ", WD_ALIGN_PARAGRAPH.LEFT, False, 16)
    
    spec_name = purchase.government_profile.specifier_name if purchase.government_profile else "............................................"
    spec_pos = purchase.government_profile.specifier_position if purchase.government_profile else "............................................"
    
    _paragraph(doc, f"                                      ({spec_name})", WD_ALIGN_PARAGRAPH.LEFT, False, 16)
    _paragraph(doc, f"                                 {spec_pos}", WD_ALIGN_PARAGRAPH.LEFT, False, 16)
"""

    old_all_block = """        doc.add_page_break()
        # 6. ใบตรวจรับ (receipt)
        _add_acceptance_receipt(doc, purchase)
    else:"""
        
    new_all_block = """        doc.add_page_break()
        # 6. ใบตรวจรับ (receipt)
        _add_acceptance_receipt(doc, purchase)
        
        doc.add_page_break()
        # 9. แบบฟอร์มกำหนด spec
        _add_spec_details(doc, purchase)
    else:"""

    if "def _build_word" in code and old_all_block in code:
        idx = code.find("def _build_word")
        code = code[:idx] + spec_details + "\n\n" + code[idx:]
        code = code.replace(old_all_block, new_all_block)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched app/routes.py successfully!")
    else:
        print("Failed to patch. Could not find target blocks.")

patch()
