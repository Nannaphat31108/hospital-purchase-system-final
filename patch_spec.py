from docx.shared import Cm

def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    # We need to replace the entire _add_spec function.
    import re
    match = re.search(r"def _add_spec\(doc, purchase\):(.*?)def _add_acceptance_receipt\(doc, purchase\):", code, re.DOTALL)
    if not match:
        print("Could not find _add_spec!")
        return

    new_spec = """def _add_spec(doc, purchase):
    _add_garuda(doc)
    _paragraph(doc, "บันทึกข้อความ", WD_ALIGN_PARAGRAPH.CENTER, True, 20)
    _paragraph(doc, "")
    _paragraph(doc, "ส่วนราชการ\\tโรงพยาบาลสิงห์บุรี กลุ่มงานเภสัชกรรม โทร. ๐๓๖๕๒ ๒๕๐๘ ต่อ ๑๑๒๙")
    _paragraph(doc, f"ที่ สห ๐๐๓๓.๒๐๕.๑๒/........................\\t\\t\\tวันที่ {purchase.document_date.strftime('%d/%m/%Y')}")
    _paragraph(doc, f"เรื่อง\\tขออนุมัติแต่งตั้งผู้กำหนดรายละเอียดคุณลักษณะเฉพาะของพัสดุ จำนวน {len(purchase.lines)} รายการ")
    _paragraph(doc, "เรียน\\tผู้ว่าราชการจังหวัดสิงห์บุรี")
    _paragraph(doc, f"\\tด้วยกลุ่มงานเภสัชกรรม โรงพยาบาลสิงห์บุรี จะดำเนินการจัดซื้อเวชภัณฑ์มิใช่ยา จำนวน {len(purchase.lines)} รายการ  ดังนี้")
    _paragraph(doc, "")

    table = doc.add_table(rows=1, cols=5)
    _apply_table_grid(table)
    headers = ["ลำดับ", "รายการ", "จำนวน", "หน่วยละ", "เป็นเงิน"]
    table.autofit = False
    widths = [Cm(1.2), Cm(9.3), Cm(2.3), Cm(2.5), Cm(2.4)]
    for i in range(5):
        table.columns[i].width = widths[i]
    for i, value in enumerate(headers):
        _set_cell_text(table.rows[0].cells[i], value, True, WD_ALIGN_PARAGRAPH.CENTER)
        table.rows[0].cells[i].width = widths[i]
    for line in purchase.lines:
        cells = table.add_row().cells
        values = [line.line_no, line.description, f"{line.quantity:g} {line.unit.name}", f"{line.unit_price:,.2f}", f"{line.amount:,.2f}"]
        for i, value in enumerate(values):
            cells[i].width = widths[i]
            _set_cell_text(cells[i], value, align=WD_ALIGN_PARAGRAPH.LEFT if i == 1 else WD_ALIGN_PARAGRAPH.CENTER)
    cells = table.add_row().cells
    _set_cell_text(cells[0].merge(cells[3]), "รวมเป็นเงินทั้งสิ้น", True, WD_ALIGN_PARAGRAPH.RIGHT)
    _set_cell_text(cells[4], f"{purchase.total_amount:,.2f}", True, WD_ALIGN_PARAGRAPH.RIGHT)
    _paragraph(doc, f"({baht_text(purchase.total_amount)})", WD_ALIGN_PARAGRAPH.CENTER, True)
    
    _paragraph(doc, "เพื่อให้ได้ร่างขอบเขตของงานหรือรายละเอียดคุณลักษณะเฉพาะของพัสดุดังกล่าว รวมทั้งกำหนดหลักเกณฑ์การพิจารณาคัดเลือกข้อเสนอ เพื่อใช้ในการจัดซื้อ/จัดจ้าง ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. 2560 ข้อ 21 วรรคหนึ่ง และมติคณะรัฐมนตรีและหลักเกณฑ์ที่เกี่ยวข้อง จึงขออนุมัติแต่งตั้งคณะกรรมการหรือผู้กำหนด หรือบุคคลใดบุคคลหนึ่ง ในการจัดทำร่างขอบเขตของงานหรือรายละเอียดคุณลักษณะเฉพาะของพัสดุที่จะซื้อ รวมทั้งกำหนดหลักเกณฑ์การพิจารณาคัดเลือกข้อเสนอ ตามรายชื่อดังนี้", first_line=True)
    _paragraph(doc, "")
    _paragraph(doc, "\\t1.\\t\\t\\tนางสาวนลินี เครือทิวา\\t\\t\\t\\tตำแหน่ง\\tเภสัชกรชำนาญการ")
    _paragraph(doc, "")
    _paragraph(doc, "\\t\\t\\tจึงเรียนมาเพื่อโปรดพิจารณาอนุมัติ")
    _paragraph(doc, "")
    _paragraph(doc, "ลงชื่อ ....................................................... เจ้าหน้าที่", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "(นางพิณนภา ศริพันธุ์)", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "เภสัชกรชำนาญการ", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "")
    _paragraph(doc, "ลงชื่อ ...................................................... หัวหน้าเจ้าหน้าที่", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "(นายชัชวาลย์ บุญญฤทธิ์)", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "เภสัชกรชำนาญการพิเศษ", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "")
    _paragraph(doc, "อนุมัติ", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "(นายพิรุณ ปิตะหงษ์นันท์)", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "ผู้อำนวยการโรงพยาบาลสิงห์บุรี ปฏิบัติงานแทน", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "ผู้ว่าราชการจังหวัดสิงห์บุรี", WD_ALIGN_PARAGRAPH.CENTER)
    
    doc.add_page_break()
    _paragraph(doc, "รายละเอียดคุณลักษณะเฉพาะ", WD_ALIGN_PARAGRAPH.CENTER, True, 20)
    _paragraph(doc, f"พัสดุจำนวน {len(purchase.lines)} รายการ")
    for line in purchase.lines:
        _paragraph(doc, f"{line.line_no}. {line.description}")
        _paragraph(doc, "คุณลักษณะเฉพาะตัวอย่าง สามารถแก้ไขได้")
        _paragraph(doc, f"จำนวน {line.quantity:g} {line.unit.name}")
    _paragraph(doc, "")
    _paragraph(doc, "\\t\\tลงชื่อ ................................................ ผู้กำหนดรายละเอียด\\n\\t\\t(................................................)")
    
def _add_acceptance_receipt(doc, purchase):"""

    code = code[:match.start()] + new_spec + code[match.end()-38:]
    with open("app/routes.py", "w") as f:
        f.write(code)
    print("Patched _add_spec!")

patch()
