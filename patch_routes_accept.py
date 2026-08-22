import re

def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    if "THAI_MONTHS_FULL" not in code:
        thai_months = """
THAI_MONTHS_FULL = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

def format_thai_date_full(dt):
    if not dt: return ""
    return f"{dt.day} {THAI_MONTHS_FULL[dt.month - 1]} {dt.year + 543}"
"""
        code = code.replace("from docx.shared import Cm, Pt, Emu\n", "from docx.shared import Cm, Pt, Emu\n" + thai_months)

    # Now let's fix _add_acceptance_receipt
    # Date line
    old_date1 = """    _paragraph(doc, f"วันที่ {purchase.document_date.strftime('%d/%m/%Y')}", WD_ALIGN_PARAGRAPH.RIGHT)"""
    new_date1 = """    _paragraph(doc, f"\\t\\tวันที่ {format_thai_date_full(purchase.document_date)}", WD_ALIGN_PARAGRAPH.CENTER)"""
    code = code.replace(old_date1, new_date1)

    # Also the text line
    old_text1 = """        f"ตามใบสั่งซื้อ เลขที่ {purchase.po_number} ลงวันที่ {purchase.document_date.strftime('%d/%m/%Y')} \""""
    new_text1 = """        f"ตามใบสั่งซื้อ เลขที่ {purchase.po_number} ลงวันที่ {format_thai_date_full(purchase.document_date)} \""""
    code = code.replace(old_text1, new_text1)

    # Now the checkboxes
    old_cb1 = """    _paragraph(doc, "     ☐ ถูกต้อง      ☐ ครบถ้วนตามสัญญา      ☐ ไม่ครบถ้วนตามสัญญา")"""
    new_cb1 = """    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    for part in ["     ", "☐", " ถูกต้อง      ", "☐", " ครบถ้วนตามสัญญา      ", "☐", " ไม่ครบถ้วนตามสัญญา"]:
        r = p.add_run(part)
        if part == "☐":
            r.font.size = Pt(14)"""
    code = code.replace(old_cb1, new_cb1)

    old_cb2 = """    _paragraph(doc, "     ☐ มีค่าปรับ      ☐ ไม่มีค่าปรับ")"""
    new_cb2 = """    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    for part in ["     ", "☐", " มีค่าปรับ      ", "☐", " ไม่มีค่าปรับ"]:
        r = p.add_run(part)
        if part == "☐":
            r.font.size = Pt(14)"""
    code = code.replace(old_cb2, new_cb2)

    with open("app/routes.py", "w") as f:
        f.write(code)
    print("Patched acceptance receipt!")
    
patch()
