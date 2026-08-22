def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    # 1. Change title font size to 16
    code = code.replace(
        '_paragraph(doc, "ใบตรวจรับการจัดซื้อ/จัดจ้าง", WD_ALIGN_PARAGRAPH.CENTER, True, 20)',
        '_paragraph(doc, "ใบตรวจรับการจัดซื้อ/จัดจ้าง", WD_ALIGN_PARAGRAPH.CENTER, True, 16)'
    )

    # 2. Blank line before and after "คณะกรรมการตรวจรับพัสดุ ได้ตรวจรับงานแล้ว ผลปรากฏดังนี้"
    code = code.replace(
        '_paragraph(doc, "คณะกรรมการตรวจรับพัสดุ ได้ตรวจรับงานแล้ว ผลปรากฏดังนี้", first_line=True)',
        '_paragraph(doc, "")\n    _paragraph(doc, "คณะกรรมการตรวจรับพัสดุ ได้ตรวจรับงานแล้ว ผลปรากฏดังนี้", first_line=True)\n    _paragraph(doc, "")'
    )

    # 3 & 4. Add \t to 1, 2, 3 and lines after them. Add blank lines after each section.
    
    # Section 1
    code = code.replace(
        '_paragraph(doc, "1. ผลการตรวจรับ")',
        '_paragraph(doc, "\\t1. ผลการตรวจรับ")'
    )
    old_cb1 = """    for part in ["     ", "☐", " ถูกต้อง      ", "☐", " ครบถ้วนตามสัญญา      ", "☐", " ไม่ครบถ้วนตามสัญญา"]:
        r = p.add_run(part)
        if part == "☐":
            r.font.size = Pt(14)
    _paragraph(doc, "2. ค่าปรับ")"""
    new_cb1 = """    for part in ["\\t", "☐", " ถูกต้อง      ", "☐", " ครบถ้วนตามสัญญา      ", "☐", " ไม่ครบถ้วนตามสัญญา"]:
        r = p.add_run(part)
        if part == "☐":
            r.font.size = Pt(14)
    _paragraph(doc, "")
    _paragraph(doc, "\\t2. ค่าปรับ")"""
    code = code.replace(old_cb1, new_cb1)

    # Section 2
    old_cb2 = """    for part in ["     ", "☐", " มีค่าปรับ      ", "☐", " ไม่มีค่าปรับ"]:
        r = p.add_run(part)
        if part == "☐":
            r.font.size = Pt(14)
    _paragraph(doc, "3. การเบิกจ่ายเงิน")"""
    new_cb2 = """    for part in ["\\t", "☐", " มีค่าปรับ      ", "☐", " ไม่มีค่าปรับ"]:
        r = p.add_run(part)
        if part == "☐":
            r.font.size = Pt(14)
    _paragraph(doc, "")
    _paragraph(doc, "\\t3. การเบิกจ่ายเงิน")"""
    code = code.replace(old_cb2, new_cb2)

    # Section 3
    old_sec3 = """    if len(purchase.lines) == 1:
        line = purchase.lines[0]
        _paragraph(doc, f"     เบิกจ่ายเงิน เป็นจำนวนเงินทั้งสิ้น {line.amount:,.2f} บาท")
    else:
        for line in purchase.lines:
            _paragraph(doc, f"     - รายการที่ {line.line_no} {line.description}")
            _paragraph(doc, f"       เบิกจ่ายเงิน งวดที่ 1 เป็นจำนวนเงินทั้งสิ้น {line.amount:,.2f} บาท")

    sig = doc.add_table(rows=3, cols=2)"""
    new_sec3 = """    if len(purchase.lines) == 1:
        line = purchase.lines[0]
        _paragraph(doc, f"\\tเบิกจ่ายเงิน เป็นจำนวนเงินทั้งสิ้น {line.amount:,.2f} บาท")
    else:
        for line in purchase.lines:
            _paragraph(doc, f"\\t- รายการที่ {line.line_no} {line.description}")
            _paragraph(doc, f"\\tเบิกจ่ายเงิน งวดที่ 1 เป็นจำนวนเงินทั้งสิ้น {line.amount:,.2f} บาท")

    _paragraph(doc, "")
    sig = doc.add_table(rows=3, cols=2)"""
    code = code.replace(old_sec3, new_sec3)

    with open("app/routes.py", "w") as f:
        f.write(code)
    print("Patched acceptance receipt sections 1,2,3!")
    
patch()
