def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    old_block = """    _paragraph(doc, "\\t\\t\\t\\t\\t\\tลงชื่อ ....................................................... เจ้าหน้าที่")
    _paragraph(doc, "\\t\\t\\t\\t\\t\\t(นางพิณนภา ศริพันธุ์)")
    _paragraph(doc, "\\t\\t\\t\\t\\t\\tเภสัชกรชำนาญการ")
    _paragraph(doc, "\\t\\t\\t\\t\\t\\tลงชื่อ ...................................................... หัวหน้าเจ้าหน้าที่")
    _paragraph(doc, "\\t\\t\\t\\t\\t\\t(นายชัชวาลย์ บุญญฤทธิ์)")
    _paragraph(doc, "\\t\\t\\t\\t\\t\\tเภสัชกรชำนาญการพิเศษ")
    _paragraph(doc, "\\t\\t\\t\\tอนุมัติ")
    _paragraph(doc, "\\t\\t(นายพิรุณ ปิตะหงษ์นันท์)")
    _paragraph(doc, "ผู้อำนวยการโรงพยาบาลสิงห์บุรี ปฏิบัติงานแทน")
    _paragraph(doc, "\\tผู้ว่าราชการจังหวัดสิงห์บุรี")"""

    new_block = """    _paragraph(doc, "ลงชื่อ ....................................................... เจ้าหน้าที่", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "(นางพิณนภา ศริพันธุ์)", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "เภสัชกรชำนาญการ", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "ลงชื่อ ...................................................... หัวหน้าเจ้าหน้าที่", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "(นายชัชวาลย์ บุญญฤทธิ์)", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "เภสัชกรชำนาญการพิเศษ", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "อนุมัติ", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "(นายพิรุณ ปิตะหงษ์นันท์)", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "ผู้อำนวยการโรงพยาบาลสิงห์บุรี ปฏิบัติงานแทน", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "ผู้ว่าราชการจังหวัดสิงห์บุรี", WD_ALIGN_PARAGRAPH.CENTER)"""

    if old_block in code:
        code = code.replace(old_block, new_block)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched signature alignment!")
    else:
        print("Could not find block!")
        
patch()
