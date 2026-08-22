def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    old_block = """    _paragraph(doc, "                                  ลงชื่อ ....................................................... เจ้าหน้าที่")
    _paragraph(doc, "                                 (นางพิณนภา ศริพันธุ์)")
    _paragraph(doc, "                                 เภสัชกรชำนาญการ")
    _paragraph(doc, "")
    _paragraph(doc, "                                             ลงชื่อ ...................................................... หัวหน้าเจ้าหน้าที่")
    _paragraph(doc, "                                 (นายชัชวาลย์ บุญญฤทธิ์)")
    _paragraph(doc, "                                 เภสัชกรชำนาญการพิเศษ")"""

    new_block = """    _paragraph(doc, "\\tลงชื่อ ....................................................... เจ้าหน้าที่", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "\\t(นางพิณนภา ศริพันธุ์)", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "\\tเภสัชกรชำนาญการ", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "")
    _paragraph(doc, "\\tลงชื่อ ...................................................... หัวหน้าเจ้าหน้าที่", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "\\t(นายชัชวาลย์ บุญญฤทธิ์)", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "\\tเภสัชกรชำนาญการพิเศษ", WD_ALIGN_PARAGRAPH.CENTER)"""

    if old_block in code:
        code = code.replace(old_block, new_block)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched 2 people signatures!")
    else:
        print("Could not find block!")
        
patch()
