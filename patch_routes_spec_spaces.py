def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    # The block starts from:
    # _paragraph(doc, "\t1.\t\t\tนางสาวนลินี เครือทิวา\t\t\t\tตำแหน่ง\tเภสัชกรชำนาญการ")
    # down to:
    # _paragraph(doc, "ผู้ว่าราชการจังหวัดสิงห์บุรี", WD_ALIGN_PARAGRAPH.CENTER)

    old_block = """    _paragraph(doc, "\\t1.\\t\\t\\tนางสาวนลินี เครือทิวา\\t\\t\\t\\tตำแหน่ง\\tเภสัชกรชำนาญการ")
    _paragraph(doc, "\\t\\t\\tจึงเรียนมาเพื่อโปรดพิจารณาอนุมัติ")
    _paragraph(doc, "ลงชื่อ ....................................................... เจ้าหน้าที่", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "(นางพิณนภา ศริพันธุ์)", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "เภสัชกรชำนาญการ", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "ลงชื่อ ...................................................... หัวหน้าเจ้าหน้าที่", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "(นายชัชวาลย์ บุญญฤทธิ์)", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "เภสัชกรชำนาญการพิเศษ", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "อนุมัติ", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "(นายพิรุณ ปิตะหงษ์นันท์)", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "ผู้อำนวยการโรงพยาบาลสิงห์บุรี ปฏิบัติงานแทน", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "ผู้ว่าราชการจังหวัดสิงห์บุรี", WD_ALIGN_PARAGRAPH.CENTER)"""

    new_block = """    _paragraph(doc, "\\t1.\\t\\t\\tนางสาวนลินี เครือทิวา\\t\\t\\t\\tตำแหน่ง\\tเภสัชกรชำนาญการ")
    _paragraph(doc, "")
    _paragraph(doc, "\\t\\tจึงเรียนมาเพื่อโปรดพิจารณาอนุมัติ")
    _paragraph(doc, "")
    _paragraph(doc, "")
    _paragraph(doc, "                                  ลงชื่อ ....................................................... เจ้าหน้าที่")
    _paragraph(doc, "                                 (นางพิณนภา ศริพันธุ์)")
    _paragraph(doc, "                                 เภสัชกรชำนาญการ")
    _paragraph(doc, "")
    _paragraph(doc, "                                             ลงชื่อ ...................................................... หัวหน้าเจ้าหน้าที่")
    _paragraph(doc, "                                 (นายชัชวาลย์ บุญญฤทธิ์)")
    _paragraph(doc, "                                 เภสัชกรชำนาญการพิเศษ")
    _paragraph(doc, "                       อนุมัติ")
    _paragraph(doc, "")
    _paragraph(doc, "")
    _paragraph(doc, "              (นายพิรุณ ปิตะหงษ์นันท์)")
    _paragraph(doc, "ผู้อำนวยการโรงพยาบาลสิงห์บุรี ปฏิบัติงานแทน")
    _paragraph(doc, "            ผู้ว่าราชการจังหวัดสิงห์บุรี")"""

    if old_block in code:
        code = code.replace(old_block, new_block)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched signature spaces!")
    else:
        print("Could not find block!")
        
patch()
