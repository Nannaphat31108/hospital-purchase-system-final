from docx.enum.text import WD_BREAK

def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    # 1. Break before item 7
    old_item7 = '        "๗. หน่วยงานของรัฐสามารถนำผลการปฏิบัติงานตามสัญญาหรือข้อตกลงมาประเมินผลการปฏิบัติงานของผู้ประกอบการ",\n    ]'
    new_item7 = '        "PAGE_BREAK",\n        "๗. หน่วยงานของรัฐสามารถนำผลการปฏิบัติงานตามสัญญาหรือข้อตกลงมาประเมินผลการปฏิบัติงานของผู้ประกอบการ",\n    ]'
    code = code.replace(old_item7, new_item7)

    old_loop = '    for text in conditions:\n        _paragraph(doc, text)'
    new_loop = '    for text in conditions:\n        if text == "PAGE_BREAK":\n            doc.add_page_break()\n        else:\n            _paragraph(doc, text)'
    code = code.replace(old_loop, new_loop)

    # 2. Add spaces before and between signatures
    old_sigs = """    _paragraph(doc, "\\t\\t\\t\\tลงชื่อ ................................................ ผู้สั่งซื้อ\\n\\t\\t\\t\\t(................................................)\\n\\t\\t\\t\\tหัวหน้าเจ้าหน้าที่\\n\\t\\t\\t\\tวันที่ ................................................")
    _paragraph(doc, "\\t\\t\\t\\t\\tลงชื่อ ................................................ ผู้รับใบสั่งซื้อ\\n\\t\\t\\t\\t\\t(................................................)\\n\\t\\t\\t\\t\\tวันที่ ................................................")"""
    new_sigs = """    _paragraph(doc, "")
    _paragraph(doc, "")
    _paragraph(doc, "\\t\\t\\t\\tลงชื่อ ................................................ ผู้สั่งซื้อ\\n\\t\\t\\t\\t(................................................)\\n\\t\\t\\t\\tหัวหน้าเจ้าหน้าที่\\n\\t\\t\\t\\tวันที่ ................................................")
    _paragraph(doc, "")
    _paragraph(doc, "")
    _paragraph(doc, "\\t\\t\\t\\t\\tลงชื่อ ................................................ ผู้รับใบสั่งซื้อ\\n\\t\\t\\t\\t\\t(................................................)\\n\\t\\t\\t\\t\\tวันที่ ................................................")"""
    code = code.replace(old_sigs, new_sigs)

    with open("app/routes.py", "w") as f:
        f.write(code)
    print("Patched spacing and page break!")

patch()
