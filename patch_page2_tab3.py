def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    old_block = """    _paragraph(doc, "\\t\\t\\t\\t\\tลงชื่อ ................................................ ผู้กำหนดรายละเอียด\\n\\t\\t\\t\\t\\t(................................................)", WD_ALIGN_PARAGRAPH.CENTER)"""

    new_block = """    _paragraph(doc, "\\t\\t\\t\\t\\tลงชื่อ ................................................ ผู้กำหนดรายละเอียด\\n\\t\\t\\t(................................................)", WD_ALIGN_PARAGRAPH.CENTER)"""

    if old_block in code:
        code = code.replace(old_block, new_block)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched page 2 signature tabs (bottom line)!")
    else:
        print("Could not find page 2 signature block!")
        
patch()
