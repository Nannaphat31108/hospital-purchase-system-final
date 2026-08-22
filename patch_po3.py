def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    # Find the current signature blocks
    old_sigs = """    _paragraph(doc, "\\t\\t\\t\\tลงชื่อ ................................................ ผู้สั่งซื้อ\\n\\t\\t\\t\\t(................................................)\\n\\t\\t\\t\\tหัวหน้าเจ้าหน้าที่\\n\\t\\t\\t\\tวันที่ ................................................")
    _paragraph(doc, "")
    _paragraph(doc, "")
    _paragraph(doc, "\\t\\t\\t\\t\\tลงชื่อ ................................................ ผู้รับใบสั่งซื้อ\\n\\t\\t\\t\\t\\t(................................................)\\n\\t\\t\\t\\t\\tวันที่ ................................................")"""
    
    new_sigs = """    _paragraph(doc, "\\t\\t\\t\\tลงชื่อ ................................................ ผู้สั่งซื้อ\\n\\t\\t\\t\\t(................................................)\\n\\t\\t\\t\\tหัวหน้าเจ้าหน้าที่\\n\\t\\t\\t\\tวันที่ ................................................", WD_ALIGN_PARAGRAPH.CENTER)
    _paragraph(doc, "")
    _paragraph(doc, "")
    _paragraph(doc, "\\t\\t\\t\\t\\tลงชื่อ ................................................ ผู้รับใบสั่งซื้อ\\n\\t\\t\\t\\t(................................................)\\n\\t\\t\\t\\tวันที่ ................................................", WD_ALIGN_PARAGRAPH.CENTER)"""

    if old_sigs in code:
        code = code.replace(old_sigs, new_sigs)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched signatures!")
    else:
        print("Could not find the exact signature block in routes.py")
        
patch()
