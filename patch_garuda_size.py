def patch():
    with open("app/routes.py", "r") as f:
        code = f.read()

    # Need to import Emu if not imported
    if "from docx.shared import Emu" not in code:
        # docx.shared import usually looks like: from docx.shared import Cm, Pt, RGBColor
        if "from docx.shared import Cm" in code:
            code = code.replace("from docx.shared import Cm, Pt, RGBColor", "from docx.shared import Cm, Pt, RGBColor, Emu")
            code = code.replace("from docx.shared import Cm, Pt", "from docx.shared import Cm, Pt, Emu")

    # Change add_picture width
    old_line = "run.add_picture(str(path), width=Cm(2.8))"
    new_line = "run.add_picture(str(path), width=Emu(635726), height=Emu(696026))"

    if old_line in code:
        code = code.replace(old_line, new_line)
        with open("app/routes.py", "w") as f:
            f.write(code)
        print("Patched garuda size!")
    else:
        print("Could not find the garuda size line to patch.")

patch()
