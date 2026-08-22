def run():
    with open("app/routes.py", "r") as f:
        code = f.read()
    
    start = code.find('    _paragraph(doc, "\\tลงชื่อ')
    end = code.find('    _paragraph(doc, "                       อนุมัติ")')
    print(code[start:end])
run()
