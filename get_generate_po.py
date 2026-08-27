def run():
    with open("app/routes.py", "r") as f:
        code = f.read()
    start = code.find('def generate_po(id):')
    end = code.find('    return send_file(', start)
    print(code[start:end])
run()
