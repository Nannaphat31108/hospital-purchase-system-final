def run():
    with open("app/routes.py", "r") as f:
        code = f.read()
    start = code.find('def ')
    for fn in ["generate_po", "export", "download"]:
        idx = code.find(f'def {fn}')
        if idx != -1:
            print(f"Found {fn} at index {idx}")
            print(code[idx:idx+500])
            break
run()
