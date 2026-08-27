import re
_DIGIT_SPAN_RE = re.compile(r"[0-9]+(?:[,.][0-9]+)*")
text = "1,150.00 บาท ๓. นางสาว"
for m in _DIGIT_SPAN_RE.finditer(text):
    print("Digit:", m.group())
