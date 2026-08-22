from decimal import Decimal, InvalidOperation, ROUND_HALF_UP


def to_decimal(value, default="0"):
    try:
        return Decimal(str(value or default)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError, TypeError):
        return Decimal(default).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def baht_text(number):
    number = to_decimal(number)
    integer, satang = f"{number:.2f}".split(".")
    digits = ["ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"]
    places = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน"]

    def convert_group(group):
        result = ""
        length = len(group)
        for i, char in enumerate(group):
            n = int(char)
            pos = length - i - 1
            if n == 0:
                continue
            if pos == 0 and n == 1 and length > 1:
                result += "เอ็ด"
            elif pos == 1 and n == 2:
                result += "ยี่สิบ"
                continue
            elif pos == 1 and n == 1:
                result += "สิบ"
                continue
            else:
                result += digits[n] + places[pos]
        return result

    def convert_integer(text):
        if int(text) == 0:
            return "ศูนย์"
        if len(text) <= 6:
            return convert_group(text)
        left = text[:-6]
        right = text[-6:]
        return convert_integer(left) + "ล้าน" + (convert_group(right) if int(right) else "")

    result = convert_integer(integer) + "บาท"
    if satang == "00":
        return result + "ถ้วน"
    return result + convert_group(satang) + "สตางค์"


def safe_filename(value):
    text = str(value or "document").strip()
    invalid = '<>:"/\\|?*'
    for char in invalid:
        text = text.replace(char, "-")
    return text or "document"
