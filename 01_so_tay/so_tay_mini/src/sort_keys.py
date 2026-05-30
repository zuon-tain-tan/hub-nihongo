import re


def get_group_and_sort_key(text):
    """Return group and natural sort key for Japanese lookup entries."""
    if not text:
        return "*", "4_"

    cleaned = text
    for char in ["[", "]", "~", "～", "(", ")", "「", "」", " ", "…"]:
        cleaned = cleaned.replace(char, "")
    cleaned = cleaned.strip()
    if not cleaned:
        return "*", "4_"

    first_char = cleaned[0]
    if "\u3040" <= first_char <= "\u309f":
        category = "1_"
    elif "\u30a0" <= first_char <= "\u30ff":
        category = "2_"
    else:
        category = "3_"

    sort_text = re.sub(r"\d+", lambda match: match.group(0).zfill(5), cleaned.lower())
    return first_char.upper(), category + sort_text


def get_vietnamese_sort_key(text):
    """Return group and natural sort key for Vietnamese lookup entries."""
    if not text:
        return "#", "zzz"

    cleaned = text
    for char in ["[", "]", "~", "～", "(", ")", "「", "」", " ", "…", "-", '"', "'"]:
        cleaned = cleaned.replace(char, "")
    cleaned = cleaned.strip()
    if not cleaned:
        return "#", "zzz"

    source = (
        "ÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚÝàáâãèéêìíòóôõùúý"
        "ĂăĐđĨĩŨũƠơƯư"
        "ẠạẢảẤấẦầẨẩẪẫẬậẮắẰằẲẳẴẵẶặ"
        "ẸẹẺẻẼẽẾếỀềỂểỄễỆệỈỉỊị"
        "ỌọỎỏỐốỒồỔổỖỗỘộỚớỜờỞởỠỡỢợ"
        "ỤụỦủỨứỪừỬửỮữỰựỲỳỴỵỶỷỸỹ"
    )
    target = (
        "AAAAEEEIIOOOOUUYaaaaeeeiioooouuy"
        "AaDdIiUuOoUu"
        "AaAaAaAaAaAaAaAaAaAaAaAa"
        "EeEeEeEeEeEeEeEeIiIi"
        "OoOoOoOoOoOoOoOoOoOoOoOo"
        "UuUuUuUuUuUuUuUuYyYyYyYy"
    )

    normalized = "".join(target[source.index(char)] if char in source else char for char in cleaned)
    first_char = normalized[0].upper()
    if not ("A" <= first_char <= "Z"):
        first_char = "#"

    sort_text = re.sub(r"\d+", lambda match: match.group(0).zfill(5), normalized.lower())
    return first_char, sort_text


def get_number_prefix(filename):
    match = re.search(r"^(\d+)", filename)
    return int(match.group(1)) if match else 9999

