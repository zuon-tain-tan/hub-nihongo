import json
import re
import unicodedata
from pathlib import Path

from sort_keys import get_group_and_sort_key, get_number_prefix, get_vietnamese_sort_key

VOCABULARY_LEVELS = ("n5", "n4")
VOCABULARY_LEVEL_LABELS = {
    "n5": "N5",
    "n4": "N4",
}

VOCABULARY_DISPLAY_WEIGHT_LIMIT = 48
PARENTHESIS_VARIANT_WEIGHT_LIMIT = 16
KANJI_SUFFIX_STOP_CHARS = set("はがをにへでともかのしすまさただな")

VOCABULARY_SECTION_TITLES = {
    "01_person_pronouns_and_forms_of_address.json": "1. Đại từ chỉ người & xưng hô",
    "02_family_and_social_relationships.json": "2. Gia đình & mối quan hệ xã hội",
    "03_jobs_and_positions.json": "3. Nghề nghiệp và chức vụ",
    "04_countries_cities_and_languages.json": "4. Quốc gia, thành phố & ngôn ngữ",
    "05_object_and_place_pronouns.json": "5. Đại từ chỉ vật & địa điểm",
    "06_personal_items_and_documents.json": "6. Vật dụng cá nhân & giấy tờ",
    "07_clothing_and_accessories.json": "7. Quần áo & phụ kiện",
    "08_stationery_and_tools.json": "8. Giấy bút & dụng cụ",
    "09_machines_equipment_and_parts.json": "9. Máy móc, thiết bị & linh kiện",
    "10_house_architecture_and_interior.json": "10. Kiến trúc nhà ở & nội thất",
    "11_shopping_and_money.json": "11. Mua sắm & tiền bạc",
    "12_post_office_and_shipping_services.json": "12. Bưu điện & dịch vụ gửi hàng",
    "13_places_and_locations.json": "13. Nơi chốn & địa điểm",
    "14_transportation_and_vehicles.json": "14. Giao thông & Phương tiện",
    "15_positions_and_space.json": "15. Vị trí & không gian",
    "16_time_and_clock.json": "16. Thời gian & giờ giấc",
    "17_dates_and_time_markers.json": "17. Ngày tháng & mốc thời gian",
    "18_weather_and_natural_environment.json": "18. Thời tiết & môi trường tự nhiên",
    "19_movement_and_activity_verbs.json": "19. Động từ di chuyển & hoạt động",
    "20_daily_life_and_clothing.json": "20. Sinh hoạt hàng ngày & trang phục",
    "21_study_mind_and_society.json": "21. Học tập, trí não & xã hội",
    "22_communication_giving_receiving_and_requests.json": "22. Giao tiếp, trao nhận & nhờ vả",
    "23_hand_actions_and_machine_operations.json": "23. Thao tác tay & hoạt động máy",
    "24_food_ingredients_and_tableware.json": "24. Nguyên liệu, đồ ăn & dụng cụ ăn uống",
    "25_drinks_and_seasonings.json": "25. Thức uống & gia vị",
    "26_dining_and_table_conversation.json": "26. Ăn uống & giao tiếp bàn tiệc",
    "27_entertainment_arts_and_sports.json": "27. Giải trí, nghệ thuật & thể thao",
    "28_animals.json": "28. Tên các loài vật",
    "29_adjectives_and_life_conditions.json": "29. Tính từ & tình trạng cuộc sống",
    "30_i_adjectives_and_feelings.json": "30. Tính từ đuôi i & Cảm giác",
    "31_basic_colors.json": "31. Màu sắc cơ bản",
    "32_adverbs_and_degree.json": "32. Trạng từ & mức độ",
    "33_connectors_pronouns_and_grammar.json": "33. Từ nối, đại từ & ngữ pháp",
    "34_question_words_and_question_patterns.json": "34. Từ nghi vấn & các cách hỏi",
    "35_body_parts_and_features.json": "35. Bộ phận & đặc điểm cơ thể",
    "36_illness_and_healthcare.json": "36. Bệnh tật & y tế",
    "37_counting_units.json": "37. Các đơn vị đếm",
    "38_confirmation_exclamation_and_apology.json": "38. Xác nhận, cảm thán & xin lỗi",
    "39_greetings_invitations_and_check_ins.json": "39. Chào hỏi, rủ rê & hỏi han",
}


def get_text_weight(text):
    weight = 0
    for char in text or "":
        if "\u3040" <= char <= "\u30ff" or "\u3400" <= char <= "\u9fff":
            weight += 2
        elif ord(char) > 127:
            weight += 1.2
        else:
            weight += 1
    return weight


def contains_kanji(text):
    return any("\u3400" <= char <= "\u9fff" for char in text or "")


def split_top_level(text, separator):
    parts = []
    current = []
    depth = 0
    opening_chars = {"(", "（"}
    closing_chars = {")", "）"}

    for char in text or "":
        if char in opening_chars:
            depth += 1
        elif char in closing_chars and depth > 0:
            depth -= 1

        if char == separator and depth == 0:
            part = "".join(current).strip()
            if part:
                parts.append(part)
            current = []
        else:
            current.append(char)

    part = "".join(current).strip()
    if part:
        parts.append(part)
    return parts


def has_parenthesized_variant(text):
    return bool(re.search(r"[（(][^()（）]+[）)]\s*$", text or ""))


def compact_parenthesis_variant(text):
    if not text or get_text_weight(text) <= PARENTHESIS_VARIANT_WEIGHT_LIMIT or not contains_kanji(text):
        return text

    last_kanji_index = max(index for index, char in enumerate(text) if "\u3400" <= char <= "\u9fff")
    prefix_end = min(len(text), last_kanji_index + 1)
    if prefix_end < len(text) and text[prefix_end] in "々〆ヶ":
        prefix_end += 1
    if (
        prefix_end < len(text)
        and "\u3040" <= text[prefix_end] <= "\u309f"
        and text[prefix_end] not in KANJI_SUFFIX_STOP_CHARS
    ):
        prefix_end += 1

    suffix = text[-1]
    prefix = text[:prefix_end]
    compacted = f"{prefix}...{suffix}"
    return compacted if len(compacted) < len(text) else text


def compact_long_vocabulary_text(text):
    if not text or get_text_weight(text) <= VOCABULARY_DISPLAY_WEIGHT_LIMIT:
        return text

    match = re.match(r"^(.*?)\s*[（(]([^()（）]+)[）)]\s*$", text)
    if not match:
        return text

    reading, variant = match.groups()
    compacted_variant = compact_parenthesis_variant(variant)
    if compacted_variant == variant:
        return text
    return f"{reading.strip()} ({compacted_variant})"


def get_vocabulary_display_text(item):
    tu_vung = compact_long_vocabulary_text(item.get("tu_vung", ""))
    if item.get("cap_do") == "N4" and tu_vung:
        return f"{tu_vung} - {item.get('cap_do', '')}"
    return tu_vung


def normalize_vietnamese_for_dedupe(text):
    normalized = unicodedata.normalize("NFD", (text or "").lower())
    normalized = "".join(char for char in normalized if unicodedata.category(char) != "Mn")
    normalized = normalized.replace("đ", "d")
    normalized = re.sub(r"\bkhong thich\b", "ghet", normalized)
    normalized = re.sub(r"[^0-9a-z]+", " ", normalized)
    return " ".join(normalized.split())


def get_meaning_tokens(text):
    normalized = normalize_vietnamese_for_dedupe(text)
    return set(normalized.split())


def meanings_are_near_duplicate(first_text, second_text):
    first_normalized = normalize_vietnamese_for_dedupe(first_text)
    second_normalized = normalize_vietnamese_for_dedupe(second_text)
    if not first_normalized or not second_normalized:
        return False
    if first_normalized == second_normalized:
        return True

    first_tokens = get_meaning_tokens(first_text)
    second_tokens = get_meaning_tokens(second_text)
    if not first_tokens or not second_tokens:
        return False

    smaller_tokens, larger_tokens = sorted([first_tokens, second_tokens], key=len)
    if smaller_tokens.issubset(larger_tokens):
        return True

    overlap_ratio = len(first_tokens & second_tokens) / len(first_tokens | second_tokens)
    return overlap_ratio >= 0.72


def get_meaning_quality_score(text):
    tokens = get_meaning_tokens(text)
    return (len(tokens), get_text_weight(text))


def add_index_word(all_words, index_word_map, item):
    key = (item.get("cap_do", ""), item.get("tu_vung", "").strip())
    matching_indexes = index_word_map.setdefault(key, [])

    for index in matching_indexes:
        existing_item = all_words[index]
        if not meanings_are_near_duplicate(existing_item.get("y_nghia", ""), item.get("y_nghia", "")):
            continue

        if get_meaning_quality_score(item.get("y_nghia", "")) > get_meaning_quality_score(existing_item.get("y_nghia", "")):
            all_words[index] = item
        return

    matching_indexes.append(len(all_words))
    all_words.append(item)


def capitalize_first(text):
    text = (text or "").strip()
    return text[:1].upper() + text[1:] if text else text


def split_parallel_meanings(text, expected_count):
    for pattern in (r"\s*[,;]\s*", r"\s*/\s*"):
        parts = [part.strip() for part in re.split(pattern, text or "") if part.strip()]
        if len(parts) == expected_count:
            return [capitalize_first(part) for part in parts]
    return None


def expand_vocabulary_item(item):
    vocab_parts = split_top_level(item.get("tu_vung", ""), "/")
    if len(vocab_parts) <= 1 or not all(has_parenthesized_variant(part) for part in vocab_parts):
        return [item]

    meaning_parts = split_parallel_meanings(item.get("y_nghia", ""), len(vocab_parts))
    if not meaning_parts:
        return [item]

    expanded_items = []
    for vocab_part, meaning_part in zip(vocab_parts, meaning_parts):
        item_copy = item.copy()
        item_copy["tu_vung"] = vocab_part
        item_copy["y_nghia"] = meaning_part
        expanded_items.append(item_copy)
    return expanded_items


def clean_vocabulary_meaning(text, level_label, lesson_number):
    if level_label != "N4" or lesson_number < 36:
        return text

    return re.sub(
        r"\s*[,;]\s*(?:kính ngữ|khiêm nhường ngữ) của .*$",
        "",
        text or "",
        flags=re.IGNORECASE,
    ).strip()


def get_display_title(json_path, level_label):
    if json_path.name in VOCABULARY_SECTION_TITLES:
        return VOCABULARY_SECTION_TITLES[json_path.name]

    bai_match = re.match(r"bai(\d+)_(.+)", json_path.stem)
    if bai_match:
        lesson_number, slug = bai_match.groups()
        title = slug.replace("_", " ").title()
        return f"{level_label} - Bài {lesson_number} - {title}"

    return f"{level_label} - {json_path.stem}"


def load_json_file(path, default):
    path = Path(path)
    if not path.exists():
        return default

    with path.open("r", encoding="utf-8") as file:
        return json.load(file)


def load_vocabulary_data(data_dir):
    vocabulary_dir = Path(data_dir) / "vocabulary"
    all_words = []
    json_data_map = []
    index_word_map = {}

    if not vocabulary_dir.exists():
        return all_words, json_data_map

    for level_name in VOCABULARY_LEVELS:
        level_dir = vocabulary_dir / level_name
        level_label = VOCABULARY_LEVEL_LABELS[level_name]

        if not level_dir.exists():
            continue

        json_files = sorted(
            [path for path in level_dir.iterdir() if path.suffix == ".json"],
            key=lambda path: get_number_prefix(path.name),
        )

        for json_path in json_files:
            data = load_json_file(json_path, [])
            decorated_data = []
            lesson_number = get_number_prefix(json_path.name)

            for source_item in data:
                for item in expand_vocabulary_item(source_item):
                    item_copy = item.copy()
                    item_copy["cap_do"] = level_label
                    item_copy["y_nghia"] = clean_vocabulary_meaning(
                        item_copy.get("y_nghia", ""),
                        level_label,
                        lesson_number,
                    )
                    item_copy["tu_vung_hien_thi"] = get_vocabulary_display_text(item_copy)

                    jp_group, jp_sort_key = get_group_and_sort_key(item_copy.get("tu_vung", ""))
                    item_copy["jp_group_char"] = jp_group
                    item_copy["jp_sort_key"] = jp_sort_key

                    vn_group, vn_sort_key = get_vietnamese_sort_key(item_copy.get("y_nghia", ""))
                    item_copy["vn_group_char"] = vn_group
                    item_copy["vn_sort_key"] = vn_sort_key
                    decorated_data.append(item_copy)

                    add_index_word(all_words, index_word_map, item_copy)

            json_data_map.append(
                {
                    "file_name": json_path.name,
                    "display_title": get_display_title(json_path, level_label),
                    "level": level_label,
                    "data": decorated_data,
                    "prefix_num": get_number_prefix(json_path.name),
                }
            )

    return all_words, json_data_map


def load_5s_data(data_dir):
    return load_json_file(Path(data_dir) / "topics" / "5s_5k_5c.json", [])


def load_special_topics(data_dir):
    topic_dir = Path(data_dir) / "topics"
    return {
        "horenso": load_json_file(topic_dir / "horenso.json", {}),
        "aisatsu": load_json_file(topic_dir / "aisatsu.json", []),
        "garbage_sorting": load_json_file(topic_dir / "garbage_sorting.json", []),
    }


def load_kaiwa_before_departure(data_dir):
    return load_json_file(Path(data_dir) / "topics" / "kaiwa_before_departure.json", {})
