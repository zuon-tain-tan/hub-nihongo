from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def ask_voice_type():
    while True:
        voice_type = input("Muon xoa file loai nao (F/M)? ").strip().upper()
        if voice_type in {"F", "M"}:
            return voice_type
        print("Chi nhap F hoac M.")


def parse_lesson_numbers(text):
    lesson_numbers = set()
    parts = [part.strip() for part in text.split(",")]

    for part in parts:
        if not part:
            raise ValueError("Co muc trong.")

        if "-" in part:
            range_parts = [range_part.strip() for range_part in part.split("-", 1)]
            if len(range_parts) != 2 or not range_parts[0].isdigit() or not range_parts[1].isdigit():
                raise ValueError(f"Khoang khong hop le: {part}")

            start = int(range_parts[0])
            end = int(range_parts[1])
            if start <= 0 or end <= 0 or start > end:
                raise ValueError(f"Khoang khong hop le: {part}")

            lesson_numbers.update(range(start, end + 1))
        else:
            if not part.isdigit() or int(part) <= 0:
                raise ValueError(f"So bai khong hop le: {part}")
            lesson_numbers.add(int(part))

    return sorted(lesson_numbers)


def ask_lesson_numbers():
    while True:
        lesson_text = input("Cua bai nao? Vi du 26 hoac 26,28,31 hoac 27-30: ").strip()
        try:
            return parse_lesson_numbers(lesson_text)
        except ValueError as error:
            print(error)


def main():
    voice_type = ask_voice_type()
    lesson_numbers = ask_lesson_numbers()

    files = []
    missing_dirs = []
    empty_lessons = []

    for lesson_number in lesson_numbers:
        lesson_dir = PROJECT_ROOT / "mp3" / f"L{lesson_number}"
        pattern = f"L{lesson_number}-*-{voice_type}.mp3"

        if not lesson_dir.exists():
            missing_dirs.append(str(lesson_dir))
            continue

        lesson_files = sorted(lesson_dir.glob(pattern))
        if lesson_files:
            files.extend(lesson_files)
        else:
            empty_lessons.append(lesson_number)

    if missing_dirs:
        print("\nKhong tim thay cac thu muc:")
        for lesson_dir in missing_dirs:
            print(f"- {lesson_dir}")

    if empty_lessons:
        print("\nKhong tim thay file dung loai trong cac bai:")
        for lesson_number in empty_lessons:
            print(f"- Bai {lesson_number}")

    if not files:
        print("\nKhong co file nao de xoa.")
        return

    print(f"\nSe xoa {len(files)} file:")
    for file_path in files:
        print(f"- {file_path}")

    confirm = input("\nNhap YES de xoa: ").strip()
    if confirm != "YES":
        print("Da huy, khong xoa file nao.")
        return

    deleted_count = 0
    for file_path in files:
        try:
            file_path.unlink()
            deleted_count += 1
        except OSError as error:
            print(f"Khong xoa duoc {file_path}: {error}")

    lessons_text = ", ".join(str(lesson_number) for lesson_number in lesson_numbers)
    print(f"Da xoa {deleted_count}/{len(files)} file loai {voice_type} cua bai: {lessons_text}.")


if __name__ == "__main__":
    main()
