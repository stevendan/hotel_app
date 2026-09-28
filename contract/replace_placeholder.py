"""Thay placeholder trong toàn bộ file mẫu .docx của thư mục template/.

Cách dùng (chạy từ thư mục gốc dự án hoặc thư mục contract):
    python contract/replace_placeholder.py            # thay thật
    python contract/replace_placeholder.py --dry-run  # chỉ liệt kê, không ghi file
"""

import argparse
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

OLD = "{{client_office_address_vi}}"
NEW = "{{client_address_vi}}"

TEMPLATE_DIR = Path(__file__).resolve().parent / "template"
# Các phần XML của Word có thể chứa chữ: thân văn bản, header, footer, chú thích...
WORD_XML = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")
TAG = re.compile(r"<[^>]+>")


def build_split_pattern(text):
    # Word đôi khi tách "{{abc}}" thành nhiều <w:r>; cho phép thẻ XML xen giữa từng ký tự.
    return re.compile("(?:<[^>]+>)*".join(re.escape(ch) for ch in text))


SPLIT_PATTERN = build_split_pattern(OLD)


def replace_in_xml(xml):
    """Trả về (xml mới, số lần thay)."""
    count = xml.count(OLD)
    xml = xml.replace(OLD, NEW)

    def repl(match):
        # Đặt toàn bộ giá trị mới vào đoạn chữ đầu tiên, giữ nguyên các thẻ XML để không vỡ cấu trúc.
        parts = re.split(r"(<[^>]+>)", match.group(0))
        out, placed = [], False
        for part in parts:
            if TAG.fullmatch(part):
                out.append(part)
            elif part and not placed:
                out.append(NEW)
                placed = True
        return "".join(out)

    xml, split_count = SPLIT_PATTERN.subn(repl, xml)
    return xml, count + split_count


def process_docx(path, dry_run):
    total = 0
    new_parts = {}
    with zipfile.ZipFile(path) as zin:
        for info in zin.infolist():
            if WORD_XML.match(info.filename):
                xml = zin.read(info).decode("utf-8")
                new_xml, count = replace_in_xml(xml)
                if count:
                    new_parts[info.filename] = new_xml.encode("utf-8")
                    total += count

        if not total or dry_run:
            return total

        # zipfile không sửa tại chỗ được: ghi ra file tạm rồi thay thế file gốc.
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx", dir=path.parent) as tmp:
            tmp_path = Path(tmp.name)
        with zipfile.ZipFile(tmp_path, "w") as zout:
            for info in zin.infolist():
                data = new_parts.get(info.filename)
                if data is None:
                    data = zin.read(info)
                zout.writestr(info, data, compress_type=info.compress_type)

    shutil.move(tmp_path, path)
    return total


def main():
    parser = argparse.ArgumentParser(description=f"Thay {OLD} bằng {NEW} trong các file mẫu.")
    parser.add_argument("--dry-run", action="store_true", help="Chỉ liệt kê, không ghi file")
    args = parser.parse_args()

    files = sorted(TEMPLATE_DIR.glob("*.docx"))
    if not files:
        print(f"Không tìm thấy file .docx nào trong {TEMPLATE_DIR}")
        return 1

    changed = 0
    for path in files:
        if path.name.startswith("~$"):  # file khóa tạm của Word khi đang mở
            continue
        try:
            count = process_docx(path, args.dry_run)
        except (zipfile.BadZipFile, PermissionError) as error:
            print(f"[LỖI]  {path.name}: {error}")
            continue
        if count:
            changed += 1
            print(f"[{'SẼ SỬA' if args.dry_run else 'ĐÃ SỬA'}] {path.name}: {count} chỗ")
        else:
            print(f"[BỎ QUA] {path.name}: không có {OLD}")

    action = "sẽ được sửa" if args.dry_run else "đã được sửa"
    print(f"\nTổng: {changed}/{len(files)} file {action}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
