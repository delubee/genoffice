"""Example transformation for a GenOffice DOCX JSON export; never writes DOCX."""
import argparse
import json
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    action = parser.add_mutually_exclusive_group(required=True)
    action.add_argument("--uppercase", action="store_true")
    action.add_argument("--replace", nargs=2, metavar=("OLD", "NEW"))
    args = parser.parse_args()
    if args.input.resolve() == args.output.resolve():
        parser.error("write a separate result file")
    if args.input.stat().st_size > 2 * 1024 * 1024:
        parser.error("input exceeds 2 MiB")
    data = json.loads(args.input.read_text(encoding="utf-8"))
    if data.get("schemaVersion") != 1 or not isinstance(data.get("blocks"), list):
        parser.error("expected a GenOffice DOCX export")
    changed = 0
    for block in data["blocks"]:
        if block.get("protected") or block.get("type") not in ("docParagraph", "docHeading", "docListItem"):
            continue
        text = block.get("text", "")
        result = text.upper() if args.uppercase else text.replace(*args.replace)
        if result != text:
            block["text"] = result
            changed += 1
    with args.output.open("x", encoding="utf-8") as output:
        json.dump(data, output, ensure_ascii=False, indent=2)
    print(f"Prepared {changed} changed paragraphs; review with prepare_import before applying.")


if __name__ == "__main__":
    main()
