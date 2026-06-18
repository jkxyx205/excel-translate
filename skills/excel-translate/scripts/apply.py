#!/usr/bin/env python3
"""Apply a translation map to an xlsx, preserving images, formulas, styles,
and rich-text run formatting.

Usage:
    python apply.py <input.xlsx> <translations.json> <output.xlsx>

translations.json format:
    {"original text": "translated text", ...}

Only <t> text nodes whose content exactly matches a key are replaced.
All other XML, images (including wmf), formulas, styles, namespaces, and
relationship entries are copied byte-for-byte into the output.
"""
import html
import json
import os
import re
import sys
import zipfile


_T_PATTERN = re.compile(r'(<t(?:\s[^>]*)?>)([^<>]*)(</t>)')


def build_escaped_map(translate_map: dict) -> dict:
    return {html.escape(k, quote=False): html.escape(v, quote=False)
            for k, v in translate_map.items()}


def replace_t_text(xml_text: str, escaped_map: dict) -> str:
    def repl(m):
        open_tag, content, close_tag = m.group(1), m.group(2), m.group(3)
        # Normalize CRLF/CR -> LF before lookup. openpyxl normalizes line
        # endings on read (XML spec 2.11), so keys in escaped_map use LF,
        # but the raw XML bytes we match against here may still be CRLF.
        normalized = content.replace('\r\n', '\n').replace('\r', '\n')
        if normalized in escaped_map:
            return open_tag + escaped_map[normalized] + close_tag
        return m.group(0)
    return _T_PATTERN.sub(repl, xml_text)


def main():
    if len(sys.argv) != 4:
        print("Usage: python apply.py <input.xlsx> <translations.json> <output.xlsx>",
              file=sys.stderr)
        sys.exit(1)

    input_path, translations_path, output_path = sys.argv[1], sys.argv[2], sys.argv[3]

    with open(translations_path, "r", encoding="utf-8") as f:
        translate_map = json.load(f)

    escaped_map = build_escaped_map(translate_map)
    print(f"Loaded {len(translate_map)} translations")

    with zipfile.ZipFile(input_path, "r") as zin, \
         zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            is_strings = item.filename == "xl/sharedStrings.xml"
            is_sheet = (item.filename.startswith("xl/worksheets/")
                        and item.filename.endswith(".xml"))
            if is_strings or is_sheet:
                text = data.decode("utf-8")
                text = replace_t_text(text, escaped_map)
                data = text.encode("utf-8")
            zout.writestr(item, data)

    print(f"Wrote translated workbook to {output_path}")


if __name__ == "__main__":
    main()
