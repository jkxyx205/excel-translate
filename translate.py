
import html
import os
import re
import zipfile

from openpyxl import load_workbook


def translate_excel(path: str):
  """
    Extract all unique text values from an Excel file.
  """
  wb = load_workbook(path)
  texts = set()

  for sheet in wb.worksheets:
      # 添加 sheet 名称到 texts 集合中
      if sheet.title:
        texts.add(sheet.title)

      for row in sheet.iter_rows():
          for cell in row:
              if isinstance(cell.value, str):
                  value = cell.value.strip()

                  if value:
                      texts.add(value)

  return texts

separator = '======'

def print_translate(path: str):
    texts = translate_excel(path)
    for text in sorted(texts):
        # text + 换行 + separator 保存到 1.txt 文件中
        with open("1.txt", "a") as f:
            f.write(text + "\n" + separator + "\n")
        # print(text)
        # print(separator)

def loadData(path: str):
    with open(path, 'r') as f:
        content = f.read()
    
    separator = "======"
    parts = content.split(separator)
    
    result = {}
    number = 1
    for part in parts:
        text = part.strip()
        if text:
            result[number] = text
            number += 1
    
    return result

def get_translate_map(source_path: str, target_path: str):
    # Load the source map
    source_map = loadData(source_path)
    # Load the target map
    target_map = loadData(target_path)
    
    # 遍历 source_map
    combined_map = {}
    for key, value in source_map.items():
        combined_map[value] = target_map.get(key, value)


    # 获取 combined_map 的大小
    print(f"source_map map size: {len(source_map)}")
    print(f"target_map map size: {len(target_map)}")
    print(f"Combined map size: {len(combined_map)}")

    return combined_map

_T_PATTERN = re.compile(r'(<t(?:\s[^>]*)?>)([^<>]*)(</t>)')
_SHEET_NAME_PATTERN = re.compile(r'(<sheet\b[^>]*?\sname=")([^"]*)(")')


def _normalize(content: str) -> str:
    return content.replace('\r\n', '\n').replace('\r', '\n').strip()


def replace_t_text(xml_text: str, escaped_map: dict) -> str:
    def repl(m):
        open_tag, content, close_tag = m.group(1), m.group(2), m.group(3)
        normalized = _normalize(content)
        if normalized in escaped_map:
            return open_tag + escaped_map[normalized] + close_tag
        return m.group(0)
    return _T_PATTERN.sub(repl, xml_text)


def replace_sheet_names(xml_text: str, escaped_map: dict) -> str:
    def repl(m):
        prefix, content, suffix = m.group(1), m.group(2), m.group(3)
        normalized = _normalize(content)
        if normalized in escaped_map:
            return prefix + escaped_map[normalized] + suffix
        return m.group(0)
    return _SHEET_NAME_PATTERN.sub(repl, xml_text)


def excel_cell_replace(source_path: str, target_path: str, path: str):
    """
    Replace the text in an xlsx at the zip/XML level, preserving images
    (including wmf), formulas, styles, and rich-text run formatting.
    """
    translate_map = get_translate_map(source_path, target_path)
    escaped_map = {html.escape(k, quote=False): html.escape(v, quote=False)
                   for k, v in translate_map.items()}

    # 获取 path 中原文件名后再添加"-translated"作为新的文件名
    file_name = os.path.splitext(os.path.basename(path))[0] + "-translated.xlsx"
    print(f"Saving translated file to {file_name}")

    with zipfile.ZipFile(path, "r") as zin, \
         zipfile.ZipFile(file_name, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            is_strings = item.filename == "xl/sharedStrings.xml"
            is_sheet = (item.filename.startswith("xl/worksheets/")
                        and item.filename.endswith(".xml"))
            is_workbook = item.filename == "xl/workbook.xml"
            if is_strings or is_sheet:
                text = data.decode("utf-8")
                text = replace_t_text(text, escaped_map)
                data = text.encode("utf-8")
            elif is_workbook:
                text = data.decode("utf-8")
                text = replace_sheet_names(text, escaped_map)
                data = text.encode("utf-8")
            zout.writestr(item, data)

if __name__ == "__main__":
    path = "./excel/translate.xlsx"
    # 获取翻译的 excel 文本到 1.txt 中
    print_translate(path)

    # 大模型翻译，提示词：将 @1.txt 的文本由中文翻译成英文，存储到 2.txt 文件中

    # 替换 Excel 中的文本
    # excel_cell_replace("./1.txt", "./2.txt", path)
    print('Completed')

    # 注意：图片中的文字无法翻译