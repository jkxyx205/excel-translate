import os

from openpyxl import load_workbook


def translate_excel(path: str):
  """
    Extract all unique text values from an Excel file.
  """
  wb = load_workbook(path)
  texts = set()

  for sheet in wb.worksheets:
      for row in sheet.iter_rows():
          for cell in row:
              if isinstance(cell.value, str):
                  value = cell.value.strip()

                  if value:
                      texts.add(value)

  return texts

def print_translate(path: str):
    texts = translate_excel(path)
    for text in sorted(texts):
        print(text)
        print('======')

separator = '======'

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
        # 判断 key 是否包含字符串 “3. 5S要求：”
        # print(f"Processing key: {key}, value: {value}")
        # if "3. 5S要求：" in value:
            # print("====value '3. 5S要求：' exists in the combined map.")  
    # print(source_map)


    # 获取 combined_map 的大小
    print(f"source_map map size: {len(source_map)}")
    print(f"target_map map size: {len(target_map)}")
    print(f"Combined map size: {len(combined_map)}")

    return combined_map

def excel_cell_replace(source_path: str, target_path: str, path: str):
  """
    Replace the text in an Excel file with the translated text from the target map.
  """
  translate_map = get_translate_map(source_path, target_path)

  wb = load_workbook(path)
  for sheet in wb.worksheets:
      for row in sheet.iter_rows():
          for cell in row:
              if isinstance(cell.value, str):
                  value = cell.value.strip()

                  if value:
                      # 设置 cell 的值
                      cell.value = translate_map.get(value, value)

  # 获取 path 中原文件名后再添加"-translated"作为新的文件名
  file_name = os.path.splitext(os.path.basename(path))[0] + "-translated.xlsx"
  print(f"Saving translated file to {file_name}")
  wb.save(file_name)  

if __name__ == "__main__":
    path = "./excel/translate.xlsx"
    
    # 获取翻译的 excel 文本，复制到 1.txt 中
    # print_translate(path)
    # 大模型翻译，提示词：将 @1.txt 的文本由中文翻译成英文，存储到 2.txt 文件中
    # target_map = loadData("./2.txt")
    excel_cell_replace("./1.txt", "./2.txt", path)
    print('Completed')