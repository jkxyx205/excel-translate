# Excel 内容翻译工具

## 如何变成 skill
```
 /document-skills:skill-creator 将该项目变成 skill，通过提示词完成翻译。比如：将文件 ./excel/translate.xlsx 由中文翻译成英文，新的文件保存到 ./translate-translated.xlsx 中。skill 用中文书写
```

## 通过 skill 翻译如何使用

创建文件夹 `excel` ，将要翻译的文件放入文件夹中 `translate.xlsx`
```
将文件 ./excel/translate.xlsx 由中文翻译成英文，新的文件保存到 ./translate-translated.xlsx 中
```

## 手动翻译如何使用

创建文件夹 `excel` ，将要翻译的文件放入文件夹中 `translate.xlsx`
### 1. 获取需要翻译的信息
```python
    path = "./excel/translate.xlsx"
    # 1. 获取翻译的 excel 文本到 1.txt 中
    print_translate(path)
```
### 2. 大模型翻译，提示词：将 @1.txt 的文本由中文翻译成英文，存储到 2.txt 文件中
### 3. Excel 翻译替换
```python
    # 3. 替换 Excel 中的文本
    excel_cell_replace("./1.txt", "./2.txt", path)
```