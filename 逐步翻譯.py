import pandas as pd
import os
from translate import Translator

# 初始化Translator，将文本从中文翻译到英文
translator = Translator(to_lang='en', from_lang='zh')


def translate_text(text):
    try:
        text = str(text)
        translated = translator.translate(text)
        return translated
    except Exception as e:
        print(f"翻译时发生错误: {e}")
        return ""

# Excel文件
source_folder = "C:\\Users\\teter\\Desktop\\翻譯檔"
excel_files = [f for f in os.listdir(source_folder) if f.startswith("excel") and f.endswith(".xlsx")]

# 分別處理
for excel_file in excel_files:
    file_path = os.path.join(source_folder, excel_file)
    df = pd.read_excel(file_path)

    # 假设要翻译的内容在第三列，如果不是，请更改列索引
    df['Translated_Text'] = df.iloc[:, 2].apply(translate_text)
    
    # 新的文件名字
    base_name = excel_file.replace("excel", "eng").replace(".xlsx", "")
    output_file = f"{source_folder}\\{base_name}.xlsx"

    df.to_excel(output_file, index=False)
    print(f"文件 {output_file} 已保存。")

print("翻譯完畢記得檢查。")
