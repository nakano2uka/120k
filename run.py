from datetime import datetime, date
import pathlib
import docx

yyyymmdd = date.today().strftime("%Y/%m/%d")
docx_path = pathlib.Path(__file__).parent / "all.docx" 
txt_path = pathlib.Path(__file__).parent / "all.txt"
csv_path = pathlib.Path(__file__).parent / "statistic.csv"

doc = docx.Document(docx_path)
num_of_char = 0
with open(txt_path, "w", encoding='utf-8') as f:
  for para in doc.paragraphs:
    encoded_text = para.text
    print(para.text, file=f)
    num_of_char += len(para.text)
    "{:,}".format(num_of_char)

with open(csv_path, "a") as f:
  add_content = f'{yyyymmdd},{num_of_char}'
  print(add_content, file=f)