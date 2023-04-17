from pathlib import  Path
import pandas as pd
from docxtpl import DocxTemplate
base_dir=Path(__file__).parent
word_template= base_dir/"word template name"
excel_path=base_dir/"excel_name"
output_dir=base_dir /"folder_name"


output_dir.mkdir(exist_ok=True)
df=pd.read_excel(excel_path,sheet_name="sheet_name",engine='openpyxl')
for record in df.to_dict(orient="record"):
  doc=DocxTemplate(word_template)
  doc.render(record)
  output_path=output_dir/f"{record['word_name detected from excel coloumn']}.docx"
  doc.save(output_path)
