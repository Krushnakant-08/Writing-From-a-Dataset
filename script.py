import pandas as pd
from docxtpl import DocxTemplate

df = pd.read_excel("Sample.xlsx")

template_path = "student.docx"

nam = input('Enter Name of Student :')

for index, row in df.iterrows():
    if row['Name'] == nam:
      context = {
          'Name': row['Name'],
          'MotherName': row['MotherName'],
          'StudentID': row['StudentID'],
          'Gen_Reg_no': row['Gen_Reg_no'],
          'Class': row['Class'],
          'DOB': str(row['DOB']),
      }
      doc = DocxTemplate(template_path)
      doc.render(context)

      output_path = f"Reg_{row['StudentID']}.docx"
      doc.save(output_path)
      print(f"Saved: {output_path}")
