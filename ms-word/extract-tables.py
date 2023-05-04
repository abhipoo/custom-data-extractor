from docx import Document
import pandas as pd
import os

files = os.listdir('data\\payslip\\')
df_final = pd.DataFrame()
for file in files:
    try:
        filename = 'data\\payslip\\{}'.format(file)
        document = Document(filename)

        for table in document.tables:
            df = [['' for i in range(len(table.columns))] for j in range(len(table.rows))]
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    if cell.text:
                        df[i][j] = cell.text
            df_raw = pd.DataFrame(df)
            #access i, j index of dataframe
            #df.loc[row_indexer,column_indexer] https://pandas.pydata.org/docs/user_guide/indexing.html
            data = {
                        'Name': df_raw.loc[1,1].strip(),
                        'Emp. No.': df_raw.loc[2,0].strip(),
                        'Location': df_raw.loc[2,4].strip(),
                        'Department': df_raw.loc[3,3].strip(),
                        'Designation': df_raw.loc[4,2].strip(),
                        'Grade': df_raw.loc[5,1].strip(),
                        'Bank A/c No.': df_raw.loc[7,1].strip(),
                        'PAN No.': df_raw.loc[7,3].strip(),
                        'AADHAR No. ': df_raw.loc[8,0].strip(),
                        'UAN No.': df_raw.loc[8,2].strip(),
                        'Filename': filename.replace('data\\payslip\\', '')
                    }
            #add extracted data to dataframe
            df_final = pd.concat([df_final, pd.DataFrame([data])], ignore_index=True)
    except:
        print('Error: {}'.format(filename))
#write to excel
df_final.to_excel('data\\payslip\\consolidated.xlsx', index=False)

