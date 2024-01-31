import os
import pandas as pd
from datetime import datetime
from pytz import timezone

DATA = pd.ExcelFile(os.path.join("ШК отправка.xlsx"))
pd_file = pd.read_excel(DATA)
art_list = pd_file['Артикул'].tolist()
shk_list = pd_file['Штрихкод'].tolist()
# print(art_list)
# for i, row in pd_file.iterrows():
#     art = str(row["Артикул"])
#     shk = str(row["Штрихкод"])
data = {'Код': [899787930149, 8997879301494, 8997879351499, 458027058618, 4580270586184],
        'Наименование': ['F1111_Sakura', 'F1111_Sakura', 'F1111_Sakura', 'PN8808', 'PN8808']}

df = pd.DataFrame(pd_file)
new_df = df.groupby('Артикул')['Штрихкод'].apply(list).reset_index()

new_df.to_excel('result_shk.xlsx', index=False)