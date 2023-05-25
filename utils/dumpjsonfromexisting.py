import json
import pandas as pd

df = pd.read_excel("Расстановка1.xlsm")

dic = dict()
df = df.iloc[:, :3]

subject_order = ['Начальные классы','Русский язык и литература', 'Иностранный язык', 
                'Математика и информатика', 'Общественные науки',
                'Естественные науки', 'Технология', 'Искусство',
                'Физическая культура, ОБЖ', 'Курсы по выбору']

current_area = ''

for index, row in df.iterrows():
    area = str(row[0])
    if area != 'nan':
        current_area = area
    for i in subject_order:
        if current_area == i:
            dic[str(row[2])] = current_area
with open('subject_areas.json', 'w') as f:
    json.dump(dic, f)       
print(dic)

