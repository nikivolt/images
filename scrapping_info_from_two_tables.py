"""
Todo list:
1) rename columns [done]
2) reoder columns
3) remove the whole row at blank lines in column "Описание"
"""

import pandas as pd

elect_cols = [0, 1, 3, 4, 5, 6, 9, 17, 21]
schneider_cols = [1, 3, 4, 5, 6, 10, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]

elect_data = pd.read_excel('from_elect.xlsx', skiprows=2, usecols=elect_cols, index_col=0)
schneider_data = pd.read_excel('web_import.xlsx', usecols=schneider_cols, index_col=0)

# elect_data.to_excel('from_elect_small.xlsx')
# schneider_data.to_excel('web_import_small.xlsx')

# combine_type = inner, left, right, outer
df = pd.merge(elect_data, schneider_data, on='Partnumber', how='outer')

"""

nan_value = float("NaN")

df.replace("", nan_value, inplace=True)
df.dropna(subset=["Описание"], inplace=True)

"""

# print(df.columns)

df = df.rename(columns={"Partnumber": "Референция",
                        "Описание ": "Описание",
                        0.65: "Покупна цена\nбез ДДС",
                        "Цена след 010721\n лв.без ДДС": "Базова цена\nбез ДДС",
                        "CatalogName": "Продуктова гама",
                        "Category": "Категория",
                        "LinkOfImage": "Линк към снимка",
                        "Price": "iVolt цена",
                        "Description": "Уеб описание",
                        "Length(mm)": "Дължина(мм)",
                        "Width(mm)": "Ширина(мм)",
                        "Depth(mm)": "Дълбочина(мм)",
                        "Weight(kg)": "Тегло(кг)",
                        "VolumeWeight(m3)": "Обемно тегло(м3)",
                        "Promotion": "Промоция",
                        "NewProduct": "Нов продукт",
                        "HiddenProduct": "Скрит продукт",
                        "Attachment": "Прикачен файл",
                        "MetaTitle": "Мета заглавие",
                        "MetaDescription": "Мета описание"})

print(df.columns)

df.to_excel('test.xlsx')
