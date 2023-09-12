import json

import pandas as pd


def get_category():
    res_df = pd.read_excel('export_sakura.xlsx')
    # Убираем дубликаты по Марке, Модели двигателю и Номеру артикля
    res_df = res_df.drop_duplicates(subset=['Name', 'VM', "ArticleNumber", "Engines"])
    #
    name_df = res_df[["Name", "VM", "Engines", "TypeName", "HorsePowers", "Year"]].drop_duplicates(subset=['Name', 'VM', "Engines", "TypeName"])
    print(name_df)

    # Создаем список всех категорий артиклей
    articles = res_df[["GenericArticle"]].drop_duplicates(subset=['GenericArticle'])['GenericArticle'].tolist()

    # Из списка артиклей создаем поля
    for article in articles:
        res_df.insert(loc=len(name_df.columns), column=article, value=article)

    for i, row in res_df.iterrows():
        article = row["GenericArticle"]
        res_df.loc[i, row[article]] = row['ArticleNumber']
    res_df = res_df.drop_duplicates(subset=['Name', 'VM', "Engines", "Year", "HorsePowers", 'ArticleNumber'])

    # Проходим по начальному датафрейму и заполняем созданные столбцы артиклей
    data_dicts = res_df.to_dict('records')
    for d in data_dicts:
        for x in d.keys():
            if d[x] == x:
                d[x] = '*'
    result_df = pd.DataFrame(data_dicts)

    # Группируем артикли по 'Name', "VM", "Engines"
    result_df = result_df.groupby(['Name', "VM", "Engines"])[articles].agg(','.join).reset_index()
    r_df = name_df.merge(result_df, how='inner', on=["Name", "VM", "Engines"])
    r_df = r_df.reset_index()

    # Очищаем ячейки от * и ','
    for art in articles:
        r_df[art] = r_df[art].str.replace('*,', '')
        r_df[art] = r_df[art].str.replace(',*', '')
        r_df[art] = r_df[art].str.replace('*', '')

    # Проверка на дубликаты в ячейках
    for i, row in r_df.iterrows():
        for art in articles:
            if row[art]:
                x = len(list(set(row[art].split(','))))
                y = len(row[art].split(','))
                if x != y:
                    print(x, y)
    print(r_df)
    r_df = r_df.drop(columns=['index'])
    r_df.to_excel("result_colum_category.xlsx", index=False)


def change_colum():
    df = pd.read_excel('result_colum_category.xlsx')

    new_data = []
    brand_dict = {}

    for i, row in df.iterrows():
        brand = row['Name']
        model = row['VM']
        engine = row['Engines']
        engine_capacity = row['TypeName']
        hp = row['HorsePowers']

        # if brand not in brand_dict:
        #     brand_dict[brand] = {}
        # if model not in brand_dict[brand]:
        #     brand_dict[brand][model] = []
        # brand_dict[brand][model].append(engine)
        if brand not in brand_dict:
            brand_dict[brand] = {}
        if model not in brand_dict[brand]:
            brand_dict[brand][model] = {}
        if engine_capacity not in brand_dict[brand][model]:
            brand_dict[brand][model][engine_capacity] = []
        brand_dict[brand][model][engine_capacity].append(engine)

    with open('brand_dict.json', 'w') as f:
        json.dump(brand_dict, f)

    #     # приводим дату к нужному формату
    #     if '-' in year:
    #         year_gm_start, year_gm_end = year.split('-')
    #         start_year, start_month = year_gm_start[:4], year_gm_start[4:]
    #         end_year, end_month = year_gm_end[:4], year_gm_end[4:]
    #         year_start = f'{start_month}.{start_year}'
    #         year_end = f'{end_month}.{end_year}'
    #     else:
    #         year_start = year_end = f'{year[4:]}.{year[:4]}'
    #
    #     new_data.append([brand, model, engine, engine_capacity, hp, year_start, year_end])
    #
    # new_df = pd.DataFrame(new_data, columns=['Brand', 'Model', 'Engine', 'Engine_Capacity', 'Horsepower', 'Year_Start', 'Year_End'])
    # print(new_df)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # get_category()
    change_colum()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
