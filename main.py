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
    brand_dict = {}

    for i, row in df.iterrows():
        brand = row['Name']
        model = row['VM']
        engine = row['Engines']
        engine_capacity = row['TypeName']
        hp = row['HorsePowers']
        year = row['Year']

        if brand not in brand_dict:
            brand_dict[brand] = {}
        if model not in brand_dict[brand]:
            brand_dict[brand][model] = {}
            if '-' in year:
                year_start, year_end = year.split('-')
                if year_start != '':
                    brand_dict[brand][model]["start_date"] = year_start
                else:
                    brand_dict[brand][model]["start_date"] = year
                if year_end != '':
                    brand_dict[brand][model]["end_date"] = year_end
                else:
                    brand_dict[brand][model]["end_date"] = year
            else:
                year_start = year_end = year
                brand_dict[brand][model]["start_date"] = year_start
                brand_dict[brand][model]["end_date"] = year_end
        else:
            if '-' in year:
                year_start, year_end = year.split('-')
                if year_start != '' and year_start < brand_dict[brand][model]["start_date"]:
                    brand_dict[brand][model]["start_date"] = year_start
                if year_end != '' and year_end > brand_dict[brand][model]["end_date"]:
                    brand_dict[brand][model]["end_date"] = year_end
            else:
                year_start = year_end = year
                if year_start < brand_dict[brand][model]["start_date"]:
                    brand_dict[brand][model]["start_date"] = year_start
                if year_end > brand_dict[brand][model]["end_date"]:
                    brand_dict[brand][model]["end_date"] = year_end

        if engine_capacity not in brand_dict[brand][model]:
            brand_dict[brand][model][engine_capacity] = {}
        if engine not in brand_dict[brand][model][engine_capacity]:
            brand_dict[brand][model][engine_capacity][engine] = hp

    with open('brand_dict.json', 'w', encoding='utf-8') as f:
        json.dump(brand_dict, f, ensure_ascii=False)

    new_df = pd.DataFrame()
    prev_brand = None
    prev_model = None
    for brand in brand_dict.keys():
        for model in brand_dict[brand].keys():
            start_date = brand_dict[brand][model]["start_date"]
            end_date = brand_dict[brand][model]["end_date"]
            for engine_cap in brand_dict[brand][model].keys():

                if isinstance(brand_dict[brand][model][engine_cap], dict):
                    engine_list = [engines for engines in brand_dict[brand][model][engine_cap].keys()]
                    hp = [brand_dict[brand][model][engine_cap][hp] for hp in brand_dict[brand][model][engine_cap].keys()]
                    if brand != prev_brand:
                        brand_value = brand
                    else:
                        brand_value = ''
                    if model != prev_model:
                        start_date = start_date[-2:] + "." + start_date[:-2]
                        end_date = end_date[-2:] + "." + end_date[:-2]
                        model_value = f'{model} {start_date}-{end_date}'
                    else:
                        model_value = ''
                    df = pd.DataFrame.from_dict({'МОДЕЛЬ': [brand_value, model_value, engine_cap],
                                                 'КОД ДВИГАТЕЛЯ': ['', '', ', '.join(engine_list)],
                                                 'Мощность Л.С': ['', '', ', '.join(hp)],
                                                 'Name': ['', '', brand],
                                                 'VM': ['', '', model],
                                                 'TypeName': ['', '', engine_cap],
                                                 }, orient='index')
                    df = df.transpose()
                    new_df = pd.concat([new_df, df], ignore_index=True)
                    prev_brand = brand
                    prev_model = model
    pd.set_option('display.width', 500)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    out_df = pd.read_excel('result_colum_category.xlsx')
    out_df = out_df.drop(columns=["HorsePowers"])
    out_df = out_df.drop_duplicates(subset=['Name', 'VM', 'TypeName'])

    new_df = new_df.merge(out_df, how='left', on=["Name", "VM", "TypeName"])
    new_df = new_df.drop(columns=["Name"])
    new_df = new_df.drop(columns=["VM"])
    new_df = new_df.drop(columns=["TypeName"])
    new_df = new_df.drop(columns=["Engines"])
    new_df.rename(columns={'Фильтр, воздух во внутренном пространстве': 'Салонный фильтр'}, inplace=True)
    new_df = new_df.drop(columns=["Year"])
    new_df = new_df[new_df['МОДЕЛЬ'] != '']
    # new_df.style.set_properties(subset=['МОДЕЛЬ'], **{'font-weight': 'bold'})
    df.style.applymap(lambda x: 'font-weight: bold', subset=['МОДЕЛЬ'])
    print(new_df.head(10))
    new_df.to_excel("res.xlsx", index=False)


if __name__ == '__main__':
    # get_category()
    change_colum()

