import pandas as pd


def get_category():
    res_df = pd.read_csv('vehicles.csv', delimiter=';')
    res_df = res_df.sort_values(['VEH_BRAND', 'VEH_TYPE_NO'])
    res_df['VEH_MODEL_NO'] = res_df['VEH_MODEL_NO'].apply(lambda x: str(int(x)) if not pd.isna(x) else '')

    res_df.to_csv("result_vehicles.csv", index=False, sep=';')


get_category()