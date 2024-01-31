import pandas as pd
import csv


def get_category():
    res_df = pd.read_csv('Archive/article_vehicle_links.csv', delimiter=';')
    res_df = res_df.sort_values(['ART_ID', 'VEH_TYPE_NO', 'ART_NUM'])
    # res_df['VEH_MODEL_NO'] = res_df['VEH_MODEL_NO'].apply(lambda x: str(int(x)) if not pd.isna(x) else '')

    res_df.to_csv("article_vehicle_links_Archive.csv", index=False, sep=';', quoting=csv.QUOTE_NONE)


get_category()