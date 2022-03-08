'''
This script takes a Matrixify export file with active products, filter those that was last updated more than
XY days ago and are currently hidden and deletes them. It also creates pages with the same handle and the same
content as deleted products and sets redirects.

Output is an Excel file to be imported by Matrixify.

Necessary columns in source file are:
ID, Handle, Command, Title, Body HTML, Tags, Variant SKU, Image Src, Image Position and all Variant Metafields.

Proceed with caution!
'''

import pandas as pd
import datetime as dt
import json
import os

from matrixify import Matrixify as XLS


BESTSELLER_PREFIX = {
    'cz': 'nejprodavanejsi',
    'sk': 'najpredavanejsie'
}


def get_reduced_df(df, column, value):
    '''Reduce source dataframe of pages to only archived products'''
    mask = df.apply(lambda row: row[column] == value, axis=1)
    return df[mask]


def get_last_update(row, today, delta_days):
    '''Return if this product was last updated more than XY days ago'''
    all_tags = row['Tags'].split(',')
    upd_tags = [tag for tag in all_tags if 'UPD:' in tag]
    if len(upd_tags) == 0:
        upd_tags = [tag for tag in all_tags if 'ADD:' in tag]
    threshold = today - dt.timedelta(days=delta_days)
    last_upd_time = None

    for tag in upd_tags:
        tag_time = dt.datetime.strptime(tag.split(':')[1][:10], '%Y-%m-%d').date()
        if last_upd_time == None or tag_time > last_upd_time:
            last_upd_time = tag_time

    if last_upd_time != None and last_upd_time < threshold:
        return True

    return False


def get_main_collection_handle(product_tags, handle_list, bf_prefix):
    '''Return collection handle based on MCI product tag'''
    for tag in product_tags.split(','):
        if 'MCI' in tag:
            col_id = tag.split('MCI:')[1]
            if col_id in handle_list.keys():
                return handle_list[col_id], f'{bf_prefix}-{handle_list[col_id]}'
    return '', ''
    

def run(handle_list, lang, bf_prefix):

    # read source xls file
    source = XLS.read_source(f'source_{lang}.xlsx')
    products = source['Products'].fillna('')

    # group products by id to get list of all images
    grouped = products[['ID', 'Image Src']].groupby('ID')
    grouped_images = grouped['Image Src'].apply(list).to_frame()
    grouped_images = grouped_images.rename(columns={'Image Src': 'Addtl Images'})

    primary_img_mask = products['Image Position'] == 1
    products = products[primary_img_mask]
    print(products.shape)

    products = pd.merge(
        left=products,
        right=grouped_images['Addtl Images'],
        left_on='ID',
        right_on='ID',
        how='left'
    )

    # get only products that have tags PRD:Hidden
    mask = products['Tags'].str.contains('PRD:Hidden')
    hidden = products[mask]

    # get only products where last update is 60 days ago
    today = dt.date.today()
    hidden['mask'] = hidden.apply(lambda row: get_last_update(row, today, 60), axis=1)
    to_archive = get_reduced_df(hidden, 'mask', True)

    # build new output dataframe
    output = to_archive[['ID', 'Command', 'Handle', 'Title', 'Body HTML']]
    output['Command'] = 'DELETE'
    output['Template Suffix'] = 'archived-goods'
    output['Metafield: mf_pg_ap.Image_Src [string]'] = to_archive['Image Src']
    output['Metafield: mf_pg_ap.Addtl_Images [string]'] = to_archive.apply(lambda row: ';'.join(row['Addtl Images'][1:]), axis=1)
    output['Metafield: mf_pg_ap.Shpsys_ID [integer]'] = to_archive['Variant Metafield: mf_pvp.MKT_ID_SHOPSYS [number_integer]']
    output['Metafield: mf_pg_ap.Variant SKU [string]'] = to_archive['Variant SKU']
    output['Metafield: mf_pg_ap.main_category [string]'] = to_archive.apply(lambda row: get_main_collection_handle(row['Tags'], handle_list, bf_prefix)[0], axis=1)
    output['Metafield: mf_pg_ap.related_products_col [string]'] = to_archive.apply(lambda row: get_main_collection_handle(row['Tags'], handle_list, bf_prefix)[1], axis=1)
    output['Metafield: mf_pg_ap.SHPF_BENEFITS [string]'] = to_archive['Variant Metafield: mf_pvp.SHPF_BENEFITS [multi_line_text_field]']
    output['Metafield: mf_pg_ap.SHPF_SHORT_DESCRIPTION [string]'] = to_archive['Variant Metafield: mf_pvp.SHPF_SHORT_DESCRIPTION [multi_line_text_field]']
    output['Path'] = to_archive.apply(lambda row: f'/products/{row["Handle"]}', axis=1)
    output['Target'] = to_archive.apply(lambda row: f'/pages/{row["Handle"]}', axis=1)
    print(output.shape)

    output_schema = {
        'Products': ['ID', 'Command', 'Handle', 'Title'],
        'Pages': ['Handle', 'Title', 'Body HTML', 'Template Suffix', 'Metafield: mf_pg_ap.Image_Src [string]', 
                'Metafield: mf_pg_ap.Addtl_Images [string]', 'Metafield: mf_pg_ap.Shpsys_ID [integer]', 
                'Metafield: mf_pg_ap.Variant SKU [string]', 'Metafield: mf_pg_ap.main_category [string]', 
                'Metafield: mf_pg_ap.related_products_col [string]', 'Metafield: mf_pg_ap.SHPF_BENEFITS [string]', 
                'Metafield: mf_pg_ap.SHPF_SHORT_DESCRIPTION [string]'],
        'Redirects': ['Path', 'Target']
    }

    XLS.build_output(output, output_schema, f'output_{lang}.xlsx')


if __name__ == '__main__':
    pd.options.mode.chained_assignment = None

    all_files = list(filter(lambda file: os.path.isfile(file), os.listdir()))
    filenames = [file for file in all_files if file.startswith('source')]
    if len(filenames) == 0:
        raise Exception('Wrong source filename format, use source_cz.xlsx or source_sk.xlsx only.')

    for filename in filenames:
        lang = filename.split('.')[0].split('_')[1].lower()
        if not lang:
            raise Exception('Wrong source filename format, use source_cz.xlsx or source_sk.xlsx only.')
            
        with open(f'collections_{lang}.json', encoding='utf8') as file:
            handle_list = json.load(file)

        run(handle_list, lang, BESTSELLER_PREFIX[lang])
