import pandas as pd
import datetime as dt
import json


BESTSELLER_PREFIX = 'najpredavanejsie'
with open('collections.json', encoding='utf8') as file:
    HANDLE_LIST = json.load(file)


def read_source_xlsx(filename):
    '''Read source xls file into separate dataframes'''
    xls = pd.ExcelFile(filename)
    data = {}
    for sheet in xls.sheet_names:
        data[sheet] = pd.read_excel(xls, sheet)
    return data


def build_output_xlsx(df):
    '''Write output to xls to import via Matrixify'''
    xls_writer = pd.ExcelWriter('output.xlsx')
    df[['ID', 'Command', 'Handle', 'Title']].to_excel(xls_writer, 'Products', index=False)
    df[
        [
            'Handle', 
            'Title', 
            'Body HTML',
            'Template Suffix', 
            'Metafield: mf_pg_ap.Image_Src [string]', 
            'Metafield: mf_pg_ap.Addtl_Images [string]', 
            'Metafield: mf_pg_ap.Shpsys_ID [integer]', 
            'Metafield: mf_pg_ap.Variant SKU [string]', 
            'Metafield: mf_pg_ap.main_category [string]', 
            'Metafield: mf_pg_ap.related_products_col [string]', 
            'Metafield: mf_pg_ap.SHPF_BENEFITS [string]', 
            'Metafield: mf_pg_ap.SHPF_SHORT_DESCRIPTION [string]'
        ]
    ].to_excel(xls_writer, 'Pages', index=False)
    df[['Path', 'Target']].to_excel(xls_writer, 'Redirects', index=False)
    xls_writer.save()


def get_reduced_df(df, column, value):
    '''Reduce source dataframe of pages to only archived products'''
    mask = df.apply(lambda row: row[column] == value, axis=1)
    return df[mask]


def get_last_update(row, today, delta_days):
    '''Return if this product was last updated more than XY days ago'''
    all_tags = row['Tags'].split(',')
    threshold = today - dt.timedelta(days=delta_days)
    for tag in all_tags:
        if 'UPD' in tag:
            upd_time = dt.datetime.strptime(tag.split('UPD:')[1], '%Y-%m-%d-%H-%M-%S')
            if upd_time < threshold:
                return True
    return False


def get_main_collection_handle(product_tags):
    '''Return collection handle based on MCI product tag'''
    for tag in product_tags.split(','):
        if 'MCI' in tag:
            col_id = tag.split('MCI:')[1]
            if col_id in HANDLE_LIST.keys():
                return HANDLE_LIST[col_id], f'{BESTSELLER_PREFIX}-{HANDLE_LIST[col_id]}'
    return '', ''
    

def main():
    # read source xls file
    source = read_source_xlsx('source.xlsx')
    products = source['Products'].fillna('')

    # group products by id to get list of all images
    grouped = products[['ID', 'Image Src']].groupby('ID')
    grouped_images = grouped['Image Src'].apply(list).to_frame()
    grouped_images = grouped_images.rename(columns={'Image Src': 'Addtl Images'})

    primary_img_mask = products['Image Position'] == 1
    products = products[primary_img_mask]

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

    # get only products where last update is 30 days ago
    today = dt.datetime.now()
    hidden['mask'] = hidden.apply(lambda row: get_last_update(row, today, 30), axis=1)
    to_archive = get_reduced_df(hidden, 'mask', True)

    # build new output dataframe
    output = to_archive[['ID', 'Command', 'Handle', 'Title', 'Body HTML']]
    output['Command'] = 'DELETE'
    output['Template Suffix'] = 'archived-goods'
    output['Metafield: mf_pg_ap.Image_Src [string]'] = to_archive['Image Src']
    output['Metafield: mf_pg_ap.Addtl_Images [string]'] = to_archive.apply(lambda row: ';'.join(row['Addtl Images'][1:]), axis=1)
    output['Metafield: mf_pg_ap.Shpsys_ID [integer]'] = to_archive['Variant Metafield: mf_pvp.MKT_ID_SHOPSYS [number_integer]']
    output['Metafield: mf_pg_ap.Variant SKU [string]'] = to_archive['Variant SKU']
    output['Metafield: mf_pg_ap.main_category [string]'] = to_archive.apply(lambda row: get_main_collection_handle(row['Tags'])[0], axis=1)
    output['Metafield: mf_pg_ap.related_products_col [string]'] = to_archive.apply(lambda row: get_main_collection_handle(row['Tags'])[1], axis=1)
    output['Metafield: mf_pg_ap.SHPF_BENEFITS [string]'] = to_archive['Variant Metafield: mf_pvp.SHPF_BENEFITS [multi_line_text_field]']
    output['Metafield: mf_pg_ap.SHPF_SHORT_DESCRIPTION [string]'] = to_archive['Variant Metafield: mf_pvp.SHPF_SHORT_DESCRIPTION [multi_line_text_field]']
    output['Path'] = to_archive.apply(lambda row: f'/products/{row["Handle"]}', axis=1)
    output['Target'] = to_archive.apply(lambda row: f'/pages/{row["Handle"]}', axis=1)

    build_output_xlsx(output)


if __name__ == '__main__':
    main()
