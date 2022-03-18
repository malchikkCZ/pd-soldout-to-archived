'''
This app takes a Matrixify export file with active products, filter those that was last updated more than
XY days ago and are currently hidden and deletes them. It also creates pages with the same handle and the same
content as deleted products and sets redirects.

Output is an Excel file to be imported by Matrixify.

Necessary columns in source file are:
ID, Handle, Command, Title, Body HTML, Tags, Variant SKU and all Variant Metafields.

Proceed with caution!
'''

import pandas as pd
import datetime as dt
import json
import os

from handleizer import Handleizer
from matrixify import Matrixify
from pim_connector import PimConnector
from secrets import HOST, USER, PASS


BESTSELLER_PREFIX = {
    'cz': 'nejprodavanejsi',
    'sk': 'najpredavanejsie'
}


class ProductArchiver:

    def __init__(self, handle_list, lang, prefix, delta_days=60):
        self.handle_list = handle_list
        self.lang = lang
        self.prefix = prefix
        self.delta_days = delta_days
        self.today = dt.date.today()
        self.source = Matrixify.read_source(f'source_{self.lang}.xlsx')

        self.pim = PimConnector(
            host=HOST,
            user=USER,
            password=PASS
        )
        self.galery = self.pim.get_df_from_table('galery')

    def run(self):
        # read source xls file
        products = self.source['Products'].fillna('')
        drop_indexes = products[(products['Variant Metafield: mf_pvp.MKT_ID_SHOPSYS [number_integer]'] == '')].index
        products.drop(drop_indexes, inplace=True)
        print(products.shape)

        # get only products that have tags PRD:Hidden
        mask = products['Tags'].str.contains('PRD:Hidden')
        hidden = products[mask]

        # get only products where last update is 60 days ago
        hidden['mask'] = hidden.apply(lambda row: self.get_last_update(row), axis=1)
        to_archive = self.get_reduced_df(hidden, 'mask', True)

        # build new output dataframe
        output = to_archive[['ID', 'Command', 'Handle', 'Title', 'Body HTML']]
        output['Command'] = 'DELETE'
        output['Template Suffix'] = 'archived-goods'
        output['Metafield: mf_pg_ap.Image_Src [string]'] = to_archive.apply(lambda row: self.get_images_srcs(row, 0, 1), axis=1)
        output['Metafield: mf_pg_ap.Addtl_Images [string]'] = to_archive.apply(lambda row: self.get_images_srcs(row, 1), axis=1)
        output['Metafield: mf_pg_ap.Shpsys_ID [integer]'] = to_archive['Variant Metafield: mf_pvp.MKT_ID_SHOPSYS [number_integer]']
        output['Metafield: mf_pg_ap.Variant SKU [string]'] = to_archive['Variant SKU']
        output['Metafield: mf_pg_ap.main_category [string]'] = to_archive.apply(lambda row: self.get_main_collection_handle(row['Tags'])[0], axis=1)
        output['Metafield: mf_pg_ap.related_products_col [string]'] = to_archive.apply(lambda row: self.get_main_collection_handle(row['Tags'])[1], axis=1)
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

        Matrixify.build_output(output, output_schema, f'output_{lang}.xlsx')

    def get_reduced_df(self, df, column, value):
        '''Reduce source dataframe of pages to only archived products'''
        mask = df.apply(lambda row: row[column] == value, axis=1)
        return df[mask]

    def get_last_update(self, row):
        '''Return if this product was last updated more than XY days ago'''
        all_tags = row['Tags'].split(',')
        upd_tags = [tag for tag in all_tags if 'UPD:' in tag]
        if len(upd_tags) == 0:
            upd_tags = [tag for tag in all_tags if 'ADD:' in tag]
        threshold = self.today - dt.timedelta(days=self.delta_days)
        last_upd_time = None

        for tag in upd_tags:
            tag_time = dt.datetime.strptime(tag.split(':')[1][:10], '%Y-%m-%d').date()
            if last_upd_time == None or tag_time > last_upd_time:
                last_upd_time = tag_time

        if last_upd_time != None and last_upd_time < threshold:
            return True

        return False

    def get_main_collection_handle(self, product_tags):
        '''Return collection handle based on MCI product tag'''
        for tag in product_tags.split(','):
            if 'MCI' in tag:
                col_id = tag.split('MCI:')[1]
                if col_id in self.handle_list.keys():
                    return handle_list[col_id], f'{self.prefix}-{self.handle_list[col_id]}'
        return '', ''

    def get_images_srcs(self, row, start=0, limit=9999): 
        id = row['Variant Metafield: mf_pvp.MKT_ID_SHOPSYS [number_integer]']
        handleized_title = Handleizer.run(row['Title'])

        images = self.galery.copy()
        drop_indexes = images[~(images['good'] == int(id))].index
        images.drop(drop_indexes, inplace=True)
        images.sort_values(by=['pos'])
        
        images_srcs = []
        for id, row in images.iterrows():
            image_src = f'https://img.okay.cz/gal/{handleized_title}-original-{row["id"]}.jpg'
            images_srcs.append(image_src)
        
        del images
        if len(images_srcs) == 0 or len(images_srcs) < start + 1:
            return ''

        end = start + limit
        return ';'.join(images_srcs[start:end])


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

        product_archiver = ProductArchiver(handle_list, lang, BESTSELLER_PREFIX[lang], 60)
        product_archiver.run()
