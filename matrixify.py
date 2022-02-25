import pandas as pd


class Matrixify:

    def __init__(self):
        pass

    @classmethod
    def read_source(cls, filename):
        '''Read source xls file exported from Matrixify into separate dataframes'''
        xls = pd.ExcelFile(filename)
        data = {}
        for sheet in xls.sheet_names:
            data[sheet] = pd.read_excel(xls, sheet)
        return data

    @classmethod
    def build_output(cls, df, schema, filename):
        '''Write output to xls to import via Matrixify'''
        xls_writer = pd.ExcelWriter(filename)
        for key in schema.keys():
            df[schema[key]].to_excel(xls_writer, key, index=False)
        xls_writer.save()
