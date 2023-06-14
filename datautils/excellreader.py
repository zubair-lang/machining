import pandas as pd

class ExcelToGenerator:
    def __init__(self, file_name, sheet_name):
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.df = self._load_data()
    
    def _load_data(self):
        xl = pd.ExcelFile(self.file_name)
        df = xl.parse(self.sheet_name)
        return df

    def data_generator(self):
        for index, row in self.df.iterrows():
            yield row




