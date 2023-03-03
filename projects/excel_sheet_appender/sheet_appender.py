#!/usr/bin/env python3
"""
Autored by Venkatesh 
"""
import os
import sys
import pandas as pd


class SheetAppender:
    def __init__(self, path, sheet_name):
        self.path = path
        self.sheet_name = sheet_name
        self.sheets = []
        if not os.path.exists(self.path):
            print(f"The path {self.path} is invalid")
            sys.exit(-1)
        cwd = os.getcwd()
        output_file = os.path.join(cwd, "output_file.xlsx")
        self.writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

    def parse_data(self):
        files = os.listdir(self.path)
        print(f"files {files} are available at the directory{self.path}")

        for file in files:
            self.sheets.append(pd.read_excel(os.path.join(self.path, file),
                                             sheet_name=self.sheet_name,
                                             engine='openpyxl'))
        print(self.sheets)

    def write_to_excel(self):
        # writing datas to output excel file
        row_count = 0
        workbook = self.writer.book
        worksheet = workbook.add_worksheet('desired_sheet_name')
        self.writer.sheets['desired_sheet_name'] = worksheet
        for df in self.sheets:
            df.to_excel(self.writer, sheet_name='desired_sheet_name', startrow=row_count, startcol=0)
            row_count += (len(df) + 2)
        self.writer.close()


if __name__ == "__main__":
    APPENDER = SheetAppender(path=r"C:\path\to\your\file\directory",
                             sheet_name="Sheet1")
    APPENDER.parse_data()
    APPENDER.write_to_excel()
