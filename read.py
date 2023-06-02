# encoding='utf-8

# @Time: 2022-09-22
# @File: %
#!/usr/bin/env
import pandas as pd
import polars as pl
from icecream import ic
import os
os.chdir(os.path.abspath(os.path.dirname(__file__)))
# change cwd to current file dir
# df = pl.read_excel('装箱单.xlsx', sheet_name='template',
#                    # read_csv_options={'ignore_errors':True,'skip_rows':0,'dtypes':{} }
#                    )
# # ic(df)
# dft = df.pivot(index=['SKU'], columns=['箱号'],
#                values='数量', aggregate_fn='sum',)
# dft.write_csv('file_name_1.csv')
# # df_lb = df.select([pl.col('箱号'),
# #                    pl.col('磅'),
# #                    pl.col('SKU'),
# #                    pl.col('箱型'),
# #                    # pl.col('col_name').cast.alias('col_name'),
# #                    ])
# df_lb = df.clone()
# df_lb = df_lb.unique(maintain_order=True, subset=['箱号'], keep='first')
# df_lb_t = df_lb.pivot(index=['SKU'], columns=['箱号'],
#                       values='磅', aggregate_fn='sum',).sum()
# df_lb_t[0, 'SKU'] = "包装箱重量（磅）："
# df_lb_size = df_lb.unique(maintain_order=True, subset=['箱号'], keep='first')
# df_lb_size = df_lb_size.pivot(index=['SKU'], columns=['箱号'],
#                               values='箱型', aggregate_fn='sum',).sum()
# df_lb_size[0, 'SKU'] = "包装箱型号："
#
#
# ic(df_lb_t)
# ic(df_lb_size)

import openpyxl

wb = openpyxl.load_workbook('装箱单.xlsx')
ws = wb['template']
ws['A1'] = 'hello world'

wb.save('save.xlsx')
