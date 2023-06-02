# encoding='utf-8

# @Time: 2022-09-24
# @File: %
#!/usr/bin/env
import getopt
import sys
from Genshipment import *
from icecream import ic
import polars as pl
from datetime import datetime
import pandas as pd
import os
os.chdir(os.path.abspath(os.path.dirname(__file__)))
# change cwd to current file dir


def main(argv):
    inputfile = ''
    packagetemplate = ''
    boxtemplate = ''
    packageoutfile = ''
    boxoutfile = ''
    DirPath = ''

    try:
        opts, args = getopt.getopt(argv[1:], "ha:i:o:t:O:T:D:", [
                                   "ifile=", "ofile=", "tfile=", "ofile=", "Tfile=", "Ofile=", "Dfile="])
    except getopt.GetoptError:
        print('test.py -i <inputfile> -o <outputfile> ')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            useage(argv[0])
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg
        elif opt in ("-o", "--ofile"):
            packageoutfile = arg
        elif opt in ("-t", "--tfile"):
            packagetemplate = arg
        elif opt in ("-O", "--Ofile"):
            boxoutfile = arg
        elif opt in ("-T", "--Tfile"):
            boxtemplate = arg
        elif opt in ("-D", "--Dfile"):
            DirPath = arg
            if '今天' in DirPath:
                DirPath = DirPath.replace(
                    '今天', datetime.today().strftime('%Y%m%d'))
    # print('输入的文件为：', inputfile)
    # print('输出的文件为：', outputfile)
    try:
        if os.path.exists(DirPath):
            pass
        else:
            os.mkdir(DirPath)
        packageoutfile = DirPath + '/' + packageoutfile
        boxoutfile = DirPath + '/' + boxoutfile
        ic(packageoutfile)
        ic(boxoutfile)
    except Exception as e:
        print(e)
    try:
        df = pl.read_excel(inputfile, sheet_name='template',
                           # read_csv_options={'ignore_errors':True,'skip_rows':0,'dtypes':{} }
                           )
        # drop empty rows
        if not df.is_empty():
            df = df.filter(
                ~pl.fold(
                    True,
                    lambda acc, s: acc & s.is_null(),
                    pl.all(),
                )
            )

        df_list = Read_files(df)
        dft = df_list[0]
        # dft.write_csv('dft.csv')
        SKU_NO = Gen_total_sku_NO(df)
        df_check = SKU_NO.join(dft, how='left', on='SKU', suffix='_x')
        df_box_info = Gen_Box_info(df)
        writer = pd.ExcelWriter(DirPath + "//简易装箱单.xlsx", engine='xlsxwriter')
        df_box_info.to_pandas().to_excel(writer, sheet_name='box', index=False)
        df_check.to_pandas().to_excel(writer, sheet_name='pacakge', index=False)
        writer.close()

        # df_check.write_csv(DirPath + '//装箱单简易信息.csv')
        # df_box_info.write_csv(DirPath + '//箱子信息.csv')
        Write_package_info(packagetemplate, SKU_NO, packageoutfile)
        df_joined = Gen_Detail_package(boxtemplate, df_list, inputfile)
        # ic(df_joined)
        Write_box_info(boxtemplate, df_joined, boxoutfile)
    except Exception as e:
        print(e)


if __name__ == "__main__":
    main(sys.argv)
