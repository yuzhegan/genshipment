# encoding='utf-8

# @Time: 2023-06-02
# @File: %
#!/usr/bin/env
from icecream import ic
import os
os.chdir(os.path.abspath(os.path.dirname(__file__)))
#change cwd to current file dir

# %%
#  1 先生成货件信息文件, 2 生成装箱单信息文件
# python GenShipData -i <装箱单信息> -t <装箱单信息的模板文件> -o 输入货件信息上传表格 -T <下载亚马逊当票货件的包装信息表格> -O <生成填写好的装箱单信息表格> -D <需要将文件生成到哪个文件夹下面>
!python GenShipData.py -i "装箱单.xlsx" -t "货件信息模板.xlsx" -o "货件信息上传_生成.xlsx" -D "./output/今天_718"

# %%

!python GenShipData.py -i "装箱单.xlsx" -t "货件信息模板.xlsx" -o "货件信息上传_生成.xlsx" -T "inputfiles/2023-05-27.xlsx" -O "包装信息_生成.xlsx" -D "./output/今天_test"
