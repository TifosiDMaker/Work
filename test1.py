# -*- coding: utf-8 -*-
import openpyxl
import os
from os.path import join
for root, dirs, files in os.walk('d:/Tifosi/12月/OT5865/OT5865-泰安特种车-中译英-手册、说明书/External Review/en-US/零部件图册'):
    #print(root)
    #print(files)
    for name in files:
        print(name)
#for 1 in range(10):
