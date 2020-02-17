# -*- coding: utf-8 -*-

import os
import pathlib
import sys

import csv
import glob
import openpyxl
import xlrd

""" import local liblarys """
from mylibs import filelist
from mylibs import excellibs

def main():
    argv = sys.argv
    argc = len(argv)

    if argc < 1:
        exit_msg(argv[0])

    if argc > 2:
        Path = argv[1]
    else:
        Path = "../"

    if argc > 3:
        Pattern = argv[2]
    else:
        Pattern = "*実装表*.xls*"

    if argc > 4:
        Exclude = argv[3]
    else:
        Exclude = ""

    if argc > 5:
        Depth = argv[4]
    else:
        Depth = "20"    # not use

    if Depth.isnumeric() == False:
        print("{0} is not numeric.".format(Depth))
        exit(1)

    """
    print("Path = {0}".format(Path))
    print("Pattern = {0}".format(Pattern))
    print("Exclude = {0}".format(Exclude))
    print("Depth = {0}".format(Depth))
    print("Hit Enter key...")
    input()
    """

    loopmain(Path, Pattern, Exclude, Depth)

def loopmain(Path, Pattern, Exclude, Depth):
    files = []
    for fileName in filelist.find_all_files(Path, Pattern, Exclude):
        files.append(fileName)

    # print(files)
    for file in files:
        print(file)

    for file in files:
        print(file)
        if file[-5:] == ".xlsx":
            T_WB = openpyxl.load_workbook(file)
            excellibs.readExcel_xlsx(T_WB)

        elif file[-4:] == ".xls":
            T_WB = xlrd.open_workbook(file)
            excellibs.readExcel_xls(T_WB)
        else:
            print("Error...")
            exit(-1)


def exit_msg(argv0):
    print("Usage: python {0} <Path> <Pattern> <Exclude> <Depth>".format(argv0))
    exit(0)


if __name__ == "__main__":
    main()
