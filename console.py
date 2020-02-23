#!/usr/bin/python3
# -*- coding:utf-8 -*-
from writeTable import *
from readTable import *
import sys
import getopt
import os

def printUsage():
    helppage='''
    usage: console.py -i <input file> -o <output file>
           console.py --in=<input file> --out=<output file>
    '''
    print(helppage)


def mainconsole():
    inputarg = ""
    outputarg = ""
    try:
        opts, args = getopt.getopt(sys.argv[1:], "hi:o:", ["in=", "out="])
    except getopt.GetoptError:
        printUsage()
        sys.exit(-1)
    for opt, arg in opts:
        if opt == '-h':
            printUsage()
        elif opt in ("-i", "--in"):
            inputarg = arg
        elif opt in ("-o", "--out"):
            outputarg = arg

    if inputarg == '' or outputarg == '':
        print('Missing parameters')
        sys.exit(-1)

    path, suffix = os.path.splitext(inputarg)
    filesuffix = suffix.strip().lower()

    if filesuffix.lower() == '.csv':
        try:
            content = readCSV(inputarg)
        except Exception as e:
            print('Error occurred when opening csv')
            sys.exit(-1)
        else:
            print('read csv successfully')
    elif filesuffix.lower() == '.xls':
        try:
            content = read03xls(inputarg)
        except Exception as e:
            print('Error occurred when opening xls')
            sys.exit(-1)
        else:
            print('read xls successfully')
    elif filesuffix.lower() == '.xlsx':
        try:
            content = read07xlsx(inputarg)
        except Exception as e:
            print('Error occurred when opening xlsx')
            sys.exit(-1)
        else:
            print('read xlsx successfully')
    else:
        print('Unsopported input file suffix!\nOnly ".xls", ".xlsx", ".csv" are permitted')

    opath, osuffix = os.path.splitext(outputarg)
    ofilesuffix = osuffix.strip().lower()

    if ofilesuffix.lower() == '.csv':
        try:
            writeCSV(outputarg[0:-4] + '.csv', content)
        except Exception as e:
            print('Error occurred when saving csv')
            sys.exit(-1)
        else:
            print('write csv successfully')
    elif ofilesuffix.lower() == '.xls':
        try:
            write03xls(outputarg[0:-4] + '.xls', content)
        except Exception as e:
            print('Error occurred when saving xls')
            sys.exit(-1)
        else:
            print('write xls successfully')
    elif ofilesuffix.lower() == '.xlsx':
        try:
            write07xlsx(outputarg[0:-5] + '.xlsx', content)
        except Exception as e:
            print('Error occurred when saving xlsx')
            sys.exit(-1)
        else:
            print('write xlsx successfully')
    elif ofilesuffix.lower() == '.tex':
        try:
            writeLaTeX(outputarg[0:-4] + '.tex', content)
        except Exception as e:
            print('Error occurred when saving tex')
            sys.exit(-1)
        else:
            print('write tex successfully')
    else:
        print('Unsopported output file suffix!\nOnly ".xls", ".xlsx", ".csv", ".tex" are permitted')



if __name__ == "__main__":
    mainconsole()