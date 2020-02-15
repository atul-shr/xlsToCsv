import pandas as pd
import argparse
import xlrd
import os

def create_parser():
    # print("In create arg")
    parser = argparse.ArgumentParser(description='Script to convert Excels to csv')
    parser.add_argument('--inpfile', dest='inpfile',
                        help='Input Excel file name')
    parser.add_argument('--outfile', dest='outfile',
                        help='Output CSV file name')
    parser.add_argument('--worksheet', dest='worksheet',
                        help='worksheet can be ALL, null or 1,2,3')
    return parser


def parse_arg():
    # print("In parse arg")
    parser = create_parser()
    args = parser.parse_args()
    if not args.inpfile or not args.outfile:
        print("Argument not available")
    print("Input Params are " + '-' * 60)
    print(args)
    print("Input Params are " + '-' * 60)
    return args


def createCsv(inf,outf,worksheet):
    try:
        if worksheet is None:
            final_file_name = outf
            read_file = pd.read_excel (inf)
        else :
            final_dir_name = os.path.dirname(inf)
            try:
                w_nm = int(worksheet)
                b_o_nm = os.path.basename(outf).split('.')[0]
                final_file_name = os.path.join(final_dir_name,b_o_nm+"_"+worksheet)+".csv"
            except:
                w_nm = worksheet
                final_file_name = os.path.join(final_dir_name,worksheet)+".csv"
            read_file = pd.read_excel (inf,sheet_name=w_nm)
        read_file = read_file.replace('\n','',regex=True)
        read_file.to_csv (final_file_name,index=False)
        return True
    except Exception as e:
        print(e)
        print("WorkSheet",w_nm,"Not Found")
        return False
    

args = parse_arg()

if args.worksheet is None:
    createCsv(args.inpfile,args.outfile,None)
else:
    sheetname = xlrd.open_workbook(args.inpfile)
    if args.worksheet.upper() == 'ALL':
        for sheet in sheetname.sheets():
            print(sheet.name)
            createCsv(args.inpfile,args.outfile,sheet.name)
    else :
        w_s_name = args.worksheet.split(',')
        for nm in w_s_name:
            createCsv(args.inpfile,args.outfile,nm)

