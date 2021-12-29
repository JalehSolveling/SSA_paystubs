import ssa_functions  as sf
import pandas as pd
import argparse
import pathlib

from gooey import Gooey, GooeyParser

@Gooey   
def main():
    # Create the parser
    #my_parser = argparse.ArgumentParser(description='Collect parameters from user (Pay Period, input file, outputfile,etc.')
    my_parser = GooeyParser(description='Collect parameters from user (Pay Period, input file, outputfile,etc.')
    # Add the arguments
    my_parser.add_argument('--Period',
                        metavar='Pay Period',
                        type=str,
                        help='The pay period we are calculating: ex/ Fall 2021')
    my_parser.add_argument('--Input_File',
                        metavar='Input File',
                        type=str,
                        help='the path and name of the input file used to calculate wages',
                        default='Payroll Data.xlsx',
                        widget='FileChooser')  
    my_parser.add_argument('--Output_Directory',
                        metavar='Output Directory',
                        type=str,
                        help='the path for the output file used to write paystubs',
                        default='./output/',
                        widget='DirChooser') 
    


    # Execute the parse_args() method
    args = my_parser.parse_args()

    in_file_path = args.Input_File
    out_file_path = args.Output_Directory
    pay_period = args.Period

    #get current directory for excel
    in_file_path = str(pathlib.Path.cwd() / in_file_path)
    out_file_path = str(pathlib.Path.cwd() / out_file_path)

    #put data from excel sheets into dataframes
    employees, deductions = sf.import_data(in_file_path)

    #calculate wages and store in a new dataframe
    df_wages = sf.calc_wages(employees, deductions)

    #create a new excel file for the results
    excel_file = out_file_path + '\Employees SSA Paystub - ' + pay_period + '.xlsx'
    
    sf.write_wages(df_wages, deductions,excel_file, pay_period)

    #the excel and pdf file that holds the results of calculated wages
    pdf_file = excel_file.replace(".xlsx", ".pdf")

    

    #create a pdf for each page of the excel file-each employee
    sf.create_pdf(df_wages,excel_file,pdf_file)

main()

