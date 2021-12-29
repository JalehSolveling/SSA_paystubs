import pandas as pd
from xlsxwriter import Workbook
import win32com.client
import xlsxwriter
# Import Module
from win32com import client



def import_data(file_path):
    employees = pd.read_excel(file_path, sheet_name='Employees')
    deductions = pd.read_excel(file_path, sheet_name='Deductions')
    return (employees, deductions)


def calc_wages(employees, deductions):
    employees['Total Planned Pay'] = employees['Wages Planning'] * employees['Sessions Planned']
    employees['Total Taught Pay'] = employees['Wages Teaching'] * employees['Sessions Taught']
    employees['Total Other Pay'] = employees['Wages Other'] * employees['Sessions Other']
    employees['Gross Pay'] = employees['Total Planned Pay'] + employees['Total Taught Pay'] +employees['Total Other Pay']
    employees['SSN'] = employees['Gross Pay'] * deductions.loc[0,'Social Security Tax']
    employees['Medicare'] = employees['Gross Pay'] *deductions.loc[0,'Medicare Tax']
    employees['Total Deductions'] = employees['Gross Pay'] *deductions.loc[0,'Total Taxes Withheld']
    employees['Net Pay'] = employees['Gross Pay'] - employees['Total Deductions']

    return(employees)

def write_wages(wages, deductions, out_file_path, pay_period):
    with xlsxwriter.Workbook(out_file_path) as workbook:
        for i, row in wages.iterrows():
           #create a worksheet for each employee
            emp_name = row["Employee Name"]
            worksheet = workbook.add_worksheet(emp_name)

            #create formatting
            title_format = workbook.add_format({'font_size': 20, 'bold':True})
            header_format = workbook.add_format({'font_size': 11, 'bold':False, 'bg_color': '#EEECE1', 'align':'center'})
            header_format_left = workbook.add_format({'font_size': 11, 'bold':False, 'bg_color': '#EEECE1', 'align':'left'})
            header_percent_format = workbook.add_format({'font_size': 11, 'bold':False, 'bg_color': '#EEECE1','num_format': '0.00%', 'align':'center'})
            header_currency_format = workbook.add_format({'font_size': 11, 'bold':False, 'bg_color': '#EEECE1','num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', 'align':'center'})
            data_bold_format = workbook.add_format({'font_size': 11, 'bold':True, 'align':'center'})
            data_bold_format_left = workbook.add_format({'font_size': 11, 'bold':True, 'align':'left'})
            data_bold_currency_format = workbook.add_format({'font_size': 11, 'bold':True,'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', 'align':'center'})
            data_format = workbook.add_format({'font_size': 11, 'bold':False, 'align':'center'})
            data_format_left = workbook.add_format({'font_size': 11, 'bold':False, 'align':'left'})
            data_currency_format = workbook.add_format({'font_size': 11, 'bold':False, 'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', 'align':'center'})
            data_percent_format = workbook.add_format({'font_size': 11, 'bold':False, 'num_format': '0.00%', 'align':'center'})
            data_center_align = workbook.add_format({'align':'center', 'align':'center'})
           

            # Widen the first column to make the text clearer.
            worksheet.set_column('A:A', 40)
            worksheet.set_column('B:C', 10)
            worksheet.set_column('D:D', 15)
           
            
            #write the title           
            worksheet.write('A1', 'Swedish School of Atlanta', title_format)
            # Insert logo image.
            worksheet.insert_image('B1', 'SSA_logo.png', {'x_scale': 0.09, 'y_scale': 0.07})

            #remove tax info for volunteers
            if (row["Status"] == "V"):
                Total_Taxes_Withheld = 0
                Tax_Status = 0
                Federal_Allowance = 0
                row['Total Deductions'] = 0
                row['SSN'] = 0
                row['Medicare'] = 0
                row["Net Pay"] = row["Gross Pay"] 
            else:
                Total_Taxes_Withheld = deductions['Total Taxes Withheld']
                Tax_Status = deductions["Tax Status"]
                Federal_Allowance = deductions["Federal Allowance (From W-4)"]
                
            

        
            #write headers
            cells=['A3', 'B3','C3','D3']
            headers = ['Name','Pay Period','', 'Net Pay']
            format = [header_format_left,header_format,header_format,header_format]
            write_cells(worksheet, zip(cells,headers,format))
           

            # #write data row
            cells=['A4', 'B4','C4','D4']
            headers = [row["Employee Name"],pay_period,'',row["Net Pay"] ]
            format = [data_bold_format_left,data_bold_format, '', data_bold_currency_format]
            write_cells(worksheet, zip(cells,headers, format))
         
            if (row["Status"] == "E"):
                # #write headers
                cells=['A6', 'A7']
                headers = ['Tax Status','Tax Federal Allowance']
                format = [data_format_left,data_format_left]
                write_cells(worksheet, zip(cells,headers, format))

                # #write data
                cells=['B6', 'B7']
                headers = [Tax_Status,Federal_Allowance]
                format = [data_format,data_format]
                write_cells(worksheet, zip(cells,headers, format))
         

            # #write headers
            cells=['A9', 'B9','C9','D9']
            headers = ['Gross Pay','Lessons','Rate', 'Total']
            format = [header_format_left,header_format,header_format,header_format]
            write_cells(worksheet, zip(cells,headers, format))

            # #write header
            cells=['A10', 'A11','A12']
            headers = ['Teaching','Planning','Other']
            format = [data_format_left,data_format_left,data_format_left]
            write_cells(worksheet, zip(cells,headers, format))

            cells=['A13','B13','C13','D13']
            headers = ['Total Gross Pay','','',row['Gross Pay']]
            format = [data_format_left,'','',data_currency_format]
            write_cells(worksheet, zip(cells,headers, format))

            # #write data
            cells=['B10', 'B11','B12']
            headers = [row["Sessions Taught"], row["Sessions Planned"],row["Sessions Other"]]
            format = [data_format,data_format,data_format]
            write_cells(worksheet, zip(cells,headers, format))

            cells=['C10','C11','C12','D10','D11','D12']
            if (row["Sessions Planned"] == 0):
                wages_planning = 0
            else:
                wages_planning = row["Wages Planning"]
            headers = [row["Wages Teaching"],wages_planning, row["Wages Other"],row["Total Taught Pay"],row["Total Planned Pay"],row["Total Other Pay"]]
            format = [data_currency_format,data_currency_format,data_currency_format,data_currency_format,data_currency_format,data_currency_format]
            write_cells(worksheet, zip(cells,headers, format))

            if (row["Status"] == "E"):
                cells=['A15', 'B15','C15','D15']
                headers = ['Deductions','','Percent', 'Amount']
                format = [header_format_left,header_format,header_format,header_format]
                write_cells(worksheet, zip(cells,headers, format))

                # #write headers
                cells=['A16', 'A17','A18','A19']
                headers = ['Social Security Tax','Medicare Tax','State Tax', 'Federal Tax']
                format = [data_format_left,data_format_left,data_format_left,data_format_left]
                write_cells(worksheet, zip(cells,headers, format))

                # #write headers
                cells=['A20', 'B20','C20','D20']
                headers = ['Total Deductions','',Total_Taxes_Withheld,row['Total Deductions']]
                format = [data_format_left,data_percent_format,data_percent_format,data_currency_format]
                write_cells(worksheet, zip(cells,headers, format))

                # #write data
                cells=['C16', 'C17','C18','C19']
                headers = [deductions['Social Security Tax'],deductions['Medicare Tax'],deductions['State Tax'], deductions['Federal Income Tax']]
                format = [data_percent_format,data_percent_format,data_percent_format,data_percent_format]
                write_cells(worksheet, zip(cells,headers, format))

                # #write data
                cells=['D16', 'D17', 'D18', 'D19']
                headers = [row['SSN'],row['Medicare'],0.0,0.0]
                format = [data_currency_format,data_currency_format,data_currency_format,data_currency_format]
                write_cells(worksheet, zip(cells,headers, format))

            

def write_cells(ws,cell_data):
    for i,h in enumerate(cell_data):
        ws.write(h[0], h[1], h[2]) 


#Create a PDF for each sheet in the excel file
def create_pdf(df_wages, excel_path, path_to_pdf):
    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = False
    wb = o.Workbooks.Open(excel_path)
    
    for index, row in df_wages.iterrows():
        Employee = row['Employee Name']
        pdf_path = path_to_pdf.replace('Employees',Employee )
        wb.WorkSheets(Employee).Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
    
    wb.Close()
    o.Application.Quit()



    
