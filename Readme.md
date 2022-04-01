# Swedish School Paystub Generator

## What it does
This script will take a list of employees along with their pay rate and hours worked from an Excel file and create a paystub in PDF format for each employee.

## How it works
1. Open "sample/Payroll Data Template.xlsx".
2. Enter the following for each employee on the **Employees** sheet:
    - Employee ID (a unique number for each employee)
    - Status (E for Employee or V for Volunteer)
    - Employee Name	( the employee name as it will appear on paystub)
    - Wages Planning (wages employee received for planning a class)
    - Wages Teaching (wages employee received for teaching a class)	
    - Wages Other (wages employee received for other services)	
    - Sessions Planned (number of classes planned that semester)	
    - Sessions Taught (number of classes taught that semester)	
    - Sessions Other (number of classes for other services that semester)   
     <br>
   *Here is an example*
   ![Excel Template Sample](/docs/Excel_sample1.png)
   
   <br>
3. Modify the tax status for all the employees on the **Deductions** sheet (optional)


    *Here is the default*
    ![Excel Template Sample](/docs/Excel_sample2.png)

4. In a terminal, run:  `python ssa.py`. This will generate a GUI window where you need to enter the following parameters:
    1. Pay Period that will be printed on paystub
    2. Location of Excel file with Employee data
    3. The location of the folder to write the paystubs

    <br>

     ![Window Pop-up](/docs/Window_sample1.png)

5. Click **Start**.
   The scipt will generate a PDF paystub document for each employee in your list and put the file in the output folder you selected.

   When the script is finished, a new window will be generated.

    ![Window Pop-up](/docs/Window_sample2.png)

    Click **OK** and **Close**

Here are examples of the different paystubs generated, one for employees and one for volunteers.

![Window Pop-up](/docs/Paystub_sample1.png)

![Window Pop-up](/docs/Paystub_sample2.png)
   
***Troubleshooting Tips***
1. Make sure to enter a value for all fields in the Excel file, even if it is zero.
2. Check that you have saved your changes to the Excel file before running the script.
3. Check that you are in the correct directory to run the script.

