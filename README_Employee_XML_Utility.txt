EMPLOYEE MASTER EXCEL TO XML UTILITY

FILES TO KEEP IN THE SAME FOLDER
1. Employee_Master_Full_Template.xlsx
2. employee_full_xlsx_to_xml.py [MAIN ENGINE DRIVING OUR BUS]
3. Convert_Employee_Excel_To_XML.cmd

HOW TO USE
1. Open Employee_Master_Full_Template.xlsx
2. Replace the sample row with employee data. MAKE SURE NO DUPLICATION OF EMPLOYEE DATA. KEEP EmployeeCategory TO DEFAULT OR CREATE CATEGORIES AND LIST IN EXCEL. IF THERE ARE EMPLOYEE GROUPS LIKE "PAYSCALE-1, PAYSCALE-2 ETC" LIST THEM UNDER EmployeeGroup AND THERE IS NO GROUP, KEEP THE EMPLOYEE GROUP EMPTY.
3. Save and close Excel.
4. Double-click Convert_Employee_Excel_To_XML.cmd.
5. The utility creates Employee_Master_Full_Import.xml. Right click the file and copy the path. make sure to remove inverted comas by pasting in tally ctrl +alt + v
6. In a copied Tally company, go to Gateway of Tally > Import Data > Masters. enter enter enter
7. Import the XML and review Tally.imp.

IMPORTANT
- Python 3 is required.
- The CMD utility installs openpyxl if it is missing.
- Employee Category and Employee Group must already exist in Tally.
- Keep the payslip TDL active for NCRCEmploymentStatus UDF index 19001.
- Dates must be YYYYMMDD.
- Existing employees are matched by exact Employee Name.
- Use Modify with New Data only after a complete backup.

DEVELOPED BY: MUHAMMAD NASR
