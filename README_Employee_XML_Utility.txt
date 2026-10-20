EMPLOYEE MASTER EXCEL TO XML UTILITY

RUN Convert_Employee_Excel_To_XML AS ADMINISTRATOR IT WILL CONVERT ALL THE EXCEL FILE RECORDS INTO .XML FILE

FILES TO KEEP IN THE SAME FOLDER
1. Employee_Master_Full_Template.xlsx
2. employee_full_xlsx_to_xml.py
3. Convert_Employee_Excel_To_XML.cmd

HOW TO USE
1. Open Employee_Master_Full_Template.xlsx.
2. Replace the sample row with employee data. MAKE SURE NO DUPLICATION OF EMPLOYEE DATA
3. Save and close Excel.
4. Double-click Convert_Employee_Excel_To_XML.cmd.
5. The utility creates Employee_Master_Full_Import.xml.
6. In a copied Tally company, go to Gateway of Tally > Import Data > Masters.
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
