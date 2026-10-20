Keep Following my youtube channel for practical import process: https://youtu.be/Iv-prbCsev8
# Tally.ERP 9 Employee Master XML Toolkit

An unofficial, open-source toolkit for transferring Employee Master records between **Tally.ERP 9 XML** and **Microsoft Excel**.

The project is intended to help users export Employee Masters from Tally.ERP 9 to Excel for review and structured data maintenance, and convert completed Excel records back into Tally-compatible XML for import through the standard **Import Data → Masters** facility.

> [!IMPORTANT]
> This project is independent and unofficial. It is not developed, endorsed, certified, maintained, or supported by Tally Solutions Pvt. Ltd. or any government organization.
>
> Always test generated XML in a copied or backup company before importing it into live payroll data.

---

## Features

### Excel to Tally XML

- Converts a structured Excel Employee Master workbook into Tally-compatible XML.
- Creates new employee Cost Centre masters.
- Supports updating existing Employee Masters when Tally's **Modify with New Data** import behavior is deliberately selected.
- Validates mandatory employee information before creating XML.
- Detects duplicate Employee Names, Employee Numbers, and CNIC values in the workbook.
- Omits blank optional fields rather than writing unnecessary empty XML tags.
- Includes a Windows CMD launcher for users who do not want to run Python commands manually.

### Tally XML to Excel

The planned reverse-conversion utility is intended to:

- Read a Tally-exported `Master.xml` file.
- Identify Employee Cost Centre masters.
- Exclude Employee Groups and unrelated Cost Centres.
- Extract Employee Master fields into one Excel row per employee.
- Generate separate validation and summary worksheets.
- Support review, correction, and controlled re-import of employee records.

> The XML-to-Excel utility should be treated as experimental until it has been separately tested with fictional Employee Masters from the relevant Tally release.

---

## Employment Status and TDL Dependency

**Employment Status is not a standard field created independently by this XML toolkit.**

The Employment Status value is pulled from, and written to, a custom User Defined Field declared in the separately loaded TDL customization:

```tdl
[System : UDF]

    NCRCEmploymentStatus : String : 19001
```

The Excel column:

```text
EmploymentStatus
```

is mapped to the following Tally UDF XML structure:

```xml
<UDF:NCRCEMPLOYMENTSTATUS.LIST
    DESC="`NCRCEmploymentStatus`"
    ISLIST="YES"
    TYPE="String">

    <UDF:NCRCEMPLOYMENTSTATUS DESC="`NCRCEmploymentStatus`">
        Regular
    </UDF:NCRCEMPLOYMENTSTATUS>

</UDF:NCRCEMPLOYMENTSTATUS.LIST>
```

### Important requirements for Employment Status

- The TDL file defining `NCRCEmploymentStatus` must be loaded and active in Tally before importing employee XML.
- The UDF index must remain `19001` unless the TDL and XML converter are deliberately updated together.
- Another loaded TDL must not use UDF index `19001` for a different purpose.
- If the TDL is not loaded, native Employee Master fields may import while Employment Status may be ignored or unavailable.
- Existing employees with no stored Employment Status may display `Regular` only because the payslip TDL applies `Regular` as a fallback value.
- Employment Status should be verified through individual Employee Alteration after each test import.

Supported Employment Status values in the supplied template are:

- Regular
- Contract
- Deputation
- Daily Wages
- Intern
- Project Employee
- Retired
- Terminated

The UDF name is retained for compatibility with the original implementation. Users creating a new independent implementation may rename the UDF, but the TDL definition, XML converter, and existing stored records must be handled consistently.

---

## Supported Employee Fields

The full Excel template includes the following fields.

### Basic employment information

- Employee Name
- Employee Category
- Employee Group
- Employee Number
- Date of Joining
- Employment Status
- Designation
- Function
- Location

### Personal information

- Gender
- Date of Birth
- Blood Group
- Father / Mother Name
- Spouse Name
- Address Line 1
- Address Line 2
- Contact Number
- Email

### Banking information

- Bank Name
- Bank Branch
- Bank Account Number
- MICR Code

### CNIC, passport, visa, and contract information

- CNIC
- Country of Issue
- Passport Expiry Date
- Visa Number
- Visa Expiry Date
- Work Permit Number
- Contract Start Date
- Contract Expiry Date

### Separation information

- Date of Leaving
- Reason for Leaving

### Statutory information

- PAN Number
- PF Account Number
- FPF Account Number
- PF Joining Date
- PF Relieving Date
- ESI Number
- ESI Dispensary Name

---

## Field Mapping

| Excel Column | Tally XML Method or Storage |
|---|---|
| EmployeeName | `COSTCENTRE NAME` and `LANGUAGENAME.LIST` |
| EmployeeCategory | `CATEGORY` |
| EmployeeGroup | `PARENT` |
| EmployeeNumber | `MAILINGNAME.LIST` |
| DateOfJoining | `DATEOFJOIN` |
| EmploymentStatus | `NCRCEMPLOYMENTSTATUS` UDF from the loaded TDL |
| Designation | `DESIGNATION` |
| Function | `FUNCTION` |
| Location | `LOCATION` |
| Gender | `GENDER` |
| DateOfBirth | `DATEOFBIRTH` |
| BloodGroup | `BLOODGROUP` |
| FatherMotherName | `FATHERNAME` |
| SpouseName | `SPOUSENAME` |
| AddressLine1 / AddressLine2 | `ADDRESS.LIST` |
| ContactNumber | `CONTACTNUMBERS` |
| Email | `EMAILID` |
| BankName | `BANKDETAILS` |
| BankBranch | `BANKBRANCH` |
| BankAccountNumber | `BANKACCOUNTNUMBER` |
| MICRCode | `MICRCODE` |
| CNIC | `PASSPORTDETAILS` |
| CountryOfIssue | `COUNTRYOFISSUE` |
| PassportExpiryDate | `PASSPORTEXPIRYDATE` |
| VisaNumber | `VISANUMBER` |
| VisaExpiryDate | `VISAEXPIRYDATE` |
| WorkPermitNumber | `WORKPERMITNUMBER` |
| ContractStartDate | `CONTRACTSTARTDATE` |
| ContractExpiryDate | `CONTRACTEXPIRYDATE` |
| DateOfLeaving | `DEACTIVATIONDATE` |
| ReasonForLeaving | `REASONSFORLEAVING` |
| PANNumber | `PANNUMBER` |
| PFAccountNumber | `PFACCOUNTNUMBER` |
| FPFAccountNumber | `FPFACCOUNTNUMBER` |
| PFJoiningDate | `PFJOININGDATE` |
| PFRelievingDate | `PFRELIEVINGDATE` |
| ESINumber | `ESINUMBER` |
| ESIDispensaryName | `ESIDISPENSARYNAME` |

> CNIC is intentionally mapped to the native `PassportDetails` method because the associated payslip customization retrieves CNIC from that storage.

---

## Repository Structure

```text
tally-erp9-employee-master-xml-toolkit/
│
├── README.md
├── LICENSE
├── DISCLAIMER.md
├── CONTRIBUTING.md
├── SECURITY.md
├── CHANGELOG.md
├── .gitignore
│
├── excel-to-xml/
│   ├── Employee_Master_Full_Template.xlsx
│   ├── employee_full_xlsx_to_xml.py
│   ├── Convert_Employee_Excel_To_XML.cmd
│   └── README.md
│
├── xml-to-excel/
│   ├── tally_employee_xml_to_xlsx.py
│   ├── Convert_Tally_XML_To_Excel.cmd
│   └── README.md
│
├── samples/
│   ├── Sample_Employee_Master.xml
│   └── Sample_Employee_Master.xlsx
│
└── docs/
    ├── installation.md
    ├── field-mapping.md
    ├── troubleshooting.md
    └── privacy.md
```

All files in `samples/` must contain fictional information only.

---

## Requirements

### For Excel-to-XML conversion

- Windows 10 or Windows 11
- Python 3
- Python package `openpyxl`
- Microsoft Excel or another compatible spreadsheet application
- Tally.ERP 9 with Payroll and Cost Centres enabled
- Existing Employee Categories and Employee Groups
- The Employment Status TDL loaded if the Status UDF is used

Install the required Python package with:

```bash
python -m pip install openpyxl
```

The supplied CMD utility can also attempt to install `openpyxl` when it is missing.

---

## Excel-to-XML Usage

### 1. Download or clone the repository

```bash
git clone https://github.com/YOUR-USERNAME/tally-erp9-employee-master-xml-toolkit.git
```

Replace `YOUR-USERNAME` with the actual GitHub username or organization name.

### 2. Open the Excel template

Open:

```text
excel-to-xml/Employee_Master_Full_Template.xlsx
```

Enter one employee per row in the `Employees` worksheet.

Do not rename or remove the column headings unless the converter is updated accordingly.

### 3. Follow the data rules

- Employee Name is mandatory.
- Employee Category is mandatory.
- Employee Group is mandatory.
- Employee Category and Employee Group must already exist in Tally.
- Use `YYYYMMDD` for all dates.
- Example: `20260401` means 1 April 2026.
- Leave optional fields blank when not applicable.
- Employee Name must be unique in the workbook.
- Employee Number must be unique when provided.
- CNIC must be unique when provided.
- Avoid formulas in input cells. Enter final values.

### 4. Generate the XML using CMD

Save and close Excel, then double-click:

```text
Convert_Employee_Excel_To_XML.cmd
```

The utility creates:

```text
Employee_Master_Full_Import.xml
```

### 5. Generate the XML using Python directly

```bash
python employee_full_xlsx_to_xml.py \
  Employee_Master_Full_Template.xlsx \
  Employee_Master_Full_Import.xml
```

To specify the exact Tally company name:

```bash
python employee_full_xlsx_to_xml.py \
  Employee_Master_Full_Template.xlsx \
  Employee_Master_Full_Import.xml \
  --company "Your Exact Company Name"
```

### 6. Import into Tally

Use a copied or test company first.

```text
Gateway of Tally
→ Import Data
→ Masters
```

Select:

```text
Employee_Master_Full_Import.xml
```

For new employees, first test:

```text
Ignore Duplicates
```

Use:

```text
Modify with New Data
```

only when intentionally updating existing Employee Masters and only after taking a complete backup.

### 7. Verify the import

After importing:

1. Review `Tally.imp`.
2. Open each test employee through individual Employee Display or Alteration.
3. Verify Employee Group and Category.
4. Verify Employee Number.
5. Verify joining and leaving dates.
6. Verify CNIC.
7. Verify Employment Status.
8. Verify bank and statutory information.
9. Print a test payslip where relevant.

---

## Exporting Employee Masters from Tally

To prepare XML for reverse conversion:

```text
Gateway of Tally
→ Display
→ List of Accounts
→ Export
```

Recommended export settings:

```text
Format                     : XML (Data Interchange)
Type of Masters            : All Masters or Cost Centres
Include dependent masters  : Yes
Output File Name           : Master.xml
```

The reverse converter should identify employees using values equivalent to:

```xml
<FORPAYROLL>Yes</FORPAYROLL>
<ISEMPLOYEEGROUP>No</ISEMPLOYEEGROUP>
```

Employee Groups should normally have:

```xml
<ISEMPLOYEEGROUP>Yes</ISEMPLOYEEGROUP>
```

---

## Compatibility

The toolkit was developed from Employee Master XML produced by a legacy:

```text
Tally.ERP 9
Series A
Release 1.1
```

Other releases may use different field names, XML ordering, native definitions, or import behavior.

Compatibility is not guaranteed for:

- Other Tally.ERP 9 releases
- TallyPrime
- Different country editions
- Different payroll configurations
- Custom TCP files
- Other TDL files using the same UDF index
- Employee Masters using additional custom UDFs

Always test with fictional records in a copied company.

---

## Import Behavior and Record Matching

The generated XML identifies an Employee Master primarily by exact Employee Name.

This means:

- A new name may create a new Employee Master.
- An exact existing name may be altered when **Modify with New Data** is selected.
- Changing an employee's name in Excel may create another master instead of renaming the original master.
- Employee Number and CNIC are validated in Excel but are not the primary Tally master identifier in this workflow.

Never use **Modify with New Data** without a complete backup.

---

## Blank Values

The converter omits blank optional fields from generated XML.

This reduces the risk that an empty workbook cell will overwrite an existing value. However, behavior can vary by Tally release and import mode.

When exporting XML to Excel, a blank Excel cell may mean:

- no value is stored in Tally; or
- Tally omitted the empty method from the exported XML.

Users must verify critical fields before re-importing records.

---

## Known Limitations

- Employee Categories and Employee Groups are not automatically created by the employee converter.
- Existing records are matched primarily by exact Employee Name.
- The utility does not import payroll salary structures or monthly payroll vouchers.
- The utility does not assign pay heads or rates to employees.
- The utility does not calculate salary, tax, deductions, or arrears.
- The utility does not replace Tally backups.
- The Employment Status field depends on the separately loaded TDL UDF.
- Custom UDFs other than `NCRCEmploymentStatus` require additional mapping.
- The XML schema may require adjustment for other Tally releases.
- The reverse XML-to-Excel utility must be tested independently before production reliance.

---

## Data Protection and Privacy

Employee Master exports can contain highly sensitive personal and financial information.

Never upload real files containing:

- Employee names
- CNIC or identity numbers
- Bank account numbers
- Dates of birth
- Residential addresses
- Telephone numbers
- Email addresses
- Salary information
- Payroll registers
- Tally serial numbers
- Government identifiers

Do not upload the following real files to GitHub:

```text
Master.xml
Tally.imp
Employee_Master_Export.xlsx
Employee_Master_Full_Import.xml
```

Use fictional or fully anonymized data in public samples, screenshots, issues, and pull requests.

---

## Recommended `.gitignore`

```gitignore
# Confidential Tally exports
Master.xml
master.xml
Tally.imp
tally.imp
*.900
*.tsf

# Generated employee output
Employee_Master_Export.xlsx
Employee_Master_Full_Import.xml
Employee_Master_Import.xml

# Private data folders
private/
production/
backups/
employee-data/
real-data/

# Python
__pycache__/
*.pyc
.venv/
venv/

# Excel temporary files
~$*.xlsx
*.tmp

# Local configuration
.env
config.local.json

# Logs
*.log
```

---

## Testing Checklist

Before a production import, confirm:

- [ ] A complete Tally company backup exists
- [ ] Testing is being performed in a copied company
- [ ] Payroll is enabled
- [ ] Cost Centres are enabled
- [ ] Employee Categories exist
- [ ] Employee Groups exist
- [ ] The Employment Status TDL is active
- [ ] UDF index `19001` is not conflicting
- [ ] Employee Names are unique
- [ ] Employee Numbers are unique
- [ ] CNIC values are unique
- [ ] Dates use `YYYYMMDD`
- [ ] The XML contains the expected company name
- [ ] `Tally.imp` has been reviewed
- [ ] Imported Employee Masters have been opened individually
- [ ] A test payslip has been printed
- [ ] Employee counts before and after import have been compared

---

## Contributing

Contributions are welcome.

Before submitting a pull request:

1. Use fictional employee records only.
2. State the exact Tally release used for testing.
3. Describe the XML tags or fields affected.
4. Keep changes small and reversible.
5. Include test cases for blank and populated fields.
6. Do not include proprietary native Tally source files.
7. Do not include real Tally company data.
8. Update documentation when field mappings change.

Suggested contribution areas include:

- XML-to-Excel conversion
- Better validation reports
- Additional Employee Master fields
- Configurable UDF names and indexes
- Compatibility profiles for other Tally releases
- Automated fictional test fixtures
- Improved Windows packaging

---

## Security

Do not report privacy or security issues using public employee data.

If an issue requires a sample XML file, replace all personal, payroll, banking, and company information with fictional values before sharing it.

If confidential information is accidentally committed, remove it from the repository and Git history immediately, rotate any exposed credentials, and treat copied employee data as compromised.

---

## Disclaimer

This software is provided **as is**, without warranty of any kind.

The author and contributors are not responsible for:

- Data loss
- Incorrect Employee Masters
- Duplicate masters
- Overwritten employee information
- Incorrect payroll
- Tax errors
- Financial loss
- Recovery proceedings
- Audit objections
- Employee claims
- Business interruption
- Incompatibility with a Tally release
- Conflicts with another TDL or TCP file

Users are responsible for maintaining backups, testing all changes, reviewing generated XML, reviewing `Tally.imp`, and verifying imported records.

---

## Trademark Notice

This repository is an independent and unofficial community project.

Tally, Tally.ERP 9, TallyPrime, TDL, and related product names belong to their respective owner.

Use of these names is solely to identify compatibility and the environment for which the toolkit was developed. No affiliation, endorsement, certification, or official support is claimed.

---

## License

This project is intended to be released under the **MIT License**.

Add a root-level `LICENSE` file containing the MIT License text before publishing the repository.

Suggested copyright line:

```text
Copyright (c) 2026 Nasr Shahbaz
```

Replace the public author name with a GitHub username if preferred.

---

## Project Origin

This project originated from a government-style payroll and Employee Master customization exercise and was generalized for public, educational, migration, and interoperability use.

The original implementation used a custom Employment Status field defined through a TDL UDF. That dependency is deliberately documented so users understand that the Status value is not a standalone native Employee Master field supplied by this converter.

---

## Release Status

Recommended first public release:

```text
v1.0.0-beta
```

The project should remain marked as beta until both Excel-to-XML and XML-to-Excel workflows have been tested with multiple fictional Employee Masters across the intended Tally releases.
