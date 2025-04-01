# **BASF Invoice Status Checker - Excel Automation with Selenium**

This project automates the validation and classification of logistics documents (CT-e and NOTFIS) by integrating Excel data with real-time consultation on the BASF/Neogrid web portal. It checks the status of each transport document to determine whether it is eligible for invoicing, reducing manual work and ensuring compliance with BASF’s processing standards.

---

## **KEY FEATURES**

### 1. **Automated Excel Data Refresh**
- Uses `win32com.client` to open and refresh Excel connections.
- Ensures data accuracy before starting the validation process.

### 2. **Selenium-Based Web Automation**
- Logs into the BASF Neogrid portal automatically.
- Navigates through the CT-e and NOTFIS menus.
- Queries transport documents (by control number).
- Detects if a document is fully processed (`"Recebido BASF"`) or has any divergence.

### 3. **Smart Validation Rules**
- Compares the number of CT-e records on the portal with the number from Excel.
- Checks all CT-e and NOTFIS statuses per document.
- Identifies issues such as missing records or incorrect statuses.

### 4. **Filtering by Emission Date**
- Filters Excel data to process only entries issued on the previous day.
- If today is Monday, it can process all available entries by default.

### 5. **Error Classification**
- Documents that don’t match expected statuses or counts are flagged and written to a separate “Errors” report.
- Helps teams quickly identify and correct problems before invoicing.

### 6. **Clean Output Generation**
- Generates two Excel files:
  - `A Faturar`: entries eligible for invoicing.
  - `Erros`: entries requiring manual analysis.
- Also exports the invoice-ready data to a `.csv` file for integration with other systems (e.g., ERP, Power BI).

---

## **TECHNOLOGIES USED**

### 1. **Python**
- Main programming language used for automation and data handling.

### 2. **Pandas**
- Reads and processes Excel spreadsheets efficiently.

### 3. **Selenium**
- Automates web portal interaction and data extraction.

### 4. **Openpyxl**
- Writes structured Excel reports with multiple sheets.

### 5. **Win32com**
- Controls Excel via COM automation to refresh data.

### 6. **Webdriver Manager**
- Automatically installs and manages ChromeDriver for Selenium.

---

## **HOW IT WORKS**

1. Place the input file (`00. Base Basf.xlsx`) in the designated network folder.
2. The script:
   - Refreshes data connections in Excel.
   - Reads the `queryBasf` sheet.
   - Filters the data based on emission date (usually yesterday).
3. For each unique control number:
   - Logs into the BASF portal.
   - Queries CT-e and NOTFIS statuses.
   - Validates if everything is marked as `"Recebido BASF"`.
4. Generates two output files:
   - One with entries ready for invoicing.
   - Another with errors to be reviewed.
5. A `.csv` file is created for further integration.

---

## **CONCLUSION**

The **BASF Invoice Status Checker** streamlines what is typically a manual and repetitive process. By connecting Excel data with real-time web validation, this tool guarantees accuracy in financial processing and reduces the risk of errors. It’s ideal for logistics and billing teams working with high-volume transport operations and strict deadlines.

With this automation, your team can save time, eliminate manual inconsistencies, and improve the reliability of your invoicing workflow.
