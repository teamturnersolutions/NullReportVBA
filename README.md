# Null Summary Report Automation
This project automates the processing and reporting of inventory data using VBA macros in Microsoft Excel. The workflow is designed to clean, filter, and transfer data between workbooks, providing an efficient solution for managing inventory discrepancies and ensuring accurate reporting.

## Features
### Data Cleanup:
* Deletes rows based on specific criteria (e.g., warehouse codes or date ranges).
* Removes unnecessary columns for a streamlined dataset.

### Data Transformation:
* Dynamically adjusts worksheet names and structures.
* Filters data based on inventory aging policies (e.g., records younger than 3 days since the start of the shift).

### Data Transfer:
* Transfers filtered data between source and destination workbooks.
* Avoids overwriting existing headers in the destination workbook.

### Workflow Automation:
* Combines all macros into a single execution workflow, ensuring sequential processing.

### Prerequisites
* Microsoft Excel (with support for VBA macros enabled).
* Familiarity with enabling macros in Excel (Developer tab must be enabled).
* Source workbook naming convention starts with "Null Location LPNs" or similar variations.

## Setup
### Enable Macros:
* Open Excel.
* Navigate to File > Options > Trust Center > Trust Center Settings > Macro Settings.
* Enable macros and allow VBA access.

### Load VBA Macros:
* Open the Visual Basic for Applications editor (Alt + F11).
* Import the provided VBA scripts into the desired workbook.

### Save as Macro-Enabled Workbook:
* Save the workbook as a .xlsm file to preserve macros.
* Execution Instructions

## Macro Overview
### The project includes the following macros:

### NullSummaryReport_Stage1:
* Prepares the dataset by performing initial cleanup and transformations.

### NullSummaryReport_Stage2:
* Filters and deletes data based on warehouse codes and inventory aging policies.
* Removes specific columns and renames the worksheet.

### NullSummaryReport_Stage3:
* Transfers data from the source workbook (My Null Location LPNs) to the destination workbook (Sandbox), starting from row 4.

### RunAllStages:
* Executes all stages sequentially, ensuring the workflow completes step by step.

# Troubleshooting
### Workbook Not Found:
* Ensure the source workbook is open before running the macros.
* Verify the naming convention matches the code ("Null Location LPNs").

### Incorrect Data Deletion:
* Double-check the date and time filter logic in NullSummaryReport_Stage2.
* Modify the criteria as per inventory policies.

### Macro Errors:
* Ensure macros are enabled and saved in a macro-enabled workbook (.xlsm).
* Verify worksheet names match those referenced in the macros.
