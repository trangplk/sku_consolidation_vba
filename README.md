# skuconsolidation_vba
# Import SKU Macro - GitHub README

## Overview

The `Import_SKU` macro is a specialized Excel VBA tool designed for aggregating SKU (Stock Keeping Unit) promotion data from multiple Excel workbooks into a single worksheet and performing data cleaning. 

- **Pre-Import Data Clearance Confirmation:** Prompts users to confirm before deleting existing data in the SKU worksheet.
- **Support for Multiple File Imports:** Enables the selection and importation of data from several `.xlsx` files in one operation.
- **Automated Data Formatting:** Ensures uniform copying and formatting of data from the source to the target worksheet.
- **Custom Data Validation and Cleaning:** Incorporates logic to filter out specific entries and cleanse the dataset.
- **Completion Alert:** Notifies users upon successful completion and reminds them to check for any duplicate SKUs.

## Requirements

- Microsoft Excel with macro support.
- Basic understanding of Excel VBA for executing and potentially modifying the macro.
- As GetFileOpenName does not work on MacOS Excel, there are two xlsm files accordingly for the MacOS and Windows operating systesm


## Usage
1. Open your target Excel workbook.
2. Activate the macro by pressing `Alt + F8` in Excel, selecting `Import_SKU`, and hitting 'Run'.
2. Follow the on-screen prompts to select the `.xlsx` files you want to import. There are two sample files `.xlsx` included here.
3. The macro will process the selected files and import the data into the "All SKUs" worksheet.

## Customization

Users can tailor the macro to specific needs by modifying the VBA code. Typical customizations might include changing the target cells, modifying data filtering criteria, or altering the data format settings.

---
*This README is specifically for the `Import_SKU` Excel VBA macro and is for informational purposes only.*
