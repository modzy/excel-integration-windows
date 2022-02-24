# Modzy Integration with Excel for Windows

<img src="https://www.modzy.com/wp-content/uploads/2020/06/MODZY-RGB-POS.png" alt="modzy logo" width="250"/>

<div align="center">

**This repository contains the VBA scripts required to build a Modzy integration with Excel on a Windows machine.**

![GitHub contributors](https://img.shields.io/github/contributors/modzy/excel-integration-windows)
![GitHub last commit](https://img.shields.io/github/last-commit/modzy/excel-integration-windows)
![GitHub Release Date](https://img.shields.io/github/issues-raw/modzy/excel-integration-windows)

[Excel for Mac with VBA](<https://github.com/modzy/integration-excel-mac>)
</div>

## Overview

This repository contains resources for building a Modzy integration into Excel with VBA
  
## Usage Instructions
  
  1. Clone repository: `git clone https://github.com/modzy/excel-integration-windows.git`
  2. Open the `./data/modzy-credit-default-risk-data.xlsx` Excel workbook
  3. Add a sheet and name it "ML Predictions"
  4. Enable your Developer Tab: Open up Excel --> go to Preferences --> go to Ribbon & Toolbar --> click on the "Developer" tab to enable it in the main ribbon
  5. Open up the VBA IDE: Click on the "Developer" tab -> Click on the "Visual Basic" icon
  6. Add Modzy API Class: Right click your VBA project, select "Import File," and add the `scripts/API_Client.cls` class module
  7. Add modules: Right click your VBA project, select "Import File," and add both `scripts/JsonConverter.bas` and `scripts/RunModzyModel.bas` modules.
  8. Open "Module1" under the "Modules" folder, and navigate to the `runHomeCreditModel()` method at the bottom of the script. Here replace "<add-modzy-URL>" and "<add-modzy-api-key>" with your valid Modzy credentials.
  9. Turn "Microsoft Scripting Runtime" on: Go to Tools --> References --> Scroll down to "Microsoft Scripting Runtime" and check the box --> OK
  10. Save workbook as Macro-Enabled Workbook
  11. In the "ML Predictions" sheet within your workbook, add a form control button: Developer Tab --> Insert --> Button (Form Controls), and select `runHomeCreditModel` as the macro enabled by this button.
  12. Replace the text in the button with whatever phrase you prefer (e.g., "Run Model"). 
  13. In Cell A6, type "Row Label", and in Cell B6, type "Risk Score"
  14. Press your button to kick off an inference job in Modzy.
 
This integration takes the data from the "Preprocessed Data" sheet, formats it into the required CSV format, submits it to a model within Modzy, and returns the predictions directly to the "ML Predictions" sheet, ultimately demonstrating how to integrate AI capabilities into a tool business analysts use daily.  

## Table of contents

- [Data](<https://github.com/modzy/excel-integration-windows/tree/master/data>): contains spreadsheet with raw and preprocessed sample data for this integration
- [Scripts](<https://github.com/modzy/excel-integration-windows/tree/master/scripts>): contains API class and modules required to build this integration

## Contributing

We are happy to receive contributions from all of our users. Check out our [contributing file](https://github.com/modzy/excel-integration-windows/blob/master/CONTRIBUTING.adoc) to learn more.

## Code of conduct

[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-v2.0%20adopted-ff69b4.svg)](https://github.com/modzy/excel-integration-windows/blob/master/CODE_OF_CONDUCT.md)
