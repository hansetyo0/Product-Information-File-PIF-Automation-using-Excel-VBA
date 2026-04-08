## Overview
This project is an Excel VBA-based tool designed to automate the generation of Product Information Files (PIF) using a predefined Word template.

The main goal is simple: reduce repetitive manual work and make document preparation more consistent, especially when dealing with multiple products or SKUs.

## Background
In QA/QC and regulatory environments, preparing documentation like PIFs often involves a lot of repetitive steps—copying data, formatting content, and making sure everything aligns with a standard structure.

While the task itself isn’t complex, it becomes inefficient and error-prone when done repeatedly, especially across multiple products. Small inconsistencies can also accumulate over time.

This project was built as a practical way to address that problem using a lightweight and accessible approach (Excel + VBA).

## What this tool does
Instead of creating PIF documents manually one by one, this tool lets you:
- Enter product and formulation data in Excel  
- Run a macro to process the data  
- Automatically generate a structured Word document for each SKU  

## Key Features
- Excel-based input (no additional software needed)  
- Automated Word document generation  
- Placeholder-based text replacement  
- Formula table auto-fill with sorting logic  
- Supports multiple SKUs in one run  
- Consistent output formatting across all files  

## How it works
The workflow is fairly straightforward:
1. Data is stored across structured Excel sheets (e.g. SKU master, formula, narrative)
2. VBA collects and maps the data into a dictionary-like structure
3. A Word template is opened in the background
4. Placeholders (e.g. `{{PRODUCT_NAME}}`) are replaced with actual values
5. The formulation table is dynamically filled and sorted
6. The final document is saved into the output folder

## How to use
1. Open `PIF_Template.xlsx`  
2. Fill in the required data in the relevant sheets 
3. Enable the Developer tab in Excel (if not already enabled)  
4. Open the VBA Editor (Go to Developer → Visual Basic)
5. Import the VBA module and Save 
6. Run the macro (e.g. `Generate_PIFs`)  
7. Wait for the process to complete  
8. Check the generated files in the `Output` folder  
Make sure the Word template is located in the correct directory before running the macro.

## Limitations
- Relies on consistent template structure (placeholders must match)  
- Error handling is basic and can be expanded  
- Performance may vary with very large datasets  

## Potential Improvements
If this were to be extended further, some ideas include:
- Adding a simple user interface (buttons / form controls)   
- Supporting different template versions  
- Exporting to PDF automatically   

## Disclaimer
All product names, formulations, and data used in this project are purely fictional and created for demonstration purposes only.  
The input data was generated with the assistance of AI and does not represent any real product, company, or formulation.  
No confidential or proprietary data was used in this project.  
This project is intended solely as a portfolio example.

## Notes
This tool can be adapted depending on specific documentation needs or internal workflows. The structure is intentionally kept flexible so it can be modified without too much effort.
