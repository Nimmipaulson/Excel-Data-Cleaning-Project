# Excel-Data-Cleaning-Project
Project Overview

This project focuses on cleaning, organizing, and formatting a product dataset using Microsoft Excel. The dataset contains product information such as Product ID, Product Name, Brand Name, Price, Quantity, and Category.

The main objective of this project is to improve data quality, maintain consistency, identify missing values and duplicate records, and make the dataset easier to analyze using Excel functions and formatting tools.

🎯 Objectives
Identify and handle missing values in the dataset.
Detect and remove duplicate records.
Split the Product ID into Manufacturing Date and Country Code.
Combine Brand Name and Product Name into a single column.
Format the Price column as currency.
Extract and organize the date information from the Product ID.
Apply conditional formatting to highlight price variations and specific product categories.
Improve the overall consistency and readability of the dataset.
📂 Repository Contents

This repository contains the following files:

File Name	Description
Product_Data_Cleaning.xlsx	Excel workbook containing the product dataset and data-cleaning operations.
Excel_Functions_Report.docx	Word document explaining the steps, formulas, and methods used to clean and format the dataset.
README.md	Project overview, objectives, and description of the tasks performed.
📊 Dataset Description

The dataset contains the following columns:

Column Name	Description
Product ID	Product identifier containing date and country information.
Product Name	Name of the product.
Brand Name	Brand associated with the product.
Price ($)	Product price in US Dollars.
Quantity	Available quantity of the product.
Category	Product category, such as Electronics, Fashion, and Kitchen.
Manufacturing Date	Date information extracted from the Product ID in DD-MMM format.
Country Code	Country code extracted from the Product ID.
Product Brand	Combined brand name and product name.
🛠️ Tools and Functions Used

Tool: Microsoft Excel

1. Missing Value Identification and Handling
Used the Filter option to identify blank cells in the Price and Category columns.
Used the COUNTBLANK function to count missing values.
Used comparable product prices and the overall median to estimate missing prices.
Filled missing category values based on the product type.
2. Removing Duplicate Records
Used Data → Remove Duplicates to identify and remove duplicate rows.
Selected all columns to identify records that were identical across the dataset.
3. Extracting Manufacturing Date and Country Code

Used the LEFT and RIGHT functions to split the Product ID into separate columns.

Formulas:

=LEFT(A2,6)
=RIGHT(A2,2)
LEFT extracts the date information from the beginning of the Product ID.
RIGHT extracts the country code from the end of the Product ID.
4. Combining Brand Name and Product Name

Used the CONCAT function to merge the Brand Name and Product Name columns into a single column called Product Brand.

Formula:

=CONCAT(C2," ",B2)

This formula combines the brand name and product name with a space between them.

5. Formatting the Price Column
Applied US Dollar currency formatting to the Price column.
Set the decimal places to two for consistent presentation.
6. Formatting the Manufacturing Date
Extracted the date information from the Product ID using the LEFT function.
Retained the date in DD-MMM format because the original Product ID does not contain a year.
Did not apply the DD-MM-YYYY format because the year was not available in the dataset.
7. Applying Conditional Formatting
Applied Data Bars or Color Scales to the Price column to visually compare product prices.
Created a custom conditional formatting rule to highlight cells containing the category Electronics.

Formula-based rule:

=$F2="Electronics"

This formula highlights cells where the Category column contains Electronics.

📈 Key Outcomes
Identified missing values in the Price and Category columns.
Removed duplicate records from the dataset.
Extracted date information and country codes from Product ID.
Combined Brand Name and Product Name into a single column.
Standardized the currency format of the Price column.
Organized date information in the DD-MMM format.
Applied conditional formatting to improve data visualization.
Improved the consistency and readability of the dataset.
📄 Project Documentation

The accompanying Word document, Excel_Functions_Report.docx, provides a detailed explanation of the steps followed in Excel, including objectives, formulas, procedures, and results for each task.

🎓 Learning Outcomes

Through this project, I gained practical experience in:

Excel data cleaning and data organization.
Identifying and handling missing values.
Removing duplicate records.
Using text functions such as LEFT, RIGHT, and CONCAT.
Applying currency and date formatting.
Using conditional formatting to highlight important information.
Documenting data-cleaning procedures.
👤 Author

Nimmi Paulson

This project was completed as part of my learning journey in Data Analytics.
