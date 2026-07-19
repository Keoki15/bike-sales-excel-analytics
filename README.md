# 🚴 Bike Sales Excel Analytics Portfolio

## Project Overview

This repository contains a collection of Microsoft Excel projects completed using a Bike Sales dataset. Each workbook represents a different stage of the data analytics process, including data preparation, sorting and filtering, outlier detection, PivotTable analysis, and data visualization.

The purpose of this portfolio is to demonstrate how Excel can be used to transform raw business data into organized, reliable, and understandable information that supports reporting and decision-making.

---

## Project Objectives

The main goals of this project were to:

- Prepare raw data for analysis
- Identify and correct data-quality problems
- Organize records using sorting and filtering
- Detect unusual or extreme values
- Summarize business information with PivotTables
- Communicate findings through charts and visualizations
- Develop practical Excel skills used by data analysts

---

## Dataset Information

The Bike Sales dataset contains sales transaction and customer information, including fields such as:

- Sales order number
- Transaction date
- Customer age
- Age group
- Customer gender
- Country
- State
- Product category
- Product subcategory
- Product name
- Order quantity
- Unit cost
- Unit price
- Revenue
- Cost
- Profit

---

## Repository Structure

```text
bike-sales-excel-analytics/
│
├── 01-Data-Preparation/
│   ├── README.md
│   └── Bike_Sales_Data_Preparation.xlsx
│
├── 02-Sorting-and-Filtering/
│   ├── README.md
│   └── Bike_Sales_Sorting_Filtering.xlsx
│
├── 03-Outlier-Analysis/
│   ├── README.md
│   └── Bike_Sales_Outlier_Analysis.xlsx
│
├── 04-Pivot-Table-Analysis/
│   ├── README.md
│   └── Bike_Sales_Pivot_Table_Analysis.xlsx
│
├── 05-Data-Visualizations/
│   ├── README.md
│   └── Bike_Sales_Data_Visualizations.xlsx
│
└── README.md
```

---

# 1. Data Preparation and Cleaning

## Workbook

[`Bike_Sales_Data_Preparation.xlsx`](./01-Data-Preparation/Bike_Sales_Data_Preparation.xlsx)

## Goal

The goal of this workbook was to inspect the original Bike Sales dataset, identify data-quality problems, and prepare the data for reliable analysis.

Raw data may contain duplicates, blank cells, inconsistent spelling, extra spaces, abbreviated values, and formatting issues. These problems must be corrected before creating calculations, PivotTables, or visualizations.

## Work Completed

### Duplicate Detection

Conditional Formatting was used to highlight possible duplicate records.

After reviewing the highlighted records, Excel's **Remove Duplicates** feature was used where appropriate.

### Missing-Value Review

Blank and empty cells were identified and reviewed to determine whether values needed to be corrected, completed, or left blank.

### Data Parsing

Combined information was separated into individual columns using **Text to Columns**.

Delimiters included:

- Spaces
- Commas

This made the information easier to organize and analyze.

### Removing Extra Spaces

The `LEN` function was used to compare the number of characters in text values and identify country names that contained extra spaces.

```excel
=LEN(A2)
```

The `TRIM` function was then used to remove unnecessary leading, trailing, and repeated spaces.

```excel
=TRIM(A2)
```

### Standardizing Text Case

The `LOWER` function was used to convert text into lowercase during the cleaning process.

```excel
=LOWER(A2)
```

The `PROPER` and `TRIM` functions were combined to standardize country names.

```excel
=PROPER(TRIM(A2))
```

### Standardizing Gender Values

Find and Replace was used to convert abbreviated gender values into complete labels:

- `F` became `Female`
- `M` became `Male`

The Match Case option was used when necessary to prevent letters inside other words from being changed.

### Formatting Corrections

Inconsistent formatting in fields such as Age Group was removed and replaced with consistent text alignment and cell formatting.

## Excel Features Used

- Conditional Formatting
- Remove Duplicates
- Find and Replace
- Text to Columns
- Filters
- Cell Formatting
- `LEN`
- `TRIM`
- `LOWER`
- `PROPER`

## Skills Demonstrated

- Data cleaning
- Data preparation
- Duplicate detection
- Missing-value review
- Text parsing
- Data standardization
- Data validation
- Quality assurance

---

# 2. Sorting and Filtering

## Workbook

[`Bike_Sales_Sorting_Filtering.xlsx`](./02-Sorting-and-Filtering/Bike_Sales_Sorting_Filtering.xlsx)

## Goal

The goal of this workbook was to organize the Bike Sales dataset and create focused views of the data using sorting and filtering.

Sorting and filtering help analysts locate records, compare categories, identify high- and low-performing values, and answer specific business questions without changing the original dataset.

## Work Completed

The workbook demonstrates techniques such as:

- Sorting text alphabetically
- Sorting numbers from smallest to largest
- Sorting values from largest to smallest
- Sorting dates chronologically
- Applying multiple sorting levels
- Filtering records by customer characteristics
- Filtering by country or state
- Filtering by product category
- Filtering by revenue, cost, or profit
- Displaying only records that meet selected conditions

## Example Business Questions

Sorting and filtering can help answer questions such as:

- Which transactions generated the most revenue?
- Which orders produced the highest profit?
- Which countries had the largest sales values?
- Which products had the highest order quantities?
- Which customer age groups purchased certain products?
- Which records require further investigation?

## Excel Features Used

- Sort A to Z
- Sort Z to A
- Smallest to Largest
- Largest to Smallest
- Custom Sort
- Multi-Level Sort
- AutoFilter
- Number Filters
- Text Filters
- Date Filters

## Skills Demonstrated

- Data organization
- Record segmentation
- Business-data exploration
- Trend identification
- Filter-based analysis
- Attention to detail

---

# 3. Outlier Analysis

## Workbook

[`Bike_Sales_Outlier_Analysis.xlsx`](./03-Outlier-Analysis/Bike_Sales_Outlier_Analysis.xlsx)

## Goal

The goal of this workbook was to identify values that were unusually high or low compared with the rest of the dataset.

Outliers can represent genuine business events, data-entry errors, unusual customer behavior, exceptional sales performance, or transactions that require additional investigation.

## Work Completed

The workbook was used to:

- Review numerical variables for unusual values
- Identify exceptionally high or low observations
- Compare individual records with the general distribution
- Examine values such as revenue, cost, profit, unit price, and order quantity
- Visualize potential outliers
- Determine whether unusual values should be retained or investigated

## Why Outlier Analysis Matters

Outliers may affect:

- Averages
- Forecasts
- Business summaries
- Charts
- Statistical calculations
- Management decisions

An outlier should not automatically be removed. It should first be examined to determine whether it is an error or a valid business observation.

## Excel Features Used

- Sorting
- Conditional Formatting
- Summary calculations
- Charts
- Data comparison
- Minimum and maximum analysis

## Skills Demonstrated

- Outlier detection
- Data-quality analysis
- Statistical reasoning
- Pattern recognition
- Data validation
- Investigative analysis

---

# 4. PivotTable Analysis

## Workbook

[`Bike_Sales_Pivot_Table_Analysis.xlsx`](./04-Pivot-Table-Analysis/Bike_Sales_Pivot_Table_Analysis.xlsx)

## Goal

The goal of this workbook was to summarize the Bike Sales dataset using PivotTables and turn detailed transaction records into meaningful business information.

PivotTables make it possible to quickly group, calculate, filter, and compare large amounts of data without changing the original dataset.

## Analysis Performed

The workbook may be used to summarize information such as:

- Total revenue by country
- Total profit by country
- Revenue by product category
- Profit by product or subcategory
- Order quantity by customer age group
- Sales by gender
- Average unit price
- Total cost by region
- Product performance
- Customer-segment performance

## Business Value

The PivotTable analysis helps identify:

- High-performing markets
- Profitable products
- Important customer segments
- Revenue patterns
- Sales differences between categories
- Areas that may require further analysis

## Excel Features Used

- PivotTables
- Row fields
- Column fields
- Value fields
- Report filters
- Sum calculations
- Count calculations
- Average calculations
- Sorting within PivotTables
- PivotTable filtering

## Skills Demonstrated

- Data aggregation
- Business reporting
- KPI analysis
- Data summarization
- Category comparison
- Trend identification
- PivotTable development

---

# 5. Data Visualizations

## Workbook

[`Bike_Sales_Data_Visualizations.xlsx`](./05-Data-Visualizations/Bike_Sales_Data_Visualizations.xlsx)

## Goal

The goal of this workbook was to convert Bike Sales data and summary results into visual formats that are easier to understand and communicate.

Charts help business users recognize trends, comparisons, patterns, and unusual results more quickly than reviewing rows of raw data.

## Visual Analysis

The workbook demonstrates how visualizations can be used to present information such as:

- Revenue by country
- Profit by product category
- Sales by age group
- Product comparisons
- Cost and revenue relationships
- Sales trends
- High- and low-performing categories
- Potential outliers

## Visualization Principles

The workbook focuses on:

- Selecting an appropriate chart type
- Using clear and descriptive chart titles
- Labeling axes correctly
- Keeping charts easy to read
- Avoiding unnecessary visual clutter
- Highlighting important business findings
- Presenting comparisons accurately

## Excel Features Used

- Column charts
- Bar charts
- Line charts
- Chart titles
- Axis labels
- Legends
- Data labels
- Chart formatting
- Source-data selection

## Skills Demonstrated

- Data visualization
- Business storytelling
- Chart development
- Trend communication
- Visual comparison
- Reporting
- Presentation of analytical findings

---

# Technical Skills Demonstrated

- Microsoft Excel
- Data preparation
- Data cleaning
- Data validation
- Data standardization
- Conditional Formatting
- Text to Columns
- Find and Replace
- Sorting and filtering
- Outlier detection
- PivotTables
- Data aggregation
- Data visualization
- Business reporting
- Data-quality assurance

---

# Overall Project Outcome

Together, these workbooks demonstrate an Excel-based data analytics workflow:

1. Prepare and clean the raw data.
2. Organize records with sorting and filtering.
3. investigate unusual or extreme values.
4. Summarize detailed data using PivotTables.
5. Communicate results through charts and visualizations.

The completed projects demonstrate how raw transaction data can be transformed into reliable, organized, and understandable business information.

---

# Future Improvements

Possible future additions to this portfolio include:

- Creating an interactive Excel dashboard
- Adding slicers to PivotTables
- Developing KPI cards
- Performing monthly or yearly trend analysis
- Calculating profit margins
- Importing the cleaned data into SQL
- Creating a Power BI or Tableau dashboard
- Developing automated data-cleaning workflows
- Adding written business recommendations

---

# About Me

I am an aspiring Data Analyst and Business Administration graduate passionate about transforming raw data into useful information that supports business decisions.

My interests include data cleaning, business analytics, data visualization, dashboard development, workflow optimization, and process improvement.

---

# Connect With Me

**George Charles Jr.**

Data Analytics | Business Analytics | Operations Analytics

**LinkedIn:**  
[linkedin.com/in/george-charles-jr-bb8462399](https://www.linkedin.com/in/george-charles-jr-bb8462399)

**Email:**  
Gcharles15@icloud.com
