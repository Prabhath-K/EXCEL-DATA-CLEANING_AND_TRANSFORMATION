# Excel & Power Query Data Cleaning and Transformation



## 📋 Assignment Overview

This project demonstrates a rigorous approach to data preparation and hygiene using **Microsoft Excel** and **Power Query**. The dataset contains raw product information prone to common data anomalies, including missing fields, typographical errors, inconsistent text formatting, and duplicate entries.

The overarching design philosophy for this project is:

> **"Use evidence where available, use business rules where defined, and use `Unknown` or `Unclassified` when a value cannot be determined reliably."**

### Key Objectives

* Identify and rectify missing values without manufacturing false certainty.

* Standardize inconsistent text formats and eliminate typographical errors.

* Construct scalable, reusable lookup tables using Power Query.

* Implement a hierarchical imputation strategy for missing numerical data.

* Extract, transform, and format identifiers, dates, and currency values.

* Apply professional conditional formatting for visual data analysis.

## 🗂️ Dataset Schema

| 

| **Field Name** | **Description** | **Data Type** | 
| **Product ID** | Unique composite identifier (`DD-MMM-COUNTRY`) | Text | 
| **Product Name** | Name of the product (subject to capitalization inconsistencies) | Text | 
| **Brand Name** | Manufacturer or brand name | Text | 
| **Price (\$)** | Unit price of the product (contains missing values) | Numeric (Currency) | 
| **Quantity** | Inventory or transaction quantity | Numeric | 
| **Category** | Product classification group (contains missing values & typos) | Text | 

## ⚙️ Data Cleaning & Transformation Workflow

```
graph TD
    A[Load Raw Dataset] --> B[Standardize Product Names]
    B --> C[Fix Category Typos: Electroni -> Electronics]
    C --> D[Create CategoryLookup Table]
    D --> E[Fill Missing Categories via Left Outer Join]
    E --> F[Fallback Unresolved Categories to 'Unclassified']
    F --> G[Create ProductPriceLookup Table]
    G --> H[Impute Missing Prices via Product Median]
    H --> I[Fallback to Category Median for Unmatched Prices]
    I --> J[Remove Exact Duplicate Rows]
    J --> K[Split Product ID into Components]
    K --> L[Generate Manufacturing Date & Country Code]
    L --> M[Merge Brand Name + Product Name ('Product Brand')]
    M --> N[Load Cleaned Data to Excel]
    N --> O[Apply Currency & Conditional Formatting]

```

## 🔍 Key Methodological Decisions

### 1. Handling Missing Categories (Avoiding Imputation Bias)

Rather than blindly assigning categories or deleting records, a strict hierarchy was enforced:

1. **Existing Product Mapping:** Look up matching product names from trusted rows (e.g., `Sneakers | Adidas` mapped to `Fashion` based on `Sneakers | Nike`).

2. **Unclassified Fallback:** If insufficient evidence exists (e.g., a unique `Backpack | North Face` with no peer records), the category defaults strictly to `Unclassified` or `Unknown` rather than fabricating certainty.

### 2. Hierarchical Price Imputation

Global medians distort product value segments. A tiered fallback strategy was used: 

$$
\text{Original Price} \longrightarrow \text{Product-Level Median} \longrightarrow \text{Category-Level Median} \longrightarrow \text{Remain Missing / Flagged}
$$

### 3. Date & ID Parsing

The `Product ID` string (format: `DD-MMM-Country`, e.g., `28-JAN-US`) was split using delimiters. Because the string intentionally omits a year, the `Manufacturing Date` is preserved strictly as a text value (`DD-MMM`) to prevent the injection of artificial, unverified temporal data.

## 🛠️ Power Query Concepts & Operations Applied

| **Power Query Feature** | **Implementation Purpose** | 
| **Replace Values** | Corrected typos such as `Electroni` to `Electronics`. | 
| **Reference & Group By** | Built isolated reference tables (`CategoryLookup`, `ProductPriceLookup`) to calculate medians. | 
| **Merge Queries (Left Outer Join)** | Dynamically joined lookup tables back to the master dataset while preserving row integrity. | 
| **Conditional Columns** | Implemented custom fallback logic (`if [Category] = null then ... else ...`). | 
| **Split & Merge Columns** | Deconstructed `Product ID` into components and merged `Brand Name` with `Product Name`. | 
| **Remove Duplicates** | Filtered out full-row exact duplicates to maintain record accuracy. | 

## 📊 Excel Formatting & Presentation Layer

* **Currency Formatting:** Applied standard localized currency formatting (`$#,##0.00`) to the `Price` column while retaining underlying numeric properties.

* **Data Bars:** Added visual data bars to the `Price` column for rapid comparative price analysis.

* **Conditional Formatting Rules:** Automatically highlighted all rows matching `Category = Electronics` to improve report scanning efficiency.

## 🚀 Getting Started & Reproducibility

1. Open the raw dataset inside **Microsoft Excel**.

2. Open the **Power Query Editor** (`Data > Get Data > Launch Power Query Editor`).

3. Follow the sequence outlined in the [Workflow](#-data-cleaning--transformation-workflow) to ensure dependency lookup queries (`CategoryLookup`, `ProductPriceLookup`) are evaluated prior to master table merges.

4. Close and Load the transformations back into the Excel worksheet workbook.
