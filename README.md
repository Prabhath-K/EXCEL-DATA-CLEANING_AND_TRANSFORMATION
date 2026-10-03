# Excel Data Cleaning and Transformation Using Power Query

## Project Overview

This project demonstrates a practical data-cleaning and transformation workflow using **Microsoft Excel and Power Query**.

The dataset contains product-related information such as:

- Product ID
- Product Name
- Brand Name
- Price
- Quantity
- Category

The main objective of this project is to identify and resolve common data-quality issues, transform columns into useful structures, and prepare the dataset for analysis and reporting.

The project covers the following tasks:

1. Handling Missing Values
2. Correcting Inconsistent Data
3. Removing Duplicates
4. Splitting and Merging Data
5. Number Formatting
6. Conditional Formatting

The main data-quality principle followed throughout this project is:

> **Use evidence where available, use business rules where defined, and use `Unknown` or `Unclassified` when a value cannot be determined reliably.**

---

## Dataset Structure

The original dataset contains the following columns:

| Column | Description |
|---|---|
| Product ID | Encoded value containing day, month, and country code |
| Product Name | Name of the product |
| Brand Name | Brand of the product |
| Price ($) | Product price |
| Quantity | Available quantity |
| Category | Product category |

Example:

| Product ID | Product Name | Brand Name | Price ($) | Quantity | Category |
|---|---|---|---:|---:|---|
| 28-JAN-US | laptop | Dell | 1000 | 30 | Electronics |
| 15-FEB-US | Sneakers | Nike | 80 | 15 | Fashion |
| 03-MAR-US | Coffee Maker | Keurig | 130 | 40 | Kitchen |
| 11-APR-US | smartphone | Samsung | 900 | 25 | Electronics |
| 22-MAY-US | Backpack | North Face | 70 | 20 | null |
| 07-JUN-UK | Headphones | Sony | null | 45 | Electroni |

---

# 1. Handling Missing Values

## 1.1 Missing Values in Category

### Question

> If there are products with missing categories, propose a strategy to impute or deal with these missing values effectively.

### Missing Category Records

The following products contain missing Category values:

| Product Name | Brand | Category | Inference |
|---|---|---|---|
| Backpack | North Face | null | Ambiguous |
| Sneakers | Adidas | null | Fashion |
| Coffee Maker | Nespresso | null | Kitchen |
| Fitness Tracker | Xiaomi | null | Electronics |

### Reasoning

The first approach is to check whether the same `Product Name` already appears elsewhere in the dataset with a valid category.

For example:

```text
Sneakers | Nike | Fashion
```

Therefore:

```text
Sneakers | Adidas | Fashion
```

is a reasonable category mapping because the product type is the same.

Similarly:

```text
Coffee Maker | Keurig | Kitchen
```

supports:

```text
Coffee Maker | Nespresso | Kitchen
```

And:

```text
Fitness Tracker | Garmin | Electronics
```

supports:

```text
Fitness Tracker | Xiaomi | Electronics
```

Although the brands are different, the product type is the same, so the category can be inferred using the existing product mapping.

---

## 1.2 Why Backpack Is Different

The following row has no category:

```text
Backpack | North Face | 70 | 20 | null
```

There is no other `Backpack` record in the dataset with a known category.

Possible values such as:

```text
Fashion
```

or:

```text
Accessories
```

could both be reasonable.

However:

- Price does not determine the category.
- Quantity does not determine the category.
- Brand Name alone does not determine the category.
- No existing Backpack row provides a trusted mapping.

Therefore, the dataset itself does not contain enough information to classify Backpack confidently.

### Decision

Use:

```text
Unclassified
```

or:

```text
Unknown
```

until an authoritative product master table or business rule becomes available.

This avoids inventing information simply to eliminate a null value.

### Final Category Cleaning Decision

| Product Name | Original Category | Cleaned Category |
|---|---|---|
| Backpack | null | Unclassified |
| Sneakers | null | Fashion |
| Coffee Maker | null | Kitchen |
| Fitness Tracker | null | Electronics |

> **Key Principle:** Do not manufacture certainty only to remove NULL values.

---

# 2. Creating a Category Lookup Table

Instead of manually replacing missing categories row by row, I created a reusable lookup table in Power Query.

The lookup table is named:

```text
CategoryLookup
```

## 2.1 Why Use a Lookup Table?

A lookup table:

- Keeps category rules centralized
- Avoids repetitive manual replacements
- Makes cleaning easier to audit
- Can be reused in future datasets
- Scales better for larger datasets
- Works similarly to a reference table in a database

## 2.2 Preparing the Category Lookup

### Power Query Steps

1. Create a **Reference** or **Duplicate** of the master query.
2. Keep only:
   - `Product Name`
   - `Category`
3. Correct category spelling errors.
4. Remove rows where `Category` is null.
5. Remove duplicate rows.
6. Ensure only one valid mapping remains for each Product Name.
7. Rename the query to:

```text
CategoryLookup
```

Rows such as:

```text
Sneakers | null
Coffee Maker | null
Fitness Tracker | null
Backpack | null
```

should not remain in the lookup table.

A lookup table should contain only valid and trusted mappings.

---

# 3. Correcting Category Typos

The dataset contains:

```text
Electroni
```

which is a misspelling of:

```text
Electronics
```

This value appears for products such as:

- Headphones
- Smartwatch
- Bose Headphones
- Huawei Smartwatch

## Why Correct the Typo Before Creating the Lookup?

If the typo is not corrected first, the lookup may contain conflicting values such as:

```text
Headphones | Electroni
Headphones | Electronics
```

Power Query would treat these as different values.

### Power Query Steps

Go to:

```text
Transform → Replace Values
```

Replace:

```text
Electroni
```

with:

```text
Electronics
```

---

# 4. Final Category Lookup Table

After removing null categories, correcting typos, and removing duplicates, the lookup table should contain one trusted category per Product Name.

| Product Name | Category |
|---|---|
| Laptop | Electronics |
| Sneakers | Fashion |
| Coffee Maker | Kitchen |
| Smartphone | Electronics |
| Headphones | Electronics |
| T-Shirt | Fashion |
| Blender | Kitchen |
| Tablet | Electronics |
| Hiking Boots | Outdoor |
| Smartwatch | Electronics |
| Laptop Bag | Accessories |
| Sunglasses | Fashion |
| Camping Tent | Outdoor |
| Camera | Electronics |
| Microwave | Kitchen |
| Dress | Fashion |
| Toaster | Kitchen |
| Fitness Tracker | Electronics |
| Jeans | Fashion |
| Watch | Accessories |

The lookup table now contains:

- No null categories
- Standardized category names
- No duplicate mappings
- One valid category per Product Name

---

# 5. Filling Missing Categories Using Merge Queries

Once `CategoryLookup` is ready, it can be merged back into the master table.

## Power Query Steps

1. Open the master query.
2. Select:

```text
Home → Merge Queries
```

3. Match:

```text
Master[Product Name]
```

with:

```text
CategoryLookup[Product Name]
```

4. Use:

```text
Left Outer Join
```

5. Expand the lookup category field.
6. Create a custom column.

### Custom Column Formula

```powerquery
if [Category] = null
then [CategoryLookup.Category]
else [Category]
```

Name the new column:

```text
Clean Category
```

Then:

1. Remove the old Category column.
2. Remove the helper lookup column.
3. Rename `Clean Category` to `Category`.
4. Replace any remaining null value with:

```text
Unclassified
```

### Final Category Logic

```text
Original Category
       ↓
If null → CategoryLookup
       ↓
If still null → Unclassified
```

---

# 6. Handling Missing Values in Price

### Question

> Check for missing values in the Price column. How would you handle products with missing price information?

The dataset contains missing prices for:

| Product | Brand | Price |
|---|---|---:|
| Headphones | Sony | null |
| Camping Tent | Coleman | null |
| Sunglasses | Ray-Ban | null |

---

# 7. Why I Did Not Use the Overall Median Immediately

Price can vary significantly depending on:

- Product
- Brand
- Product type
- Market segment

For example:

```text
Headphones | Bose | 250
```

does not prove that:

```text
Headphones | Sony
```

also costs `$250`.

Similarly:

```text
Sunglasses | Oakley | 150
```

does not prove that:

```text
Sunglasses | Ray-Ban
```

costs exactly `$150`.

Therefore, replacing every missing value using the overall dataset median could create misleading values.

---

# 8. Preferred Price Imputation Hierarchy

The preferred hierarchy is:

1. Official Product/Master Price Table
2. Same Product + Brand from another valid record
3. Product-level median
4. Category-level median as a fallback
5. Keep the value missing if insufficient evidence exists

For this practice dataset, the fallback process is:

```text
Original Price
      ↓
Product Median
      ↓
Category Median
```

---

# 9. Creating ProductPriceLookup

To estimate missing prices using existing product prices, I created a product-level price lookup.

## Power Query Steps

1. Reference or duplicate the master query.
2. Keep only:
   - `Product Name`
   - `Price ($)`
3. Remove rows where `Price ($)` is null.
4. Go to:

```text
Transform → Group By
```

5. Group by:

```text
Product Name
```

6. Create:

```text
Median Price
```

7. Use the `Median` operation on:

```text
Price ($)
```

8. Rename the query:

```text
ProductPriceLookup
```

### Example Product Price Lookup

| Product Name | Median Price |
|---|---:|
| Laptop | 980 |
| Sneakers | 85 |
| Coffee Maker | 125 |
| Headphones | 250 |
| Smartphone | 850 |
| Sunglasses | 150 |
| Blender | 245 |
| Fitness Tracker | 140 |

---

# 10. Merging ProductPriceLookup with Master Data

## Steps

1. Open the master query.
2. Select:

```text
Home → Merge Queries
```

3. Match both queries using:

```text
Product Name
```

4. Use:

```text
Left Outer Join
```

5. Expand:

```text
Median Price
```

6. Create a custom column.

### Formula

```powerquery
if [Price ($)] = null
then [Median Price]
else [Price ($)]
```

Name the new column:

```text
Clean Price
```

### Expected Result

```text
Sony Headphones → 250
Ray-Ban Sunglasses → 150
```

These values are imputed estimates, not confirmed brand-specific prices.

---

# 11. Handling Products Without Product-Level Price

`Camping Tent` cannot receive a product-level median because there is no other valid Camping Tent price in the dataset.

Therefore, a second fallback is required.

---

# 12. Creating CategoryPriceLookup

Create another lookup query called:

```text
CategoryPriceLookup
```

## Power Query Steps

1. Reference the master table.
2. Keep:
   - `Category`
   - `Price ($)`
3. Remove null prices.
4. Go to:

```text
Transform → Group By
```

5. Group by:

```text
Category
```

6. Calculate:

```text
Median of Price ($)
```

7. Merge the result back into the master table using `Category`.

### Final Price Formula

```powerquery
if [Price ($)] <> null then [Price ($)]
else if [MedianPrice] <> null then [MedianPrice]
else [MedianCategoryPrice]
```

The logic becomes:

```text
Original Price
      ↓
Product Median
      ↓
Category Median
```

### Important Limitation

For a small dataset, a category median may be calculated from very few records.

For example:

```text
Hiking Boots → 130
Camping Tent → null
```

If these are the only Outdoor records, then:

```text
Camping Tent → 130
```

is only an estimate.

It is not the confirmed price of the Camping Tent.

For a real business dataset, the preferred solution would be an official master pricing table.

---

# 13. Correcting Inconsistent Product Name Formats

The `Product Name` column contains inconsistent capitalization.

Examples:

```text
laptop
Laptop
```

```text
smartphone
Smartphone
```

```text
headphones
Headphones
```

These values should be standardized before:

- Creating lookup tables
- Grouping
- Merging queries
- Removing duplicates

## Power Query Approach

Go to:

```text
Transform → Format
```

and apply:

```text
Capitalize Each Word
```

where appropriate.

Known inconsistencies can also be corrected using:

```text
Transform → Replace Values
```

### Why This Matters

Without standardization:

```text
laptop
```

and:

```text
Laptop
```

may behave as separate values during:

- Filtering
- Grouping
- Joins
- Lookups
- Duplicate detection
- Reporting

---

# 14. Removing Duplicate Rows

The dataset contains exact duplicate records.

Examples include:

```text
21-AUG-CA | Laptop Bag | Samsonite | 50 | 35 | Accessories
```

```text
16-APR-ES | Headphones | Bose | 250 | 20 | Electronics
```

```text
17-JUN-IN | Laptop | HP | 950 | 25 | Electronics
```

## Power Query Steps

1. Select all columns.
2. Go to:

```text
Home → Remove Rows → Remove Duplicates
```

Power Query keeps the first occurrence and removes exact repeated rows.

### Why Remove Duplicates Using the Entire Row?

Two records can legitimately have the same:

- Product Name
- Category
- Price

but still represent different records.

Therefore, duplicate removal should be based on the entire row rather than only one field.

---

# 15. Splitting Product ID

The `Product ID` column contains multiple pieces of information.

Examples:

```text
28-JAN-US
17-JUN-IN
09-JUL-FR
```

The structure is:

```text
Day-Month-Country Code
```

The assignment requires this to be separated into:

- Manufacturing Date
- Country Code

---

# 16. Split Column in Power Query

## Steps

1. Select:

```text
Product ID
```

2. Go to:

```text
Transform → Split Column → By Delimiter
```

3. Select:

```text
-
```

as the delimiter.

4. Choose:

```text
Each occurrence of the delimiter
```

Power Query creates three columns:

| Day | Month | Country Code |
|---:|---|---|
| 28 | JAN | US |
| 15 | FEB | US |
| 03 | MAR | US |
| 17 | JUN | IN |
| 09 | JUL | FR |

Rename the fields:

```text
Day
Month
Country Code
```

---

# 17. Creating Manufacturing Date

After splitting the Product ID, merge the Day and Month columns.

## Steps

1. Select:
   - `Day`
   - `Month`

2. Go to:

```text
Transform → Merge Columns
```

3. Choose:

```text
-
```

as the separator.

4. Name the result:

```text
Manufacturing Date
```

### Example

| Day | Month | Manufacturing Date |
|---:|---|---|
| 28 | JAN | 28-JAN |
| 17 | JUN | 17-JUN |
| 09 | JUL | 09-JUL |

Final structure:

| Manufacturing Date | Country Code |
|---|---|
| 28-JAN | US |
| 17-JUN | IN |
| 09-JUL | FR |

---

# 18. Manufacturing Date Limitation

The original Product ID contains only:

```text
Day-Month-Country
```

For example:

```text
28-JAN-US
```

contains:

```text
Day = 28
Month = JAN
Country Code = US
```

but it does not contain a year.

Therefore:

```text
28-JAN
```

is not a complete calendar date.

A true value in:

```text
DD-MM-YYYY
```

format cannot be created reliably unless a year is supplied from another trusted source.

## Decision

Keep `Manufacturing Date` as a text field until a valid year becomes available.

This prevents the creation of an artificial or incorrect year.

---

# 19. Merging Brand Name and Product Name

The assignment requires:

```text
Brand Name
```

and:

```text
Product Name
```

to be merged into a new field:

```text
Product Brand
```

## Power Query Steps

1. Select `Brand Name`.
2. Hold `Ctrl` and select `Product Name`.
3. Go to:

```text
Transform → Merge Columns
```

4. Choose `Space` as the separator.
5. Name the new field:

```text
Product Brand
```

### Example

| Brand Name | Product Name | Product Brand |
|---|---|---|
| Dell | Laptop | Dell Laptop |
| Nike | Sneakers | Nike Sneakers |
| Keurig | Coffee Maker | Keurig Coffee Maker |
| Samsung | Smartphone | Samsung Smartphone |
| Casio | Watch | Casio Watch |

### Why I Used This Approach

The merged field is more descriptive.

For example:

```text
Laptop
```

does not identify the brand.

But:

```text
Dell Laptop
HP Laptop
Asus Laptop
```

clearly distinguishes the products.

This makes the field more useful for:

- Reporting
- Filtering
- Dashboards
- Visualizations

---

# 20. Number Formatting

## Price Formatting

The `Price ($)` column should remain numeric during transformations.

After loading the cleaned data back into Excel:

1. Select the Price column.
2. Go to:

```text
Home → Number Format
```

3. Select:

```text
Currency
```

4. Choose `$` as the currency symbol.

### Example

Before:

```text
1000
```

After:

```text
$1,000.00
```

This improves readability without changing the underlying numeric value.

---

# 21. Manufacturing Date Formatting

If a valid year becomes available and the field is converted into a true Excel Date:

1. Select the Manufacturing Date column.
2. Press:

```text
Ctrl + 1
```

3. Select:

```text
Custom
```

4. Enter:

```text
dd-mm-yyyy
```

Example:

```text
28-01-2026
```

This should only be applied if the year comes from a valid source.

---

# 22. Conditional Formatting

## 22.1 Apply Data Bars to Price

To visually compare product prices:

1. Select the Price cells.
2. Do not include the header.
3. Go to:

```text
Home → Conditional Formatting → Data Bars
```

4. Choose the preferred style.

### Why Use Data Bars?

A high price displays a longer bar than a low price.

For example:

```text
$1,000
```

will display a longer bar than:

```text
$30
```

This makes price differences easy to compare visually.

---

# 23. Highlight Electronics Category

The requirement is to highlight Category cells where:

```text
Electronics
```

appears.

## Method 1: Format Cells That Contain

1. Select the Category column.
2. Go to:

```text
Home → Conditional Formatting → New Rule
```

3. Choose:

```text
Format only cells that contain
```

4. Set:

```text
Cell Value → equal to → Electronics
```

5. Choose a fill color.
6. Click `OK`.

## Method 2: Custom Formula

If the Category field is in column `E` and the first data row is row `2`, use:

```excel
=$E2="Electronics"
```

The `$` locks column `E`, while the row number changes automatically.

---

# 24. Overall Data Cleaning Workflow

The complete workflow is:

```text
Load Dataset
      ↓
Standardize Product Names
      ↓
Fix Category Typo
      ↓
Create CategoryLookup
      ↓
Fill Missing Categories
      ↓
Replace Remaining Category Nulls with Unclassified
      ↓
Create ProductPriceLookup
      ↓
Fill Missing Price Using Product Median
      ↓
Create CategoryPriceLookup
      ↓
Use Category Median as Fallback
      ↓
Remove Exact Duplicate Rows
      ↓
Split Product ID
      ↓
Create Manufacturing Date
      ↓
Create Country Code
      ↓
Merge Brand Name + Product Name
      ↓
Load Cleaned Data into Excel
      ↓
Apply Currency Formatting
      ↓
Apply Conditional Formatting
```

---

# 25. Key Analytical Decisions

## Missing Categories

The hierarchy used was:

```text
Existing Product Mapping
       ↓
Business / Master Mapping
       ↓
Unclassified
```

I did not assign categories randomly.

## Missing Prices

The hierarchy used was:

```text
Original Price
       ↓
Same Product + Brand
       ↓
Product-Level Median
       ↓
Category-Level Median
       ↓
Trusted External Source / Remain Missing
```

I did not immediately replace every missing price with the overall dataset median.

## Backpack

The Backpack category remains:

```text
Unclassified
```

unless a trusted business rule or master category source states otherwise.

A clearly documented unknown value is better than assigning an unsupported category simply to remove a null.

## Manufacturing Date

The original Product ID contains:

```text
Day-Month-Country
```

but no year.

Therefore, I did not create an artificial year.

The Manufacturing Date remains a day-month text value until a valid year becomes available.

---

# 26. Power Query Concepts Used

| Power Query Feature | Purpose |
|---|---|
| Replace Values | Correct `Electroni` to `Electronics` |
| Format Text | Standardize Product Name capitalization |
| Remove Nulls | Prepare lookup tables |
| Remove Duplicates | Remove duplicate mappings and records |
| Group By | Calculate product and category median prices |
| Merge Queries | Join lookup tables with the master dataset |
| Left Outer Join | Preserve all rows from the master table |
| Custom Column | Apply fallback cleaning logic |
| Split Column | Separate Product ID components |
| Merge Columns | Create Manufacturing Date and Product Brand |
| Data Types | Maintain correct numeric and text structures |

---

# 27. Data Quality Principles Followed

1. **Do not invent missing information without evidence.**
2. **Use existing product mappings when reliable.**
3. **Use lookup tables instead of repeated manual replacements.**
4. **Standardize text before grouping, merging, or duplicate detection.**
5. **Use stronger evidence before broader fallback rules.**
6. **Treat imputed prices as estimates, not confirmed facts.**
7. **Use `Unknown` or `Unclassified` when the value cannot be determined.**
8. **Document assumptions clearly.**
9. **Keep the cleaning process reproducible.**
10. **Preserve the difference between original and inferred information.**

---

# 28. Final Outcome

After completing the cleaning workflow, the dataset becomes:

- More consistent
- Easier to analyze
- Free from known exact duplicates
- Standardized in text format
- Improved in missing-value handling
- Structured into useful fields
- Ready for reporting and visualization

The final dataset includes transformed fields such as:

```text
Manufacturing Date
Country Code
Product Brand
```

and cleaned fields such as:

```text
Category
Price ($)
Product Name
```

---

# Conclusion

This project demonstrates a practical and scalable approach to **Excel Data Cleaning and Transformation using Power Query**.

Rather than simply replacing every missing value or correcting visible errors manually, the workflow uses:

- Reusable lookup tables
- Merge Queries
- Product-level mappings
- Median-based price estimation
- Category-level fallback logic
- Text standardization
- Duplicate removal
- Column splitting
- Column merging
- Currency formatting
- Conditional formatting
- Transparent handling of uncertain values

The most important lesson from this project is:

> **A clean dataset should not only contain fewer errors. It should also preserve the difference between known information, inferred information, and genuinely unknown information.**
