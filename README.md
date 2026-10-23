# 📊 PR. 2 Analyzer – Sales and Profit Analytics Project

![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Data Analysis](https://img.shields.io/badge/Data%20Analysis-Sales%20%26%20Profit-blue?style=for-the-badge)
![Status](https://img.shields.io/badge/Purpose-Educational-orange?style=for-the-badge)

---

## 📑 Table of Contents

- [Project Overview](#-project-overview)
- [Project Objectives](#-project-objectives)
- [Tools and Technologies](#-tools-and-technologies)
- [Dataset Description](#-dataset-description)
- [Project Structure](#-project-structure)
- [Key Features and Analysis](#-key-features-and-analysis)
- [Excel Functions Used](#-excel-functions-used)
- [Project Workflow](#-project-workflow)
- [Business Applications](#-business-applications)
- [How to Use the Project](#-how-to-use-the-project)
- [Learning Outcomes](#-learning-outcomes)
- [Project Limitations](#-project-limitations)
- [Future Enhancements](#-future-enhancements)
- [Conclusion](#-conclusion)
- [Author](#-author)
- [Repository](#-repository)

---

## 📌 Project Overview

**PR. 2 Analyzer – Sales and Profit Analytics Project** is a Microsoft Excel-based data analysis project designed to transform raw sales data into meaningful business insights using Excel formulas, statistical analysis, conditional formatting, PivotTables, charts, and dashboard techniques.

The project focuses on understanding:

- Sales performance
- Profitability
- Customer purchasing behavior
- Regional performance
- Product category trends
- Monthly sales growth

It demonstrates practical Excel skills for data cleaning, data analysis, business reporting, statistical modeling, and data visualization.

---

## 🎯 Project Objectives

1. Analyze sales and profit data to understand overall business performance.
2. Identify high-value customers based on total purchases.
3. Compare sales performance across different regions and product categories.
4. Calculate monthly sales totals and growth percentages.
5. Apply What-If Analysis to evaluate hypothetical discount scenarios.
6. Perform descriptive statistical analysis on sales and profit.
7. Use linear regression to study the relationship between sales and profit.
8. Apply conditional formatting to highlight important values and trends.
9. Create PivotTables for multidimensional sales analysis.
10. Develop charts and a dashboard to present business information clearly.

---

## 🛠 Tools and Technologies

| Tool or Feature | Purpose |
|---|---|
| Microsoft Excel | Data analysis and reporting |
| Excel Formulas | Calculations and data analysis |
| Conditional Formatting | Highlight important values and trends |
| What-If Analysis | Evaluate hypothetical business scenarios |
| Data Analysis ToolPak | Regression and descriptive statistics |
| PivotTables | Summarize and compare sales data |
| Excel Charts | Visualize sales, profit, and trends |
| Dashboard | Present key performance indicators and analytical results |

---

## 🗂 Dataset Description

The project uses a sales dataset containing customer information, regional details, product categories, sales values, quantities, discounts, order dates, and profits.

| Column | Description |
|---|---|
| `Customer_ID` | Unique identifier assigned to a customer |
| `Customer_Name` | Name of the customer |
| `Region` | Geographical region associated with the transaction |
| `Product_Category` | Category of the product sold |
| `Sales` | Sales amount for the transaction |
| `Quantity` | Number of units sold |
| `Discount` | Discount applied to the transaction |
| `Order_Date` | Date of the transaction |
| `Profit` | Profit associated with the transaction |

> **Note:** The dataset's exact number of records and date range depend on the supplied workbook. The analysis should be interpreted using the actual data and calculated results.

---

## 🏗 Project Structure

The workbook is organized into analytical sections, with each worksheet serving a specific purpose.

| Worksheet | Purpose |
|---|---|
| `Sales_Data` | Stores the original sales dataset |
| `Conditional_Formatting` | Highlights important values and customer performance |
| `What_If_Analysis` | Evaluates hypothetical discount scenarios |
| `Regression` | Examines the relationship between sales and profit |
| `Descriptive_Stats` | Presents statistical summaries |
| `Customer_Analysis` | Analyzes customer purchases and rankings |
| `Monthly_Analysis` | Calculates monthly sales and growth trends |
| `Pivot_Analysis` | Summarizes sales using PivotTables |
| `Dashboard` | Presents key metrics and visualizations |

> Worksheet names may differ slightly depending on the final version of the workbook.

---

## 🔍 Key Features and Analysis

### 1. Data Preparation and Formatting

The initial stage involves preparing the dataset for analysis.

- Reviewing column headings and data types
- Checking that sales, quantity, discount, profit, and date values use appropriate formats
- Formatting the source data as an Excel Table
- Applying consistent number and date formats
- Preparing the dataset for formulas, PivotTables, and charts

**Objective:** Establish a structured dataset that supports reliable calculations and analysis.

---

### 2. Conditional Formatting

Conditional Formatting is used to highlight important values and make the analysis easier to interpret.

- Highlighting the top-performing customers
- Identifying high-value purchase totals
- Applying different colors to distinguish important results
- Displaying visual indicators for monthly sales trends

For example, the following formula classifies customers based on their purchase ranking:

```excel
=IF(D2<=10,"High Value","Regular")
```

Here, `D2` represents the customer's rank. Customers ranked in the top ten are classified as **High Value**.

**Objective:** Improve readability and identify significant results quickly.

---

### 3. Customer Purchase Analysis

Customer Analysis evaluates how much each customer contributes to overall sales.

- Customer identification
- Customer purchase totals
- Customer ranking based on total purchases
- Identification of high-value customers
- Comparison of customer purchasing behavior

#### Total Purchases

The `SUMIF` function calculates the total sales associated with a particular customer.

```excel
=SUMIF(Sales_Data!$A$2:$A$201,A2,Sales_Data!$E$2:$E$201)
```

This formula assumes column A contains `Customer_ID` and column E contains `Sales`. Cell `A2` in the analysis sheet contains the customer ID being analyzed.

#### Customer Ranking

```excel
=RANK.EQ(C2,$C$2:$C$31,0)
```

This formula ranks customers according to their total purchase amounts, with the highest amount receiving rank 1.

> The referenced ranges should be adjusted to match the number of customers in the completed analysis.

**Objective:** Identify valuable customers and understand customer-level sales contributions.

---

### 4. What-If Analysis

What-If Analysis evaluates how hypothetical changes in discount assumptions may affect calculated profit.

Discount scenarios used:

- 0% discount
- 5% discount
- 10% discount
- 15% discount
- 20% discount

The resulting values can be compared in a table and visualized using a column chart.

#### Example Scenario Formula

```excel
=SUM(Sales_Data!$I$2:$I$201)+SUMPRODUCT(Sales_Data!$E$2:$E$201,Sales_Data!$G$2:$G$201-A2)
```

This example assumes that:

- Column E contains `Sales`
- Column G contains `Discount`
- Column I contains `Profit`
- Cell `A2` contains the scenario discount rate

> ⚠️ **Important:** This is an illustrative scenario formula, not a verified accounting formula. Its suitability depends on how the original profit and discount fields are defined and how discounts affect costs, revenue, and profit. Validate the calculation against the dataset's business rules before interpreting the results as actual projected profit.

**Objective:** Understand how different discount assumptions may influence a modeled business outcome.

---

### 5. Regression Analysis

Regression Analysis is used to study the relationship between sales and profit.

- **Independent Variable (X):** Sales
- **Dependent Variable (Y):** Profit

Excel's Data Analysis ToolPak can be used to perform linear regression.

#### Regression Procedure

1. Open the Excel workbook.
2. Enable the Analysis ToolPak if it is not already enabled.
3. Navigate to the **Data** tab.
4. Select **Data Analysis**.
5. Choose **Regression**.
6. Set the Sales range as the **Input X Range**.
7. Set the Profit range as the **Input Y Range**.
8. Select **Labels** if the selected ranges include column headings.
9. Choose an output location.
10. Click **OK** to generate the regression results.

#### Important Regression Metrics

| Metric | Purpose |
|---|---|
| R Square | Measures the proportion of variation in the dependent variable explained by the model |
| Multiple R | Represents the correlation coefficient reported by Excel's regression output |
| Coefficients | Describe the fitted relationship between the independent and dependent variables |
| P-value | Helps assess statistical evidence for an individual coefficient |
| Standard Error | Measures uncertainty in the estimated regression relationship |

The actual regression output should be used when reporting coefficients, R Square, p-values, and other numerical results.

**Objective:** Explore whether sales values help explain variations in profit.

> **Limitation:** A regression relationship does not establish causation, and its usefulness depends on the dataset, assumptions, and model fit.

---

### 6. Descriptive Statistics

Descriptive Statistics provides a statistical summary of selected numerical columns, including Sales and Profit.

| Statistic | Description |
|---|---|
| Mean | Arithmetic average |
| Median | Middle value in an ordered dataset |
| Mode | Most frequently occurring value |
| Standard Deviation | Measures the spread of values around the mean |
| Variance | Measures squared dispersion around the mean |
| Minimum | Lowest observed value |
| Maximum | Highest observed value |
| Range | Difference between maximum and minimum |
| Count | Number of numerical observations included |

#### Procedure

1. Navigate to the **Data** tab.
2. Click **Data Analysis**.
3. Select **Descriptive Statistics**.
4. Select the required Sales or Profit input range.
5. Enable **Labels in First Row** if headings are included.
6. Select **Summary Statistics**.
7. Choose the output location.
8. Click **OK**.

**Objective:** Understand the distribution, average values, and variability of sales and profit.

---

### 7. Monthly Sales Analysis

Monthly Analysis summarizes sales by month and calculates changes relative to the previous month.

The analysis includes: **Month**, **Total Sales**, **Previous Month Sales**, **Sales Growth**, and **Trend Indicator**.

#### Monthly Sales Calculation

```excel
=SUMIFS(Sales_Data!$E$2:$E$201,Sales_Data!$H$2:$H$201,">="&A2,Sales_Data!$H$2:$H$201,"<"&EDATE(A2,1))
```

This formula sums sales for the month beginning at the date in cell `A2`. It assumes column E contains `Sales`, column H contains `Order_Date`, and `A2` contains the first day of the month.

#### Previous Month Sales

```excel
=B2
```

Enter this in the appropriate next-month row, such as `C3`, to reference the previous month's total.

#### Monthly Sales Growth

```excel
=IFERROR((B3-C3)/C3,0)
```

Format the result as a percentage. This calculation measures the percentage change between the current month and the previous month. If the previous month's sales are zero, the formula returns zero rather than a division error; interpret such cases carefully.

#### Trend Indicator

```excel
=IF(B3>C3,"↑",IF(B3<C3,"↓","→"))
```

| Symbol | Meaning |
|---|---|
| ↑ | Current monthly sales exceed the previous month's sales |
| ↓ | Current monthly sales are lower |
| → | Both values are equal |

Conditional Formatting can apply different colors to the trend indicators.

**Objective:** Track monthly performance and make changes in sales easier to identify.

---

### 8. PivotTable Analysis

PivotTables summarize large datasets into meaningful categories without manually writing separate formulas for every combination.

#### Example PivotTable Configuration

| PivotTable Area | Field |
|---|---|
| Rows | `Region` |
| Columns | `Product_Category` |
| Values | `Sales` |
| Aggregation | Sum of Sales |

This configuration displays total sales for each product category within each region. Additional PivotTables can summarize sales by month, customer, or other available fields.

**Objective:** Compare sales performance across multiple dimensions and identify differences between categories.

---

### 9. Data Visualization

Charts help communicate analytical results clearly.

| Chart Type | Purpose |
|---|---|
| Column Chart | Compare sales across regions or product categories |
| Bar Chart | Rank categories or regions by sales |
| Line Chart | Show monthly sales trends |
| Pie Chart | Illustrate category contributions to total sales |
| Column Chart (What-If) | Compare modeled results under different discount scenarios |

Charts should include meaningful titles, readable labels, suitable number formats, and appropriate legends.

**Objective:** Convert calculated results into visual summaries that are easier to understand.

---

### 10. Dashboard and Key Performance Indicators

The Dashboard worksheet brings together key metrics and visualizations to provide an overview of business performance.

| KPI | Formula | Description |
|---|---|---|
| Total Sales | `=SUM(Sales_Data!E2:E201)` | Sum of the sales amounts in the specified range |
| Total Profit | `=SUM(Sales_Data!I2:I201)` | Sum of the profit values |
| Transaction Records | `=COUNTA(Sales_Data!A2:A201)` | Counts non-empty `Customer_ID` cells; gives the number of transaction records, not unique orders or customers |
| Average Sales per Record | `=AVERAGE(Sales_Data!E2:E201)` | Average sales amount for the selected records |

The dashboard can also incorporate charts for regional performance, product category comparisons, and monthly sales trends.

**Objective:** Provide a consolidated view of important sales and profit indicators.

---

## 🧮 Excel Functions Used

| Function or Feature | Application |
|---|---|
| `SUM` | Calculate total sales and profit |
| `SUMIF` | Calculate customer-level totals |
| `SUMIFS` | Calculate monthly sales using date criteria |
| `SUMPRODUCT` | Evaluate calculations across multiple ranges |
| `IF` | Apply logical conditions |
| `IFERROR` | Handle formula errors |
| `RANK.EQ` | Rank customers by purchase totals |
| `INDEX` | Retrieve a value from a specified position |
| `MATCH` | Find the position of a matching value |
| `EDATE` | Calculate dates at monthly intervals |
| `DATE` | Create date values |
| `NOW` | Return the current date and time |
| Conditional Formatting | Highlight important results |
| PivotTables | Summarize data |
| Data Analysis ToolPak | Perform statistical analysis |

---

## 🔄 Project Workflow

1. Prepare and format the sales dataset.
2. Organize the workbook into analytical worksheets.
3. Calculate customer-level sales totals.
4. Identify high-value customers using rankings.
5. Evaluate hypothetical discount scenarios.
6. Perform regression analysis on Sales and Profit.
7. Generate descriptive statistics.
8. Calculate monthly sales totals and growth.
9. Summarize regional and product-category sales with PivotTables.
10. Create charts to visualize the results.
11. Consolidate important metrics and visualizations in the Dashboard.
12. Review formulas, number formats, and calculated outputs.

---

## 💼 Business Applications

- **Sales Performance Monitoring:** Track sales totals and compare performance across periods.
- **Customer Analysis:** Identify customers with higher purchase totals.
- **Regional Analysis:** Compare performance across geographical regions.
- **Product Category Analysis:** Evaluate the contribution of different product categories.
- **Profitability Analysis:** Summarize profit and explore its relationship with sales.
- **Scenario Evaluation:** Compare hypothetical discount assumptions.
- **Statistical Analysis:** Summarize data distributions and investigate relationships.
- **Business Reporting:** Present key metrics and charts in a consolidated dashboard.

These applications demonstrate how spreadsheet analysis can support business reporting and decision-making. Actual business decisions should also consider data quality, costs, business rules, and other relevant factors.

---

## 🚀 How to Use the Project

1. Clone or download the repository:
   ```bash
   git clone https://github.com/pranamiraj083-ship-it/excel-analyzer.git
   ```
2. Open the Excel workbook included in the repository.
3. Review the `Sales_Data` worksheet to understand the source dataset.
4. Explore the customer analysis and conditional formatting results.
5. Review the What-If Analysis scenarios.
6. Examine the regression and descriptive statistics outputs.
7. Check the monthly sales calculations and trend indicators.
8. Explore the PivotTables and charts.
9. Open the Dashboard to review the consolidated results.

> Some statistical features require Microsoft Excel's **Data Analysis ToolPak**.

---

## 🎓 Learning Outcomes

This project provides practical experience in:

- Organizing and formatting structured data in Excel
- Writing formulas for conditional calculations and aggregations
- Using lookup and ranking functions
- Analyzing customer purchasing behavior
- Calculating monthly sales growth
- Applying Conditional Formatting and visual indicators
- Performing regression and descriptive statistical analysis
- Creating and interpreting PivotTables
- Building charts and KPI-based dashboards
- Presenting numerical results in a business-reporting format

---

## ⚠️ Project Limitations

- The analysis depends on the completeness and accuracy of the supplied dataset.
- Regression results describe the fitted relationship within the analyzed data and should not automatically be interpreted as proof of causation.
- What-If Analysis results depend on the assumptions and formulas used in each scenario.
- Monthly growth calculations require careful handling of zero or missing previous-month values.
- Dashboard values are only as reliable as the formulas and source data used to calculate them.
- Results should be validated before being used for real business decisions.

---

## 🔮 Future Enhancements

- Adding interactive slicers to PivotTables and dashboard charts
- Expanding customer segmentation and regional comparisons
- Adding more detailed product-level profitability analysis
- Improving the presentation of statistical results
- Adding data validation and automated data-quality checks
- Creating additional visualizations for monthly and category-level performance
- Updating the dashboard to support more flexible reporting

---

## ✅ Conclusion

The **PR. 2 Analyzer – Sales and Profit Analytics Project** demonstrates how Microsoft Excel can be used to perform structured data analysis, statistical evaluation, customer analysis, scenario modeling, and business reporting.

By combining formulas, Conditional Formatting, regression, descriptive statistics, PivotTables, charts, and dashboard techniques, the project turns transactional sales data into organized analytical outputs.

It serves as a practical learning project for developing spreadsheet analysis skills and understanding how Excel can support data-driven business reporting.

---

## 👩‍💻 Author

**rajkumar**

GitHub: [pranamiraj083-ship-it](https://github.com/pranamiraj083-ship-it)

---

## 📦 Repository

**Excel-Analyzer** → [https://github.com/pranamiraj083-ship-it/excel-analyzer](https://github.com/pranamiraj083-ship-it/excel-analyzer)

---

*This project was developed for educational and practical skill-building purposes using Microsoft Excel.*
