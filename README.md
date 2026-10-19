# 📊 Indian E-Commerce Data Analysis & Interactive Excel Dashboard

## 📌 Project Overview
This project demonstrates an end-to-end data analysis workflow using Microsoft Excel on an Indian e-commerce transaction dataset. The analysis covers the full pipeline from raw data extraction, cleaning messy field entries (currency formats, mixed date structures, inconsistent text), formula-based data transformations, and Pivot Table aggregation to building an interactive visual dashboard.

---

## 🖼️ Dashboard Preview

![Dashboard Preview](dashboard_preview.png)

---

## 🛠️ Key Technical Steps & Methodology

### 1. Data Cleaning & Standardization (Power Query & Formulas)
* **Text Formatting:** Cleaned whitespace and standard casing for Customer Names and Product Categories using `=PROPER(TRIM())`.
* **Price Standardization:** Removed currency symbols (`₹`, `$`, `Rs`) and converted raw text to numeric values using `=VALUE(SUBSTITUTE())`.
* **Date Normalization:** Fixed mixed date formats (`YYYY/MM/DD`, `DD-MM-YYYY`, `YYYY.MM.DD`) into standard Excel short date structures (`DD-MM-YYYY`).

### 2. Calculated Columns & Logic
* **Gross Revenue:** `=Clean_Unit_Price * Quantity`
* **Discount Amount:** `=Gross_Revenue * Discount_Pct`
* **Net Revenue:** `=Gross_Revenue - Discount_Amount`
* **Delivered Revenue Flag:** `=IF(Order_Status="Delivered", Net_Revenue, 0)` to ensure returned/cancelled orders do not inflate total sales figures.
* **High-Value Order Flag:** `=IF(AND(Net_Revenue>20000, Order_Status="Delivered"), "High Value Delivered", "Standard")`

### 3. Summary Aggregations & Lookups
* Applied `SUMIFS` to calculate total revenue per region for delivered orders.
* Applied `COUNTIFS` to evaluate order fulfillment breakdown across statuses (`Delivered`, `Returned`, `Cancelled`).
* Built dynamic lookup searches using `XLOOKUP`.

### 4. Interactive Dashboard Assembly
* **KPI Summary Cards:** Linked dynamically to Pivot Table aggregations to display Total Revenue (₹), Total Orders, Average Order Value (AOV), and Total Quantity Sold.
* **Visual Charts:** 
  * *Monthly Revenue Trend* (Line Chart with Markers)
  * *Category Performance* (Clustered Bar Chart)
  * *Order Fulfillment Distribution* (Donut Chart)
* **Dynamic Slicers:** Connected `Region`, `Payment_Method`, and `City` slicers across all Pivot Tables for instant multi-dimensional filtering.

---

## 💡 Key Business Insights

1. **Top Revenue Driver:** Electronics generates the highest net sales volume among all categories.
2. **Payment Preferences:** Digital payment modes (UPI and Net Banking) constitute the majority of successfully delivered transactions compared to Cash on Delivery (COD).
3. **Fulfillment Tracking:** Return and cancellation tracking provides visibility into net actualized revenue versus gross order volume.

---

## 📁 Repository Structure
