# real-estate-etl-pipeline
Cleaning a messy 56,000-row house sales dataset using SQL Server, then linking the final cleaned view to an interactive Excel Power Pivot dashboard.

### Tools Used
* **Database:** SQL Server (T-SQL)
* **Dashboard:** Excel Power Pivot, Data Modeling, and DAX formulas


## Part 1: Cleaning the Data with SQL

The raw data had a lot of mistakes, missing info, and formatting bugs. These 10 scripts show how I fixed the tables step-by-step without losing any important data.

### 1. Fixing Date Formats
* **What it does:** Converts messy long date-time stamps into a clean, simple date format.
``` sql
-- Standardise Date Format

Alter TABLE nashvillehousing
add SalesDateConverted Date;

UPDATE [PROJECT1NASHVILLEDATACLEANING].DBO.nashvillehousing 
SET SalesDateConverted = Convert(Date, SaleDate)


```
