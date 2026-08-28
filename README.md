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
### 2 .Filling Missing Property Addresses (Self-JOIN)
* **What it does:** Finds records where the property address was left completely blank. It looks for other rows with the exact same Parcel ID but a different Unique ID, and uses that matching data to automatically fill in the missing address.
```sql
-- Select pass to verify the missing data and its replacement
Select isnull(nvh1.propertyaddress, nvh2.PropertyAddress) as CorrectedPropertyAddress
from [PROJECT1NASHVILLEDATACLEANING].dbo.nashvillehousing as nvh1
join [PROJECT1NASHVILLEDATACLEANING].dbo.nashvillehousing as nvh2
on nvh1.ParcelID = nvh2.ParcelID
where nvh1.PropertyAddress is null
and nvh1.[UniqueID ] <> nvh2.[UniqueID ]

-- Active update pass to permanently fix the blank fields
UPDATE nvh1
SET PropertyAddress = isnull(nvh1.propertyaddress, nvh2.PropertyAddress)
from [PROJECT1NASHVILLEDATACLEANING].dbo.nashvillehousing as nvh1
join [PROJECT1NASHVILLEDATACLEANING].dbo.nashvillehousing as nvh2
on nvh1.ParcelID = nvh2.ParcelID
where nvh1.PropertyAddress is null
and nvh1.[UniqueID ] <> nvh2.[UniqueID ]
```
