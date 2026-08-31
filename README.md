# Property Valuation And Market Trends 
Cleaning a messy 56,000-row house sales dataset using SQL Server, then linking the final cleaned view to an interactive Excel Dashboard

### Tools Used
* **Database:** SQL Server (T-SQL)
* **Dashboard:**  EXCEL Pivot Charts including Slicers and Timelines + KPIs


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

### 3. Splitting Address Components (String Math)
* **What it does:** Breaks a single, squashed address string into clean, separate columns for the street name and the city. It uses string functions to find the comma delimiter, slices the text at that exact boundary, and updates the new fields.
```sql
-- Breaking ADDRESS into Individual Coloumns (Address, City, State) Function --  the delimiter is a comma


ALTER  TABLE NASHVILLEHOUSING
ADD CITY varchar(255),
SPLITADDRESS VARCHAR(255);
UPDATE nashvillehousing
SET CITY = SUBSTRING(PropertyAddress, CharINDEX(',', PropertyAddress) +1, Len(PropertyAddress)), 
SPLITADDRESS = SUBSTRING(PropertyAddress, 1, CharINDEX(',', PropertyAddress) -1)
```
### 4. Splitting Address Columns (Alternative Method)
* **What it does:** Uses string matching functions to cut text apart without relying on normal commas or periods.

```sql
--Split OwnerAddress into 3 coloumns (address, City, State)

Select OwnerAddress
from [PROJECT1NASHVILLEDATACLEANING].dbo.nashvillehousing

Select ParseName(Replace(OwnerAddress,',','.'),3) as ownerSplitAddress,
parsename(Replace(OwnerAddress, ',','.'), 2) as ownerSplitCity,
parsename(replace(owneraddress, ',', '.'), 1) as ownderSplitState
from [PROJECT1NASHVILLEDATACLEANING].dbo.nashvillehousing

begin transaction 
alter table nashvillehousing
add OwnerSplitAddress varchar(255),
OwnerSplitCity varchar(255),
OwnerSplitState varchar(255)

commit transaction

Update nashvillehousing
SET OwnerSplitAddress = ParseName(Replace(OwnerAddress,',','.'),3),
OwnerSplitCity = parsename(Replace(OwnerAddress, ',','.'), 2),
OwnerSplitState = parsename(replace(owneraddress, ',', '.'), 1)
```
### 5. Fixing Yes/No Typos
* **What it does:** Finds messy data entry columns (like "Y", "N", "Yes", "No") and changes them all to match perfectly.
```sql
-- Change Y to Yes and N to No in SoldAsVacant


select distinct (SoldAsVacant), Count(SoldAsVacant) as NumberOfRows
from [PROJECT1NASHVILLEDATACLEANING].dbo.nashvillehousing 
Group by SoldAsVacant
order by NumberOfRows

Select distinct(S.A)
From
(
Select case
when SoldAsVacant = 'Y' then 'Yes'
when SoldAsVacant = 'N' then 'No'
Else SoldAsVacant
End  as A
from [PROJECT1NASHVILLEDATACLEANING].dbo.nashvillehousing)  as S

Select case
when SoldAsVacant = 'Y' then 'Yes'
when SoldAsVacant = 'N' then 'No'
Else SoldAsVacant
End 
from [PROJECT1NASHVILLEDATACLEANING].dbo.nashvillehousing

begin transaction
ALTER TABLE Nashvillehousing
ADD NewSoldAsVacant varchar(255)
commit transaction


UPDATE nashvillehousing
SET NewSoldAsVacant = case
when SoldAsVacant = 'Y' then 'Yes'
when SoldAsVacant = 'N' then 'No'
Else SoldAsVacant
End 

select *
from [PROJECT1NASHVILLEDATACLEANING].dbo.nashvillehousing
```
### 6. Removing Duplicates Using Row Numbers
* **What it does:** Finds and deletes duplicate entries where identical property details were logged multiple times. It assigns unique row numbers inside matching data groups and uses a CTE loop to permanently remove the extra rows from the physical table.

```sql
-- REMOVING duplicates using row numbers!



WITH ROWSS AS (
Select *, 
Row_Number() over (partition by ParcelID, PropertySplitAddress, SalePrice, SalesDateConverted, Legalreference order by uniqueid) as RNN
from [PROJECT1NASHVILLEDATACLEANING].dbo.nashvillehousing
)
,
FINAL AS (
SELECT *
FROM ROWSS
WHERE RNN > 1) 


DELETE 
FROM FINAL
WHERE RNN > 1
```
### 7. Removing Unused Columns
* **What it does:** Drops the old, uncorrected, or unparsed columns that are no longer needed for reporting. This cleans up the table layout and optimizes database storage.

```sql
--- deleting unused coloumns that we wont use in analysis (all the uncorrected  columns)
-- usually do not delete any coloumns from permenant data
-- 

ALTER TABLE nashvillehousing
DROP  COLUMN  SaleDate, PropertyAddress, OwnerAddress, OwnerName, legalreference, taxdistrict

select *
from [PROJECT1NASHVILLEDATACLEANING].dbo.nashvillehousing
```
### 8. Creating the Final View
* **What it does:** Saves the entire clean dataset as a shortcut view so Excel can read it live from the server.
```sql
USE PROJECT1NASHVILLEDATACLEANING
go

DROP VIEW IF EXISTS V_CLEANEDHOUSINGDATA
GO

CREATE VIEW V_CLEANEDHOUSINGDATA
AS
SELECT *
FROM [PROJECT1NASHVILLEDATACLEANING].DBO.nashvillehousing

GO

SELECT *
FROM V_CLEANEDHOUSINGDATA
```
## Part 2: Property Valuation & Market Trends Dashboard (Real Estate Analytics Layer)

This section shows the interactive real estate dashboard built inside Microsoft Excel. It turns thousands of messy housing rows into a clean, easy-to-read screen that helps managers see price trends, neighbor values, and building age patterns instantly.

###  Section 2a: Top Summary Numbers (KPI Cards)
* **KPI 1 - Total Homes Sold:** Counts every property sale tracked in the data to show the complete size of the housing market.
* **KPI 2 - Record Breaking Sale Price:** Finds the single most expensive property ever sold to highlight the maximum market ceiling.
* **KPI 3 - Average Sale Price:** Calculates the middle-ground, standard value of a home to show what a normal house costs.
* **KPI 4 - Number of New Developments:** Counts exactly how many brand-new properties were built and sold in recent years.

---

### Section 2b: The Main Charts
* **Chart 1 - Home Sales Over Time (Combo Chart):** Tracks monthly home sales and average prices at the same time on a timeline to show summer spikes and winter dips.
* **Chart 2 - Property Types Per Neighbourhood (100% Stacked Chart):** Shows the mix of houses (like Single Family vs. Duplex) in each town using percentages so you can compare tiny zones against big cities easily.
* **Chart 3 - Average Sale Price Per Neighbourhood (Horizontal Bar):** Ranks the priciest towns from top to bottom so real estate investors get a quick cheat sheet of expensive zones.
* **Chart 4 - Average Sale Price Vs Year Built (Column Chart):** Groups homes by the decade they were built to prove whether buyers pay a premium for old historic charm or brand-new modern builds.
