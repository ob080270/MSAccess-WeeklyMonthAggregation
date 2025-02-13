# MSAccess-WeeklyMonthAggregation
VBA function for determining the month to which the week belongs, taking into account the dominant days. It is used in MS Access.
# 📊 GetMonthForWeek Function for MS Access

## 📌 Overview
This VBA function **determines the dominant month for a given ISO week** in a specific year.  
It is designed for use in **Microsoft Access** SQL queries to **group weekly data by month**.

## 🔹 How It Works
- Takes a **year** and **week number (1-53)** as input.
- Optionally, allows setting the **start day of the week** (`Monday = 2` by default).
- **Returns the month number (1-12)** where most days of the week belong.

---

## **📌 Function Code (VBA)**
```vba
Public Function GetMonthForWeek(targetYear As Integer, targetWeek As Integer, Optional startDayOfWeek As Integer = 2) As Integer
' This function determines the month to which most of the given week belongs.
' 
' Arguments:
'   targetYear      - The year for which the calculation is performed.
'   targetWeek      - The ISO week number (1-53).
'   startDayOfWeek  - The first day of the week (1 = Sunday, 2 = Monday, default = 2).
'
' Returns:
'   The month number (1-12) to which most of the week belongs.
'
' Example:
'   GetMonthForWeek(2024, 5)  ' Returns 2 (February)
'
'-------------------------------------------------------------------------------------------------------------------

    ' Validate input values
    If targetYear < 1900 Or targetYear > 2100 Then Err.Raise 5, "GetMonthForWeek", "Invalid year"
    If targetWeek < 1 Or targetWeek > 53 Then Err.Raise 5, "GetMonthForWeek", "Invalid week number"
    If startDayOfWeek <> 1 And startDayOfWeek <> 2 Then Err.Raise 5, "GetMonthForWeek", "Invalid start day (must be 1 or 2)"

    ' Calculate the start date of the given week
    Dim startDate As Date
    Dim midWeekDate As Date
    
    ' Determine the first day of the requested week
    startDate = DateAdd("ww", targetWeek - 1, DateSerial(targetYear, 1, 1))
    startDate = DateAdd("d", -Weekday(startDate, startDayOfWeek) + 1, startDate)
    
    ' Find the midpoint of the week (used for determining the dominant month)
    midWeekDate = DateAdd("d", 3, startDate)

    ' Return the month of the midpoint date
    GetMonthForWeek = Month(midWeekDate)

End Function

📌 SQL Query for MS Access
This SQL query groups data by week and assigns the corresponding month using the GetMonthForWeek function:
SELECT 
    Year([smpDate]) AS smpYear, 
    GetMonthForWeek(Year([smpDate]), DatePart("ww", [smpDate], 0)) AS smpMonth, 
    DatePart("ww", [smpDate], 0) AS smpWeek, 

    Sum([smpValue1]) AS smpSumA, 
    Sum([smpValue2]) AS smpSumB

FROM smpTable

GROUP BY 
    Year([smpDate]), 
    GetMonthForWeek(Year([smpDate]), DatePart("ww", [smpDate], 0)), 
    DatePart("ww", [smpDate], 0);

📌 Parameters:
Parameter	Description
targetYear	The year of the calculation (default range: 1900-2100)
targetWeek	The ISO week number (1-53)
startDayOfWeek	First day of the week: 1 = Sunday, 2 = Monday (default = 2)
smpDate	The date field used in the SQL query
smpValue1	Aggregated value field 1
smpValue2	Aggregated value field 2

📌 Example Usage

1️⃣ Calling the Function in VBA:
Dim result As Integer
result = GetMonthForWeek(2024, 5) ' Returns 2 (February)
MsgBox "The dominant month is: " & result

2️⃣ Using in an MS Access SQL Query:
SELECT smpYear, smpMonth, smpWeek, smpSumA, smpSumB
FROM qAggregatedWeeks
ORDER BY smpYear DESC, smpWeek ASC;
