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
