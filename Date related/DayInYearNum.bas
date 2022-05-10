Attribute VB_Name = "DayInYearNum"
Option Explicit
Option Base 1

Function DayNumInYear()
'This function returns the day number in the year based on the
'   current date ie system time
'   (ie how many days have passed in the year so far)

'Define variables to be used
Dim YearNum As Integer, MonthNum As Integer, DayInMonth As Integer
Dim MonthIndex As Integer
Dim TempDayNum As Integer
Dim MaxDaysInMonth As Variant

'Set required variables
YearNum = Year(Date)
MonthNum = Month(Date)
DayInMonth = Day(Date)

'Adjust MaxDaysInMonth array to accomodate leap years. These values are added
'   to find the final day number.
'   Note: Performance could be improved by using the total day count in the
'       MaxDaysInMonth arrays rather than looping through but it is left
'       as is for clarity.
If (YearNum - 1900 Mod 4) = 0 Then
    MaxDaysInMonth = Array(31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
ElseIf (YearNum - 1900 Mod 4) <> 0 Then
    MaxDaysInMonth = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
ElseIf (YearNum - 1900 Mod 100) = 0 Then
    MaxDaysInMonth = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
End If

TempDayNum = 0 'Initialise variable used to store summed days

'Add the number of days in each of the previous months in the year
For MonthIndex = 1 To MonthNum - 1
    TempDayNum = TempDayNum + MaxDaysInMonth(MonthIndex)
Next MonthIndex
'Add the days passed in the current month
TempDayNum = TempDayNum + DayInMonth

'Return result
DayNumInYear = TempDayNum

End Function

