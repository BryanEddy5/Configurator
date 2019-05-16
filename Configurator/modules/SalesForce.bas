Option Compare Database
Option Explicit

Public Function CalculateDaysOpen( _
  ByVal datDateFrom As Date, _
  ByVal datDateTo As Date, _
  Optional ByVal booExcludeHolidays As Boolean) _
  As Double
'#39; Purpose: Calculate number of working days between dates datDateFrom and datDateTo.
' Assumes: 5 or 6 working days per week. Weekend is (Saturday and) Sunday.
' May be freely used and distributed.
' 1999-04-23. Gustav Brock, Cactus Data ApS, Copenhagen
' 2000-10-03. Constants added.
'             Option for 5 or 6 working days per week added.
' 2008-06-12. Option to exclude holidays from the count of workdays.
  Const cbytWorkdaysOfWeek  As Byte = 5
  ' Name of table with holidays.
  Const cstrTableHoliday    As String = "tblHoliday"
  ' Name of date field in holiday table.
  Const cstrFieldHoliday    As String = "HolidayDate"
  Dim bytSunday             As Byte
  Dim intWeekdayDateFrom    As Double
  Dim intWeekdayDateTo      As Double
  Dim lngDays               As Double
  Dim datDateTemp           As Date
  Dim strDateFrom           As String
  Dim strDateTo             As String
  Dim lngHolidays           As Double
  Dim strFilter             As String
  
  ' Reverse dates if these have been input reversed.
  If datDateFrom > datDateTo Then
    datDateTemp = datDateFrom
    datDateFrom = datDateTo
    datDateTo = datDateTemp
  End If
  
  ' Find ISO weekday for Sunday.
  bytSunday = Weekday(vbSunday, vbMonday)
  
  ' Find weekdays for the dates.
  intWeekdayDateFrom = Weekday(datDateFrom, vbMonday) + Hour(datDateFrom) / 24 + Minute(datDateFrom) / 24 / 60
  intWeekdayDateTo = Weekday(datDateTo, vbMonday) + Hour(datDateTo) / 24 + Minute(datDateTo) / 24 / 60
  
  ' Compensate weekdays' value for non-working days (weekends).
  intWeekdayDateFrom = intWeekdayDateFrom + (intWeekdayDateFrom = bytSunday)
  intWeekdayDateTo = intWeekdayDateTo + (intWeekdayDateTo = bytSunday)
  
  ' Calculate number of working days between the two weekdays, ignoring number of weeks.
  lngDays = intWeekdayDateTo - intWeekdayDateFrom - (cbytWorkdaysOfWeek * (intWeekdayDateTo < intWeekdayDateFrom))
  
    Select Case Weekday(datDateTo)
        Case vbSaturday
            lngDays = lngDays - 1
        Case vbSunday
            lngDays = lngDays - 2
    End Select
  ' Add number of working days between the weeks of the two dates.
  lngDays = lngDays + (cbytWorkdaysOfWeek * DateDiff("w", datDateFrom, datDateTo, vbMonday, vbFirstFourDays))
  If booExcludeHolidays And lngDays > 0 Then
    strDateFrom = Format(datDateFrom, "yyyy\/mm\/dd")
    strDateTo = Format(datDateTo, "yyyy\/mm\/dd")
    strFilter = cstrFieldHoliday & " Between #" & strDateFrom & "# And #" & strDateTo & "# And Weekday(" & cstrFieldHoliday & ", 2) <= " & cbytWorkdaysOfWeek & ""
    lngHolidays = DCount("*", cstrTableHoliday, strFilter)
  End If
  
  CalculateDaysOpen = Round(lngDays - lngHolidays, 3)
End Function