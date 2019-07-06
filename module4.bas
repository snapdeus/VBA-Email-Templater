Public Function NetWorkdays2(StartDate As Date, EndDate As Date, _
    ExcludeDaysOfWeek As Long, Optional Holidays As Variant) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NetWorkdays2
' This function calcluates the number of days between StartDate and EndDate
' excluding those days of the week specified by ExcludeDaysOfWeek and
' optionally excluding dates in Holidays. ExcludeDaysOfWeek is a
' value from the table below.
'       1  = Sunday     = 2 ^ (vbSunday - 1)
'       2  = Monday     = 2 ^ (vbMonday - 1)
'       4  = Tuesday    = 2 ^ (vbTuesday - 1)
'       8  = Wednesday  = 2 ^ (vbWednesday - 1)
'       16 = Thursday   = 2 ^ (vbThursday - 1)
'       32 = Friday     = 2 ^ (vbFriday - 1)
'       64 = Saturday   = 2 ^ (vbSaturday - 1)
' To exclude multiple days, add the values in the table together. For example,
' to exclude Mondays and Wednesdays, set ExcludeDaysOfWeek to 10 = 8 + 2 =
' Monday + Wednesday.
' If StartDate is less than or equal to EndDate, the result is positive. If
' StartDate is greater than EndDate, the result is negative. If either
' StartDate or EndDate is less than or equal to 0, the result is a
' #NUM error. If ExcludeDaysOfWeek is less than 0 or greater than or
' equal to 127 (all days excluded), the result is a #NUM error.
' Holidays is optional and may be a single constant value, an array of values,
' or a worksheet range of cells.
' This function can be used as a replacement for the NETWORKDAYS worksheet
' function. With NETWORKDAYS, the excluded days of week are hard coded
' as Saturday and Sunday. You cannot exlcude other days of the week. This
' function allows you to exclude any number of days of the week (with the
' exception of excluding all days of week), from 0 to 6 days. If
' ExcludeDaysOfWeek = 65 (Sunday + Saturday), the result is the same as
' NETWORKDAYS.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
Dim TestDayOfWeek As Long
Dim TestDate As Date
Dim Count As Long
Dim Stp As Long
Dim Holiday As Variant
Dim Exclude As Boolean

If ExcludeDaysOfWeek < 0 Or ExcludeDaysOfWeek >= 127 Then
    ' invalid value for ExcludeDaysOfWeek. get out with error.
    NetWorkdays2 = CVErr(xlErrNum)
    Exit Function
End If

If StartDate <= 0 Or EndDate <= 0 Then
    ' invalid date. get out with error.
    NetWorkdays2 = CVErr(xlErrNum)
    Exit Function
End If

' set the value used for the Step in
' the For loop.
If StartDate <= EndDate Then
    Stp = 1
Else
    Stp = -1
End If

For TestDate = StartDate To EndDate Step Stp
    ' get the bit pattern of the weekday of TestDate
    TestDayOfWeek = 2 ^ (Weekday(TestDate, vbSunday) - 1)
    If (TestDayOfWeek And ExcludeDaysOfWeek) = 0 Then
        ' do not exclude this day of week
        If IsMissing(Holidays) = True Then
            ' count day
            Count = Count + 1
        Else
            Exclude = False
            ' holidays provided. test date for holiday.
            If IsObject(Holidays) = True Then
                ' assume Excel.Range
                For Each Holiday In Holidays
                    If Holiday.Value = TestDate Then
                        Exclude = True
                        Exit For
                    End If
                Next Holiday
            Else
                ' not an Excel.Range
                If IsArray(Holidays) = True Then
                    For Each Holiday In Holidays
                        If Int(Holiday) = TestDate Then
                            Exclude = True
                            Exit For
                        End If
                    Next Holiday
                Else
                    ' not an array or range, assume single value
                    If TestDate = Holidays Then
                        Exclude = True
                    End If
                End If
            End If
            If Exclude = False Then
                Count = Count + 1
            End If
        End If
    Else
        ' excluded day of week. do nothing
    End If
Next TestDate
' return the result, positive or negative based on Stp.
NetWorkdays2 = (Count * Stp) - 1

End Function


'Credit to: http://www.cpearson.com/excel/betternetworkdays.aspx

