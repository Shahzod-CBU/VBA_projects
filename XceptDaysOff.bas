Attribute VB_Name = "XceptDaysOff"
Option Explicit

Private Function ExceptHolidays(a As Date) As Date
Dim y As Integer
y = Year(a)
If a = DateSerial(y, 1, 1) Or a = DateSerial(y, 3, 8) Or a = DateSerial(y, 3, 21) Or a = DateSerial(y, 5, 9) Or _
            a = DateSerial(y, 9, 1) Or a = DateSerial(y, 10, 1) Or a = DateSerial(y, 12, 8) Then
    ExceptHolidays = a - 1
Else
    ExceptHolidays = a
End If
End Function

Private Function ExceptWeekends(b As Date) As Date
Dim kun As Date
kun = LastSaturday(b)

If b = kun Then
    ExceptWeekends = b
    Exit Function
End If

If Weekday(b, vbMonday) = 6 Then
    ExceptWeekends = b - 1
ElseIf Weekday(b, vbMonday) = 7 Then
    ExceptWeekends = b - 2
Else
    ExceptWeekends = b
End If
End Function

Function ExceptDaysOff(c As Date) As Date
ExceptDaysOff = ExceptWeekends(ExceptHolidays(c))
End Function


Private Function ExceptHolidaysPlus(a As Date) As Date
Dim y As Integer
y = Year(a)
If a = DateSerial(y, 1, 1) Or a = DateSerial(y, 3, 8) Or a = DateSerial(y, 3, 21) Or a = DateSerial(y, 5, 9) Or _
            a = DateSerial(y, 9, 1) Or a = DateSerial(y, 10, 1) Or a = DateSerial(y, 12, 8) Then
    ExceptHolidaysPlus = a + 1
Else
    ExceptHolidaysPlus = a
End If
End Function

Private Function ExceptWeekendsPlus(b As Date) As Date

Dim kun As Date
kun = LastSaturday(b)

If b = kun Then
    ExceptWeekendsPlus = b
    Exit Function
End If

If Weekday(b, vbMonday) = 6 Then
    ExceptWeekendsPlus = b + 2
ElseIf Weekday(b, vbMonday) = 7 Then
    ExceptWeekendsPlus = b + 1
Else
    ExceptWeekendsPlus = b
End If
End Function

Function ExceptDaysOffPlus(c As Date) As Date
ExceptDaysOffPlus = ExceptWeekendsPlus(ExceptHolidaysPlus(c))
End Function

Private Function LastSaturday(Sana As Date) As Date

    Dim LastDayOfMonth As Date, a As Date, b As Integer
    
    LastDayOfMonth = DateSerial(Year(Sana), Month(Sana) + 1, 0)
    a = LastDayOfMonth - 7
    b = 6 - Weekday(a, vbMonday)
    
    If Weekday(LastDayOfMonth, vbMonday) = 6 Then
        LastSaturday = LastDayOfMonth
    ElseIf b > 0 Then
        LastSaturday = a + b
    Else
        LastSaturday = LastDayOfMonth + b
    End If

End Function

