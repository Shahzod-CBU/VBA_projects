Attribute VB_Name = "LiquidityContinued"
Option Explicit
'Option Private Module

'Public Const stPath$ = "C:\Users\Salixov_M\Desktop\Тест ликвид"
'Const Shablon$ = "C:\Users\Salixov_M\Desktop\корсчет керак\Шаблон корсчет.xlsx"    'Kun ochish u-n shablon fayl
'Const downloadsPath$ = "C:\Users\Salixov_M\Downloads"

Public Const stPath$ = "D:"                     'fayllar saqlanadigan papka
Const Shablon$ = "c:\Users\Asus\Documents\корсчет керак\Шаблон корсчет.xlsx"    'Kun ochish u-n shablon fayl
Const downloadsPath$ = "c:\Users\Asus\Downloads"

Sub Корсчет_давоми()

'25.03.2019 da Shahzod tomonidan yozildi
'10.04.2019 da yangilandi (faktorlarni mos bankka qo'yish)
'last updated on 10.04.2019 at 16:00

Dim MyPivot As PivotTable, Start As Date, Finish As Date, i&
Dim wbJamlovchi As Workbook, wbManba As Workbook, Bugun As String
Dim PivotRangeArr(), FactorsInHand(), Jami As Long, Boshqa As Long
Dim GivenFactors As Range, FactorsByBanks(), wbShablon As Workbook
Dim FaktorlarTasiri(), FaktorlarNomi(), Buyi&, Eni&, m&, n&, x&, y&
Dim Boshlanishi&, tdate As Date, manbaPath$, nameActWb$
Dim nameManbaWb$, wbJamlovchiName$, listsNum%, MsgTxt$
Dim qoldiqlar$, qoldiqlarRng As Range, TempArr(), a&, b&

Rem Qaysi kun uchun korschot qilinishini aniqlab olamiz
nameActWb = ActiveWorkbook.Name
If InStr(nameActWb, "Корсчет фактор") <> 0 Then
    tdate = Mid(nameActWb, 16, 10)
Else
    tdate = Application.InputBox("Sanani ko'rsating", "Korschot davomi", Format(Now, "dd.mm.yyyy"))
    If tdate = 0 Then Exit Sub
End If

On Error Resume Next

qoldiqlar = "Remainders_" & tdate

If Dir(downloadsPath & "\" & qoldiqlar & ".csv") = "" Then
    MsgTxt = "Kun yakunlari bo'yicha oborotlar saqlangan fayl topilmadi!" & Chr(10) & Chr(10)
    MsgTxt = MsgTxt & "'Центр расчетов'ning 'Остатки на коррсчетах (таблица)' qismidan " & Chr(10)
    MsgTxt = MsgTxt & "'Импорт данных'ni bosib kunlik oborotlarni yuklab olib, " & Chr(10)
    MsgTxt = MsgTxt & "keyin makrosni qaytadan ishga tushiring."
    MsgBox MsgTxt, vbOKOnly + vbCritical, "Xatolik"
    Exit Sub
End If

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Start = Timer

Rem Saqlash uchun manzillarni topamiz
fldrYear = Format(tdate, "yyyy")
monthNum = Format(tdate, "mm")
fldrMonthName = Format(tdate, "MMMM")
NewFolder = "кунлик корр. счет " & fldrYear
fldrpath = stPath & "\" & NewFolder
fldrpathMonth = fldrpath & "\" & monthNum & " " & fldrMonthName
fldrpathResults = fldrpathMonth & "\" & fldrMonthName

Set wbShablon = Workbooks.Open(Shablon, False)      'shablon faylni ochamiz
wbJamlovchiName = fldrMonthName & ".xlsx"           'kunlik korschot yig'ilib boradigan fayl
If Not IsBookOpen(wbJamlovchiName) Then             'agar yopiq bo'lsa ochamiz, mavjud bo'lmasa yaratamiz
    Workbooks.Open fldrpathMonth & "\" & wbJamlovchiName, False
    If Err <> 0 Then
        wbShablon.SaveAs fldrpathMonth & "\" & wbJamlovchiName
        Set wbJamlovchi = Workbooks(wbJamlovchiName)
        On Error GoTo 0
    Else
        Set wbJamlovchi = Workbooks(wbJamlovchiName)
        listsNum = wbJamlovchi.Worksheets.count
        wbShablon.Sheets("Шаблон").Copy After:=wbJamlovchi.Sheets(listsNum)
        wbShablon.Close False
    End If
Else
    Set wbJamlovchi = Workbooks(wbJamlovchiName)
    listsNum = wbJamlovchi.Worksheets.count
    wbShablon.Sheets("Шаблон").Copy After:=wbJamlovchi.Sheets(listsNum)
    wbShablon.Close False
End If

Bugun = Format(tdate, "dd")
wbJamlovchi.Sheets("Шаблон").Name = Bugun

Rem Ta'sir qilgan faktorlar yozilgan kitobni ochamiz
nameManbaWb = "Корсчет фактор " & tdate & ".xlsx"
If Not IsBookOpen(nameManbaWb) Then
    manbaPath = fldrpathResults & "\" & nameManbaWb
    Workbooks.Open manbaPath, False
End If

Set wbManba = Workbooks(nameManbaWb)
Set MyPivot = wbManba.Worksheets("PivotNet").PivotTables("PivotNet")
PivotRangeArr = MyPivot.TableRange1                                     '1) faqat faktorlarni olib olamiz

MyPivot.PivotFields("Банк").Orientation = xlColumnField
FactorsByBanks = MyPivot.TableRange1                                    '2)faktorlarni banklari bilan olamiz

On Error Resume Next
With wbJamlovchi.Sheets(Bugun)
    .Cells(2, 1).Value = tdate                                          'sarlavhaga sanani qo'yamiz
    Jami = .Cells.Find(What:="ЖАМИ", LookIn:=xlValues).Row              'mavjud faktorlarni olish uchun uning boshi va
    Boshqa = .Columns(8).Find(What:="Бош" & ChrW(1179) & "а", LookIn:=xlValues).Row     'oxirini aniqlab olamiz
    Boshlanishi = .Columns(1).Find(What:="1", LookIn:=xlValues).Row     'banklarga faktorlarni qo'yish uchun
    Set GivenFactors = .Range(.Cells(Jami, 8), .Cells(Boshqa - 1, 9))
    FactorsInHand = GivenFactors                                        'shablondan mavjud faktorlarni olamiz
    For i = 1 To UBound(FactorsInHand)                                  'mavjud faktorlarni qiymatini to'ldiramiz
        FactorsInHand(i, 2) = WorksheetFunction.VLookup(FactorsInHand(i, 1), PivotRangeArr, 2, 0)
        .Cells(Jami - 1 + i, 1).EntireRow.Hidden = False                'ta'sir qilmagan faktorlani skrit qilamiz
        If Err <> 0 Then
            FactorsInHand(i, 2) = ""                                     'Интервенция ni yozmaymiz (Anig'ini Operbankdan ko'rish k-k)
            If Not FactorsInHand(i, 1) Like "Интервенция" Then .Cells(Jami - 1 + i, 1).EntireRow.Hidden = True
            Err.Clear
        End If
    Next i
    GivenFactors.Value = FactorsInHand

    Buyi = UBound(FactorsByBanks)
    Eni = UBound(FactorsByBanks, 2)
    
    ReDim FaktorlarTasiri(1 To Eni, 1 To 2)
    Dim bankNames(), bankName As String
    bankNames = Application.Transpose(.Range(.Cells(Boshlanishi, 2), .Cells(Jami - 1, 2)).Value)
    
    Rem Faktorlarni banklar bo'yicha qo'yishga tayyorlab olamiz
    For m = 2 To Eni - 2                                'Jami va Markaziy bank uchun: -2
        bankName = Mid(FactorsByBanks(2, m), 5)         'faktorni mos bankka qo'yish uchun
        x = IndexOf(bankName, bankNames)
        For n = 3 To Buyi - 1                           'Jami uchun: -1
            y = Round(FactorsByBanks(n, m), 0)
            If Abs(y) >= 9 Then                         'Moduli 9dan katta bo'lgan faktorlarnigina yozamiz
                y = Format(y, "0")
                FaktorlarTasiri(x, 1) = IIf(IsEmpty(FaktorlarTasiri(x, 1)), FactorsByBanks(n, 1), FaktorlarTasiri(x, 1) & Chr(10) & FactorsByBanks(n, 1)) 'faktor nomi
                FaktorlarTasiri(x, 2) = IIf(IsEmpty(FaktorlarTasiri(x, 2)), y, FaktorlarTasiri(x, 2) & Chr(10) & y)                                    'faktor qiymati
            End If
        Next
    Next

    .Range(.Cells(Boshlanishi, 8), .Cells(Jami - 1, 9)).Value = FaktorlarTasiri                 'Natijani banklar bo'yicha qo'yamiz
    
    Set qoldiqlarRng = .Range(.Cells(Boshlanishi, 3), .Cells(Jami - 1, 6))
    With .QueryTables.Add(Connection:= _
        "TEXT;" & downloadsPath & "\" & qoldiqlar & ".csv", Destination _
        :=qoldiqlarRng)
        .CommandType = 0
        .Name = qoldiqlar
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(9, 9, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    wbJamlovchi.Connections(qoldiqlar).Delete
    
    TempArr = qoldiqlarRng
    For a = 1 To UBound(TempArr)
        For b = 1 To 4
            TempArr(a, b) = TempArr(a, b) / 10 ^ 9
        Next
    Next
    qoldiqlarRng = TempArr
End With
wbJamlovchi.Save

Finish = Timer

With Application
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
    .DisplayAlerts = True
    .StatusBar = Format(Finish - Start, "0.00") & " soniyada muvaffaqiyatli bajarildi!"
    .OnTime Now + TimeValue("00:00:06"), "KillStatBar"
End With

End Sub


Private Function IsBookOpen(wbName As String) As Boolean
    Dim wbBook As Workbook
    On Error Resume Next
    For Each wbBook In Workbooks
        If wbBook.Name <> ThisWorkbook.Name Then
            If Windows(wbBook.Name).Visible Then
                If wbBook.Name = wbName Then IsBookOpen = True: Exit For
            End If
        End If
    Next wbBook
End Function

Private Function IndexOf(element, arr) As Long
Dim i As Long

For i = 1 To UBound(arr)
    If element = arr(i) Then
        IndexOf = i
        Exit Function
    End If
Next i

End Function



