Attribute VB_Name = "Liquidity"
Option Explicit
Public Const stPath$ = "D:"                     'fayllar saqlanadigan papka
Public Const templatesPath$ = "c:\Users\Asus\Documents\������� �����"
Const downloadsPath$ = "c:\Users\Asus\Downloads"

'Public Const stPath$ = "C:\Users\Salixov_M\Desktop\���� ������"
'Public Const templatesPath$ = "C:\Users\Salixov_M\Desktop\������� �����"
'Const downloadsPath$ = "C:\Users\Salixov_M\Downloads"

Public tdate, ttime As Date, fldrMonthName$, fldrYear$, NewFolder$, fldrpath$, monthNum$
Public fldrpathMonth$, fldrpathResults$

Sub �������_������()

'14.01.2018 da Shahzod tomonidan yozildi
'18.02.2018 da yangilandi (Banklar nomini qo'shish)
'26.02.2018 da yangilandi (Sof ta'sirni topish)
'15.03.2018 da yangilandi (CSV fayldan olish)

Dim i As Long, RowNumber As Long, Start As Date, Finish As Date, lAllCnt As Long
Dim KattaArray(), BalansSchot(), TurProvod(), ListlarNomi(), PivotListlar(), x As Integer
Dim n As Integer, IshchiKitob As Workbook, LastRow As Long, JadvalNomi$, UstunNomi$, SummaSchot$
Dim TasirField As PivotField, downloadedCSV As String, destCell As Range
Dim Harbiy As String, Hukumat As String, tulovMarkazi$
Dim fso As Object

tdate = Application.InputBox("Sanani ko'rsating", "Korschot", Format(Now, "dd.mm.yyyy"))
If tdate = "False" Then Exit Sub

ListlarNomi = Array("Dr", "Cr")
PivotListlar = Array("PivotDr", "PivotCr")
TurProvod = Array("��", "��")
lAllCnt = 25
Call Show_PrBar_Or_No(lAllCnt, "Bajarilmoqda...")

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Start = Timer
Call MyProgresBar

'tdate = Date
'tdate = #12/20/2018#
fldrYear = Format(tdate, "yyyy")
monthNum = Format(tdate, "mm")
fldrMonthName = Format(tdate, "MMMM")
NewFolder = "������ ����. ���� " & fldrYear '& "\" & monthNum & " " & fldrMonthName & "\" & fldrMonthName
fldrpath = stPath & "\" & NewFolder

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.folderexists(fldrpath) Then fso.CreateFolder fldrpath            'Create a folder for a year
fldrpathMonth = fldrpath & "\" & monthNum & " " & fldrMonthName
If Not fso.folderexists(fldrpathMonth) Then fso.CreateFolder fldrpathMonth  'Create a folder for a month
fldrpathResults = fldrpathMonth & "\" & fldrMonthName
If Not fso.folderexists(fldrpathResults) Then MkDir fldrpathResults         'Create a folder for results

Set IshchiKitob = Application.Workbooks.Open(templatesPath & "\" & "������������������ (�����������).xlsx", False)
'Set IshchiKitob = Application.Workbooks.Open("C:\Users\�������������\Desktop\Shahzod\������� �����\������������������ (�����������).xlsx", False)

For n = 0 To 1
    IshchiKitob.Worksheets.Add
    Call MyProgresBar
    
    ActiveSheet.Name = ListlarNomi(n)
    tulovMarkazi = "001"          'Markaziy bank
    downloadedCSV = "LiquidityData_" & tulovMarkazi & "_for_" & Format(tdate, "dd.mm.yyyy") & "_" & ListlarNomi(n)
    Set destCell = Range("$A$2")
    Call MyProgresBar
    
    ImportCSV downloadsPath, downloadedCSV, destCell
    Kill downloadsPath & Application.PathSeparator & downloadedCSV & ".csv"
    
    With IshchiKitob.Worksheets(ListlarNomi(n))
        Call MyProgresBar
        
        With Range(Cells(1, 1), Cells(1, 11))
            .Value = Array("�", "���� ��", "������� ���� ��", "���� ��", "������� ���� ��", _
                            "�����" & TurProvod(n), "��", "��", "������1", "������2", "����")
            .HorizontalAlignment = xlLeft
            .Font.Bold = True
        End With

        LastRow = .UsedRange.Row + .UsedRange.Rows.count - 1
        .Range(Cells(2, 6), Cells(LastRow, 6)).Replace What:=".", Replacement:=","
        
        ReDim BalansSchot(1 To LastRow - 2, 1 To 1)

        'Debetlangan balans hisobvaraqlarini olib olamiz
        BalansSchot = Range(Cells(2, 3), Cells(LastRow, 3)).Value2
        For i = 1 To UBound(BalansSchot)
            BalansSchot(i, 1) = Left(BalansSchot(i, 1), 5)
        Next i
        Range(Cells(2, 7), Cells(LastRow, 7)).Value = BalansSchot
        
        'Kreditlangan balans hisobvaraqlarini olib olamiz
        BalansSchot = Range(Cells(2, 5), Cells(LastRow, 5)).Value2
        For i = 1 To UBound(BalansSchot)
            BalansSchot(i, 1) = Left(BalansSchot(i, 1), 5)
        Next i
        Range(Cells(2, 8), Cells(LastRow, 8)).Value = BalansSchot

        'Summani sonlarga o'giramiz
        BalansSchot = Range(Cells(2, 6), Cells(LastRow, 6)).Value
        For i = 1 To UBound(BalansSchot)
            BalansSchot(i, 1) = CDbl(BalansSchot(i, 1))
        Next i
        Range(Cells(2, 6), Cells(LastRow, 6)).Value = BalansSchot

        .Range(Cells(2, 2), Cells(LastRow, 6)).NumberFormat = "#,##0"
        JadvalNomi = ListlarNomi(n) & "Jadval"
        Call MyProgresBar
        
        Harbiy = ChrW(1202) & "����� ���"
        Hukumat = ChrW(1202) & "������"
        
        'Svodniy jadval tayyorlaymiz
        .ListObjects.Add(xlSrcRange, Range("$A$1:$K$" & LastRow), , xlYes).Name = JadvalNomi
        .ListObjects(JadvalNomi).ShowTableStyleRowStripes = False
        .ListObjects(JadvalNomi).TableStyle = "TableStyleMedium2"
        .Range(JadvalNomi & "[������1]").FormulaR1C1 = _
            "=IF([@" & TurProvod(n) & "]<>21596,IFERROR(VLOOKUP([@" & TurProvod(n) & "], ����������������, 2, 0)," & _
                "LOOKUP(LEFT([@" & TurProvod(n) & "]),{""4"",""5""},{""�������"",""�������""}))," & _
                "IFERROR(VLOOKUP(MID([@[������� ���� " & TurProvod(n) & "]],10,8), ����������,2,0)," & _
                "IFERROR(VLOOKUP(MID([@[������� ���� " & TurProvod(n) & "]],10,11), ���������������,2,0)," & Chr(34) & Harbiy & Chr(34) & ")))"
        Call MyProgresBar

        .Range(JadvalNomi & "[������2]").FormulaR1C1 = _
            "=IF([@" & TurProvod(n) & "]<>21596,IFERROR(VLOOKUP([@" & TurProvod(n) & "],����������������,3,0)," & _
                "IF(OR([@������1]=""�������"",[@������1]=""�������""),""�� ������� ������������""))," & _
                "IFERROR(VLOOKUP(MID([@[������� ���� " & TurProvod(n) & "]],10,11),���������������,2,0)," & _
                "IF([@������1]=" & Chr(34) & Harbiy & Chr(34) & ", " & Chr(34) & Hukumat & Chr(34) & ",""�� ������� ������������"")))"
        
        Call MyProgresBar
        
        .Range(JadvalNomi & "[����]").FormulaR1C1 = "=VLOOKUP([@���� " & TurProvod(1 - n) & "], BankNums, 3, 0)"

        SummaSchot = "�����" & TurProvod(n)
        IshchiKitob.Queries.Add Name:=JadvalNomi, Formula:= _
            "let" & Chr(13) & Chr(10) & "    �������� = Excel.CurrentWorkbook(){[Name=" & Chr(34) & JadvalNomi & Chr(34) & "]}[Content]," & _
                    Chr(13) & Chr(10) & "    #""���������� ���"" = Table.TransformColumnTypes(��������,{{""�"", Int64.Type}, {""���� ��"", Int64.Type}, {""������� ���� ��"", type number}, {""���� ��"", Int64.Type}, {""������� ���� ��"", type number}," & _
                    "{" & Chr(34) & SummaSchot & Chr(34) & ", type number}, {""��"", Int64.Type},{""��"", Int64.Type}, {""������1"", type text}, {""������2"", type text}, {""����"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""���������� ���"""
        IshchiKitob.Connections.Add2 _
            "������ � " & JadvalNomi, "���������� � �������� " & JadvalNomi & " � �����.", _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & JadvalNomi _
            , "SELECT * FROM [" & JadvalNomi & "]", 2

    End With
    Call MyProgresBar
    
    IshchiKitob.Worksheets.Add
    ActiveSheet.Name = PivotListlar(n)
    Call MyProgresBar
    
    IshchiKitob.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=JadvalNomi, _
       Version:=xlPivotTableVersion12).CreatePivotTable TableDestination:=PivotListlar(n) & "!R3C1", _
       TableName:="Pivot" & n + 1, DefaultVersion:=xlPivotTableVersion12
    UstunNomi = TurProvod(n) & " ������"
    Call MyProgresBar
    
    Rem Svodniy jadval parametrlarini to'g'rilaymiz
'    On Error GoTo 0

    With Sheets(PivotListlar(n)).PivotTables("Pivot" & n + 1)
        .AddDataField ActiveSheet.PivotTables("Pivot" & n + 1).PivotFields(SummaSchot), UstunNomi, xlSum
        .PivotFields(UstunNomi).NumberFormat = "#,##0"
        With .PivotFields("������2")
            .Orientation = xlRowField
            .Position = 1
            .AutoSort xlDescending, UstunNomi
            .PivotItems("�������").Visible = False
        End With
        With .PivotFields("������1")
            .Orientation = xlRowField
            .Position = 2
            .AutoSort xlDescending, UstunNomi
        End With
        .PivotFields("��").Orientation = xlPageField
        .PivotFields("��").PivotItems("27402").Visible = False
    End With
    Call MyProgresBar
    
Next n

With IshchiKitob
    .Queries.Add Name:="TotalPivot", Formula:= _
        "let" & Chr(13) & Chr(10) & "    �������� = Table.Combine({DrJadval, CrJadval})" & Chr(13) & Chr(10) & "in" & Chr(13) & Chr(10) & "    ��������"
    .Connections.Add2 _
        "������ � TotalPivot", "���������� � �������� ""TotalPivot"" � �����.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TotalPivot" _
        , "SELECT * FROM [TotalPivot]", 2
    .Worksheets.Add
    Call MyProgresBar
    
    .ActiveSheet.Name = "PivotNet"
    .PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("������ � TotalPivot"), Version:=6).CreatePivotTable _
        TableDestination:="PivotNet!R3C1", TableName:="PivotNet", _
        DefaultVersion:=6
End With
Call MyProgresBar

'On Error Resume Next
With Sheets("PivotNet").PivotTables("PivotNet")
    .CalculatedFields.Add "�������", "=(�������-�������)/10^9"
    .PivotFields("�������").Orientation = xlDataField
    .PivotFields("����� �� ���� �������").Caption = "��� �������"
    '.PivotFields("Sum of �������").Caption = "��� �������"
    .DataBodyRange.NumberFormat = "#,##0"
    '.PivotFields("����").Orientation = xlColumnField
    With .PivotFields("������2")
        .Orientation = xlRowField
        .Position = 1
        .AutoSort xlDescending, "��� �������"
        .PivotItems("�������").Visible = False
    End With
'    With .PivotFields("������1")
'        .Orientation = xlRowField
'        .Position = 2
'        .AutoSort xlDescending, "��� �������"
'    End With
    .PivotFields("��").Orientation = xlPageField
    .PivotFields("��").PivotItems("27402").Visible = False
    Set TasirField = .PivotFields("��� �������")
End With
Call MyProgresBar

'Conditional formatting
With TasirField.DataRange.FormatConditions
    .Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
    With .Item(1)
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
        .StopIfTrue = False
        .ScopeType = xlDataFieldScope
    End With
End With
Call MyProgresBar

IshchiKitob.SaveAs fldrpathResults & "\" & "������� ������ " & Format(tdate, "dd.mm.yyyy") & ".xlsx"
Call MyProgresBar

Finish = Timer
Unload frmStatusBar

With Application
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
    .DisplayAlerts = True
    .StatusBar = Format(Finish - Start, "0.00") & " soniyada muvaffaqiyatli bajarildi!"
    .OnTime Now + TimeValue("00:00:06"), "KillStatBar"
End With

End Sub

Private Function ImportCSV(filePath, Filename, destinationCell)
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & filePath & Application.PathSeparator & Filename & ".csv" _
        , Destination:=destinationCell)
        .Name = Filename
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 866
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 2, 2, 2, 2, 2)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    ActiveSheet.QueryTables(Filename).Delete
End Function

Private Sub MyProgresBar()
    dblProgressWidth = dblProgressWidth + dblStep
    frmStatusBar.FrameProgress.Width = dblProgressWidth - dblStep
    If dblProgressWidth > dblPercent Then
        frmStatusBar.lblPercentWhite.Caption = Format(dblPercent / frmStatusBar.FramePrgBar.Width, "0%")
        frmStatusBar.lblPercentBlack.Caption = frmStatusBar.lblPercentWhite.Caption
        If frmStatusBar.lblPercentWhite.Caption Like "96%" Then
            frmStatusBar.Caption = "Saqlanmoqda..."
        End If
        dblPercent = dblPercent + dblStep
        frmStatusBar.Repaint
        DoEvents
    End If
End Sub
Private Function Show_PrBar_Or_No(lCnt As Long, sUfCaption As String)
    frmStatusBar.Caption = sUfCaption
    dblStep = frmStatusBar.FramePrgBar.Width / lCnt
    frmStatusBar.lblPercentWhite.Left = 96
    frmStatusBar.lblPercentBlack.Left = frmStatusBar.lblPercentWhite.Left
    
    frmStatusBar.Show 0
    dblPercent = 0: dblProgressWidth = 0
End Function

' Private Sub KillStatBar()
'     Application.StatusBar = False
' End Sub


