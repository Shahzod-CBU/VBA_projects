Attribute VB_Name = "Exp_LegEntity"
Option Explicit

Sub InflatsionKutilmaKorxona()

'25.07.2018 da Shahzod tomonidan yozildi

Dim TempArray()
Dim JamlovchiWorkBk As Workbook, HududiyWorkBk As Workbook
Dim lr As Long, sFolder As String
Dim lAllCnt As Long, sFiles As String
Dim sFileName As String, sNewFileName As String
Dim Start As Date, Finish As Date
Dim objFSO As Object, objFile As Object
Dim Hudud As String, h As Integer, i As Integer, InvalidChar As String
Dim SurovMuddati As String, TuldirilganSana As String, TuldirganXodim As String
Const lMaxQuad As Long = 15
lAllCnt = 15
ReDim TempArray(1 To 100, 1 To 103)

Set JamlovchiWorkBk = ActiveWorkbook

With Application.FileDialog(msoFileDialogFolderPicker)
 If .Show = False Then Exit Sub
 sFolder = .SelectedItems(1)
End With
 sFolder = sFolder & IIf(Right(sFolder, 1) = Application.PathSeparator, "", Application.PathSeparator)

Start = Timer

   Application.ScreenUpdating = False
   Application.Calculation = xlCalculationManual
   Application.DisplayAlerts = False
    
sFiles = Dir(sFolder & "*.xls*")

Do While sFiles <> ""
      If Not sFiles = JamlovchiWorkBk.Name Then
        On Error Resume Next
        Workbooks.Open sFolder & sFiles, UpdateLinks:=False
        If Err <> 0 Then
            i = 1
          Do
            InvalidChar = WorksheetFunction.Find(Chr(63), sFiles, i)
                If i > 1 Then
                sNewFileName = WorksheetFunction.Replace(sNewFileName, InvalidChar, 1, "к")
                sFileName = WorksheetFunction.Replace(sFileName, InvalidChar, 1, ChrW(1178))
                Else
                sNewFileName = WorksheetFunction.Replace(sFiles, InvalidChar, 1, "к")
                sFileName = WorksheetFunction.Replace(sFiles, InvalidChar, 1, ChrW(1178))
                End If
            i = InvalidChar + 1
          Loop While Not IsError(WorksheetFunction.Find(Chr(63), sNewFileName))
            sFileName = sFolder & sFileName
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFSO.GetFile(sFileName)
            objFile.Name = sNewFileName
            sFiles = sNewFileName
            GoTo Handler
        End If
        ActiveWorkbook.Worksheets("База-Корхона").Activate
        Hudud = Left(Cells(5, 3).Value, 5)
            If Hudud = "Тошке" Then
                Hudud = Left(Cells(5, 3).Value, WorksheetFunction.Find(" ", Cells(5, 3), 3) + 1)
                Select Case Hudud
                    Case "Тошкент в": h = 10
                    Case "Тошкент ш": h = 14
                End Select
            Else
                Select Case Hudud
                    Case "Андиж": h = 2
                    Case "Бухор": h = 3
                    Case "Жизза": h = 4
                    Case "Навои": h = 5
                    Case "Наман": h = 6
                    Case "Самар": h = 7
                    Case "Сирда": h = 8
                    Case "Сурхо": h = 9
                    Case "Фар" & ChrW(1171) & "о": h = 11
                    Case ChrW(1178) & "аш" & ChrW(1179) & "а": h = 12
                    Case "Хораз": h = 13
                    Case Else: h = 1
            End Select
            End If
        SurovMuddati = Cells(6, 3)
        TuldirilganSana = Cells(5, 6)
        TuldirganXodim = Cells(6, 6)
        TempArray = Range(Cells(11, 1), Cells(60, 37)).Value
        ActiveWorkbook.Close False
        JamlovchiWorkBk.Activate
        Sheets(h + 1).Range(Sheets(h + 1).Cells(11, 1), Sheets(h + 1).Cells(60, 37)).Value = TempArray
        Sheets(h + 1).Cells(6, 3) = SurovMuddati
        Sheets(h + 1).Cells(5, 6) = TuldirilganSana
        Sheets(h + 1).Cells(6, 6) = TuldirganXodim
        If Not IsEmpty(Sheets(h + 1).Cells(6, 3)) Then Sheets(h + 1).Tab.Color = 12611584
     End If
    sFiles = Dir
    lr = lr + 1
    Application.StatusBar = "Бажарилмо" & ChrW(1179) & "да: " & Int(100 * lr / lAllCnt) & "%" & String(CLng(lMaxQuad * lr / lAllCnt), ChrW(9632)) & String(lMaxQuad - CLng(lMaxQuad * lr / lAllCnt), ChrW(9633))
Handler:
Loop

With Application
 .Calculation = xlCalculationAutomatic
 .ScreenUpdating = True
 .DisplayAlerts = True
 Finish = Timer
 .StatusBar = Format(Finish - Start, "0.00") & " секундда муваффа" & ChrW(1179) & "иятли бажарилди!"
 .OnTime Now + TimeValue("00:00:06"), "KillStatBar"
End With

End Sub

' Private Sub KillStatBar()
'     Application.StatusBar = False
' End Sub


