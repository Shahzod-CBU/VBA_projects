Attribute VB_Name = "DailyExtended"
Option Explicit

Const EMBED_ATTACHMENT As Long = 1454, RICHTEXT As Long = 1
Public Const stPath$ = "C:\Users\msd15\Desktop\Кунлик нархлар\"    'fayllar saqlanadigan papka
Const sServerName$ = "Main/ABM", sUserDB$ = "d_ext\mb\01msd14.nsf" 'lokal server va ma'lumotlar bazasi nomlari
                                                                   'Lotus => Файл => База данных => Свойства (Сервер, Имя файла)
Const Shablon$ = "C:\Users\msd15\Desktop\Кунлик нархлар\+Печатга Кунлик Шаблон.xlsx"    'Kun ochish u-n shablon fayl
Const sNotesSourceFolder$ = "$Inbox"    'Lotusda qaysi papkadan qidirish
Const lAllCnt As Long = 20      'progress bar u-n bo'laklar soni
Dim dblProgressWidth As Double, dblStep As Double, dblPercent As Double
Dim tdate, ttime As Date, fldrMonthName$, fldrDateName$, NewFolder$, fldrpath$
Dim NextTick As Date

Sub DailyPricesExtended()
'11.07.2018 da Shahzod tomonidan yozildi
'20.07.2018 da yangilandi (Fayllar nomini o'zgartirish)
'29.10.2018 da yangilandi (Lotus Notes bilan aloqa o'rnatildi)
'23.11.2018 da yangilandi (Bajarilganlikka tekshirish)

Dim NUIWorkspace As Object, noSession As Object, noDatabase As Object
Dim noView As Object, Doc As Object, noNextDocument As Object   'no - Notes object
Dim vaItem As Variant, vaAttachment As Variant, fso As Object
Dim fldrtime(14) As String
Dim NMoveDocsCollection As Object, Filename$, NewDoc As Object
Dim TayyorWb As Workbook
Dim TempArray(), Bugun, Start As Date, Finish As Date
Dim JamlovchiWorkBk As Workbook, HududiyWorkBk As Workbook
Dim sFileName$, sNewFileName$, objFSO As Object, objFile As Object
Dim Hudud$, h As Integer, i As Integer, InvalidChar$, sFolder$
Dim Natija, sFiles$, PechatDir$, LastColumn As Long, LastRow As Long

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
Start = Timer
Call Show_PrBar_Or_No(lAllCnt, "Bajarilmoqda...")
Call MyProgresBar

tdate = Date
'tdate = #11/6/2018#
fldrMonthName = Format(tdate, "MMMM")
fldrDateName = Format(tdate, "dd.mm.yyyy")
NewFolder = "Нархлар кунлик\" & fldrMonthName & "\" & fldrDateName
fldrpath = stPath & fldrMonthName

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.folderexists(fldrpath) Then fso.CreateFolder fldrpath 'Create a folder for a month
fldrpath = fldrpath & "\" & fldrDateName
If Not fso.folderexists(fldrpath) Then MkDir fldrpath 'Create a folder for a day

Set NUIWorkspace = CreateObject("Notes.NotesUIWorkspace") 'Front end UI is only exposed with OLE automation
Set noSession = CreateObject("Notes.NotesSession") 'Start a Notes session, using Lotus Domino Objects (COM classes)
Set noDatabase = noSession.GETDATABASE(sServerName, sUserDB)
If Not noDatabase.IsOpen Then noDatabase.Open
noDatabase.EnableFolder NewFolder  'Create a new folder in Lotus
Set NMoveDocsCollection = noDatabase.Search("", Nothing, 0)
Set noView = noDatabase.GetView(sNotesSourceFolder)

Set Doc = noView.GetLastDocument    'Get to last mail in inbox view
Call MyProgresBar

'Check only today's files and move them into the necessary folders
Do While Doc.GetItemValue("DeliveredDate")(0) >= tdate
    If Doc.HasEmbedded Then
        Set vaItem = Doc.GetFirstItem("Body")
        If vaItem.Type = RICHTEXT Then
            For Each vaAttachment In vaItem.EmbeddedObjects
                Filename = vaAttachment.Name
                If InStr(1, Filename, "кунлик", 3) <> 0 Then
                    NMoveDocsCollection.AddDocument Doc
                    vaAttachment.ExtractFile fldrpath & "\" & Filename
                    
                    ' *****************************************************
                    Hudud = Doc.GetItemValue("From")(0)
                    Hudud = Mid(Hudud, 4, 2)
                    Select Case Hudud
                        Case "25": h = 10
                        Case "14": h = 14
                        Case "03": h = 2
                        Case "04": h = 3
                        Case "05": h = 4
                        Case "07": h = 5
                        Case "08": h = 6
                        Case "09": h = 7
                        Case "10": h = 8
                        Case "11": h = 9
                        Case "12": h = 11
                        Case "06": h = 12
                        Case "13": h = 13
                        Case "02": h = 1
                    End Select
                    ttime = Doc.GetItemValue("DeliveredDate")(0)
                    fldrtime(h) = Format(ttime, "hh:mm")
                    ' ******************************************************
                
                End If
                Exit For
            Next vaAttachment
        End If
    End If
    Set Doc = noView.GetPrevDocument(Doc)
Loop
Call MyProgresBar

With NMoveDocsCollection
    .PutAllInFolder NewFolder, True
    .RemoveAllFromFolder sNotesSourceFolder
End With

noView.Refresh
NUIWorkspace.VIEWREFRESH

'Extract files from archives and then delete the archives
sFolder = fldrpath & IIf(Right(fldrpath, 1) = Application.PathSeparator, "", Application.PathSeparator)
sFiles = Dir(sFolder)
Do While sFiles <> ""
    If sFiles Like "*.rar" Or sFiles Like "*.zip" Or sFiles Like "*.7z" Then
        Natija = UnRAR(sFolder, sFiles)
    End If
    sFiles = Dir
Loop
Call MyProgresBar

'Open the template file and create a new folder to save it
PechatDir = "C:\Users\msd15\Desktop\Кунлик нархлар\!Печать\" & fldrMonthName
If Not fso.folderexists(PechatDir) Then fso.CreateFolder PechatDir
Workbooks.Open Shablon, UpdateLinks:=False

'Start working with the downloaded files
Set JamlovchiWorkBk = ActiveWorkbook
With Sheets("Кунлик нархлар")
    LastColumn = .UsedRange.Column + .UsedRange.Columns.Count - 1
    LastRow = .UsedRange.Row + .UsedRange.Rows.Count - 1
End With

ReDim TempArray(1 To LastRow - 4, 1)

With JamlovchiWorkBk.Sheets("Кунлик нархлар")
    .Range("D5:Q98").ClearContents
    .Range("D2") = tdate
    Bugun = .Cells(2, 4).Value
End With

sFiles = Dir(sFolder & "*.xls*")
Call MyProgresBar

Do While sFiles <> ""
      If sFiles <> JamlovchiWorkBk.Name Or InStr(1, sFiles, "кунлик", 3) <> 0 Then 'check if the names of files include the specified word
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
        Hudud = Left(Cells(2, 2).Value, 5)
            If Hudud = "Тошке" Then
                Hudud = Left(Cells(2, 2).Value, WorksheetFunction.Find(" ", Cells(2, 2), 3) + 1)
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
                    Case "ФАР" & ChrW(1170) & "О": h = 11
                    Case ChrW(1178) & "аш" & ChrW(1179) & "а": h = 12
                    Case "Хораз": h = 13
                    Case Else: h = 1
            End Select
            End If
        TempArray = Cells.Find(What:=Bugun, LookIn:=xlValues).Offset(2, 0).Range(Cells(1, 1), Cells(LastRow - 4, 1)).Value
        ActiveWorkbook.Close False
        JamlovchiWorkBk.Activate
        Sheets("Кунлик нархлар").Range(Sheets("Кунлик нархлар").Cells(5, h + 3), Sheets("Кунлик нархлар").Cells(LastRow, h + 3)).Value = TempArray
        '**************
        Sheets("Кунлик нархлар").Cells(99, h + 3) = fldrtime(h)
        '**************
     End If
    sFiles = Dir
    Call MyProgresBar
Handler:
Loop

JamlovchiWorkBk.Worksheets("Печатга").Activate
JamlovchiWorkBk.SaveAs PechatDir & "\" & "+Печатга Кунлик " & Format(tdate, "dd.mm.yyyy") & ".xlsx"
Set TayyorWb = ActiveWorkbook
TayyorWb.Save
Call MyProgresBar

'Create a new memo in Lotus Notes
If WorksheetFunction.Count(Sheets("Кунлик нархлар").Range("D5:Q5")) = 14 Then
    Dim EditDoc As Object, AttachME As Object, Attachment$, EmbedObj
    
    Set NewDoc = noDatabase.CREATEDOCUMENT
    With NewDoc
        .form = "Memo"
        .Subject = "Кунлик нархлар " & fldrDateName & " холатига"
        .SendTo = Array("01pom3/MB/Uzb@Banks, 01msd5/MB/Uzb@Banks, 01msd7/MB/Uzb@Banks")
        Attachment = TayyorWb.FullName
        Set AttachME = .CREATERICHTEXTITEM("Attachment")
        Set EmbedObj = AttachME.EMBEDOBJECT(1454, "", Attachment, "Attachment")
        .SAVEMESSAGEONSEND = True
        Set EditDoc = NUIWorkspace.EditDocument(True, NewDoc)
    End With
End If

'Release objects from memory.
Set NewDoc = Nothing
Set EditDoc = Nothing
Set noNextDocument = Nothing
Set NMoveDocsCollection = Nothing
Set Doc = Nothing
Set noView = Nothing
Set noDatabase = Nothing
Set noSession = Nothing
Set NUIWorkspace = Nothing

Call MyProgresBar

With Application
 .Calculation = xlCalculationAutomatic
 .ScreenUpdating = True
 .DisplayAlerts = True
 Finish = Timer
 .StatusBar = Format(Finish - Start, "0.00") & " секундда муваффа" & ChrW(1179) & "иятли бажарилди!"
 .OnTime Now + TimeValue("00:00:06"), "KillStatBar"
End With
Unload frmStatusBar

End Sub

Private Function UnRAR(fldrpath As String, sArhivName As String)
    Const sWinRarAppPath As String = "C:\Program Files (x86)\WinRAR\WinRAR.exe"
    Dim sWinRarApp As String
    sWinRarApp = sWinRarAppPath & " E -o+"
    UnRAR = Shell(sWinRarApp & " """ & fldrpath & "\" & sArhivName & """ """ & fldrpath & """ ", vbHide)
    Application.Wait (Now + TimeValue("0:00:01")) 'Let the shell do the commands and close
    Kill fldrpath & sArhivName
End Function
Private Sub MyProgresBar()
    dblProgressWidth = dblProgressWidth + dblStep
    frmStatusBar.FrameProgress.Width = dblProgressWidth - dblStep
    If dblProgressWidth > dblPercent Then
        frmStatusBar.lblPercentWhite.Caption = Format(dblPercent / frmStatusBar.FramePrgBar.Width, "0%")
        frmStatusBar.lblPercentBlack.Caption = frmStatusBar.lblPercentWhite.Caption
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

'Private Sub KillStatBar()
'     'Application.StatusBar = False
' End Sub

Private Sub CheckIsDone()
    Dim MsgTxt$
    tdate = Date
    fldrMonthName = Format(tdate, "MMMM")
    fldrDateName = Format(tdate, "dd.mm.yyyy")
    NewFolder = "Нархлар кунлик\" & fldrMonthName & "\" & fldrDateName
    fldrpath = stPath & fldrMonthName
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.folderexists(fldrpath) Then
       MsgTxt = "Bugun uchun kunlik narxlar qilinmagan." & vbNewLine
       MsgTxt = MsgTxt & "Yig'uvchi majros bajarilsinmi?"
       Select Case MsgBox(MsgTxt, vbYesNoCancel + vbQuestion + vbDefaultButton1, "Eslatma")
            Case vbYes
                Call DailyPricesExtended
            Case vbNo
                NextTick = Now + TimeValue("00:05:00")
                Application.OnTime NextTick, "CheckIsDone"
            Case vbCancel
                Exit Sub
       End Select
    End If
End Sub

