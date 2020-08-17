VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Balans 
   Caption         =   "Kunlik balans"
   ClientHeight    =   3828
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   9804.001
   OleObjectBlob   =   "Balans.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Balans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


 Option Explicit
 Public avFiles
 
Private Sub Chiqish_Click()
    Unload Me
    Tugatish = True
End Sub

Private Sub Kursatish_Click()
    avFiles = Application.GetOpenFilename("Excel files(*.xls*),*.xls*", 1, "Выбрать Excel файлы", , False)
    If VarType(avFiles) = vbBoolean Then Exit Sub
    Yul.Value = avFiles
End Sub

Private Sub OK_Click()
    Dim SanaManba, ListManba(), bListlar(), Sanalar(), i As Integer
    Dim ListOti As String, JoriyTest As Long, OldingiKunTest As Integer
    Dim SofFoyda As Long, TaqsimFoyda As Long
    ListManba = Array("ЦБ", "КБ")
    bListlar = Array(ЦБ, КБ)
    Sanalar = Array(tbSana, tbOldingiSana)
    
    If Len(Yul) = 0 Then
        MsgBox "Iltimos faylni ko'rsating", vbCritical, "Xatolik"
        Exit Sub
    End If
    If ЦБ.Value = False And КБ.Value = False Then
        MsgBox "Iltimos kamida bitta list tanlang", vbCritical, "Xatolik"
        Exit Sub
    End If
    If IsDate(tbSana) = False Then
        With tbSana
            MsgBox "Sana noto'g'ri kiritilgan!", vbCritical, "Xato"
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.text)
            Exit Sub
        End With
        Exit Sub
    End If
    If IsDate(tbOldingiSana) = False Then
        With tbOldingiSana
            MsgBox "Sana noto'g'ri kiritilgan!", vbCritical, "Xato"
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.text)
            Exit Sub
        End With
        Exit Sub
    End If
    
    If CDate(tbSana) <= CDate(tbOldingiSana) Then
        MsgBox "Oldingi sana Joriy sanadan kichik bo'lishi lozim!", vbCritical, "Xatolik"
        Exit Sub
    End If
    
    On Error Resume Next
    
    OldingiKunTest = wbJamlovchi.Worksheets("ЦБ(конс_new)").Rows(4).Cells.Find(What:=CDate(tbOldingiSana), LookIn:=xlFormulas).Column
    If Err <> 0 Then
        MsgBox "Ko'rsatilgan Oldingi sana uchun hali balansda kun ochilmagan", vbCritical, "Xatolik"
        Exit Sub
    End If
    
Application.ScreenUpdating = False

    Fayl = Balans.Yul.Value
    Set wbManba = Workbooks.Open(Fayl, UpdateLinks:=False)
    If Err <> 0 Then
        MsgBox Err.Description
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    For i = 0 To 1
        If bListlar(i) Then
            On Error Resume Next
            ListOti = wbManba.Worksheets(ListManba(i)).Name
            If Err <> 0 Then
                MsgBox "Siz tanlagan " & ListManba(i) & " listi Faylda topilmadi!", vbCritical, "Xatolik"
                wbManba.Close False
                Application.ScreenUpdating = True
                Exit Sub
            End If
            SanaManba = Mid(wbManba.Sheets(ListManba(i)).Cells(3, 1), InStr(1, wbManba.Sheets(ListManba(i)).Cells(3, 1), Chr(58)) + 1, 10)
            SanaManba = CDate(Format(SanaManba, "mm/dd/yyyy"))
            If CDate(tbSana) <> SanaManba Then
                MsgBox "Siz ko'rsatgan Joriy sana Fayldagi sanaga mos kelmayapti." & vbNewLine & _
                    "Ma'lumotlar to'g'ri kiritilganligini yana bir marta tekshiring!", vbCritical, "Xatolik"
                wbManba.Close False
                Application.ScreenUpdating = True
                Exit Sub
            End If
            With wbManba.Worksheets(ListManba(i))
                SofFoyda = .Cells.Find(What:="31206", LookIn:=xlValues, SearchOrder:=xlByColumns).Row
                If Err <> 0 Then
                    TaqsimFoyda = .Cells.Find(What:="31203", LookIn:=xlValues, SearchOrder:=xlByColumns).Row
                    SofFoyda = TaqsimFoyda + 1
                    .Rows(SofFoyda).Insert
                    .Cells(SofFoyda, 3) = "31206"
                    Err.Clear
                End If
            End With
        End If
    Next
    
Me.Hide
End Sub

Private Sub spbOldingiSana_Change()
    tbOldingiSana.Value = Format(spbOldingiSana.Value, "dd/mm/yyyy")
End Sub

Private Sub spbSana_Change()
    tbSana.Value = Format(spbSana.Value, "dd/mm/yyyy")
End Sub

Private Sub tbOldingiSana_Change()
Dim NewVal2 As Long, vSana2
On Error GoTo Chiqish
vSana2 = CLng(CDate(tbOldingiSana.Value))
    If IsNumeric(vSana2) Then
        NewVal2 = CLng(vSana2)
        If NewVal2 >= spbOldingiSana.Min And _
            NewVal2 <= spbOldingiSana.Max Then
            spbOldingiSana.Value = NewVal2
        End If
    End If
Chiqish:
End Sub

Private Sub tbSana_Change()
Dim NewVal As Long, vSana
On Error GoTo Chiqish
vSana = CLng(CDate(tbSana.Value))
    If IsNumeric(vSana) Then
        NewVal = CLng(vSana)
        If NewVal >= spbSana.Min And _
            NewVal <= spbSana.Max Then
            spbSana.Value = NewVal
        End If
    End If
Chiqish:
End Sub

Private Sub UserForm_Initialize()
    Dim joriy_sana As Range, new_sana As Date
    tbSana.Value = Format(ExceptDaysOff(Date - 1), "dd/mm/yyyy")
    tbOldingiSana.Value = Format(ExceptDaysOff(ExceptDaysOff(Date - 1) - 1), "dd/mm/yyyy")
'    Yul = "D:\Монетар сиёсат\ДКП\Баланс\Баланс_ЦБ_КБ_кунлик (минг)_17.08.xls"
    
'    Set joriy_sana = Workbooks("Баланс макро.xlsm").Sheets(1).Range("A1")
'    new_sana = DateSerial(Year(joriy_sana), Month(joriy_sana) + 2, 0)
'    Yul = "D:\Balance\баланс 2014-2019\" & Format(new_sana, "YYYY-MM-DD") & ".xls"
''    Debug.Print new_sana
'    tbSana.Value = new_sana
'    tbOldingiSana.Value = joriy_sana
'    joriy_sana.Value = new_sana
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Tugatish = True
        Cancel = 0
    End If
End Sub

