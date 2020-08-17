Attribute VB_Name = "DailyBalance"
Option Explicit
Public wbManba As Workbook, wbJamlovchi As Workbook
Public Tugatish As Boolean, Fayl As String
Public dblProgressWidth As Double, dblStep As Double, dblPercent As Double

Sub KunlikBalans()

'08.08.-04.09.2018 da Shahzod Maxmudov tomonidan yozildi
'18.09.2018 da yangilandi (Yangi qo'shilgan schotlarni aniqlash va qo'shish)
'29.09.2018 da yangilandi (ErrHandler optimallashtirildi)
'26.12.2018 da yangilandi (31206-schot yo'q bo'lsa, uni manba listga qo'shish)

'HukumatgaKredit listidagi sanani KB dan oladigan qilib o'zgartirish k-k

Rem O'zgaruvchilarni e'lon qilamiz
Dim LastRow As Long, LastClm As Long, ListJamlovchi(), ListManba(), ManbaQoldiqlar(1)
Dim acShJam As Worksheet, acShMan As Worksheet     'acSh - ActiveSheet; Jam - Jamlovchi, Man - Manba
Dim OldingiKun As Variant, OldingiKunClm(1) As Long, JoriySana As Date, TempRange As Range
Dim SofFoyda As String, OxirgiQtr(1) As Long, i As Integer, m As Long, n As Long, j As Long
Dim bListlar(), arDoimiySch(), arTblSch(), arJoriySch(), arJoriyQld(), arQoldiqlar()   'ar - "Array"
Dim Start As Date, Finish As Date, OxirgiKatak As Range, arFoizUzg(), f As Integer, Wanted
Dim FirstRow As Long, QatorSoni As Long, b As Integer, RowN As Long, JamFayl As String
Dim JamiAktiv As Double, JamiPassiv As Double, Farq As Double, Jamilar(), ErrMsg As String
Dim JamiKapital As Long, Schot31206 As Long, Schot31200 As Long, y(1) As Integer
Dim OyUzgarishi As Boolean, OxirgiUstun As Long, arFrm(), Uzgarish(), UzgarRow As Long
Dim ForStart As Long, ForEnd As Long, ForLen As Long, NewClm As Long   'For - "Formula"
Dim lAllCnt As Long, TempArray(), JoriyTest As Long, JoriyTestUchun, AlreadyDone As Boolean
Dim bYangiSchot(), arYangiSchot(), MsgNewAcc(1) As String, sSchot(1) As String, XatoList$
Dim arYangidanOldin(), t As Long, YangiSatr As Long, q As Long, BeforeNew(), BeforeNewRow
Dim sYangiSchot(1), bNewAcc(1) As Boolean, h As Integer, rnNewAcc As Range, arNewAcc(1)
Dim ErrManba() As Boolean
Set wbJamlovchi = ActiveWorkbook
Balans.Show

If Tugatish Then
    Tugatish = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
End If

ListJamlovchi = Array("ЦБ(конс_new)", "КБ(конс_new)", "ЧВшА", "КЭ_конс (2)", "КП", "М2", "B (M2)", "СТРУК_Норкулов", "Таблица 9 (Раис)", "Таблица 5", _
            "Таблица 3", "Таблица 4", "Вал Активы БС", "Таблица 6", "Таблица 7", "Таблица 8_", "Таблица 1", "Табл2", "Таблица 2 ", "Рассм табл к табл3")
ListManba = Array("ЦБ", "КБ")
ReDim ErrManba(0 To UBound(ListJamlovchi) - 4)
lAllCnt = (UBound(ListJamlovchi) - 4) * 2 + 2
bListlar = Array(Balans.ЦБ.Value, Balans.КБ.Value)

OldingiKun = CDate(Balans.tbOldingiSana)
JoriySana = CDate(Balans.tbSana)
If Month(JoriySana) - Month(OldingiKun) <> 0 Then OyUzgarishi = True
OldingiKun = Format(OldingiKun, "[$-419]D MMM YY;@")
JoriyTestUchun = Format(JoriySana, "[$-419]D MMM YY;@")
SofFoyda = "Чистая прибыль (убыток) (активно-пассивный)"
Unload Balans

On Error Resume Next
Wanted = JoriyTestUchun: JoriyTest = wbJamlovchi.Sheets(ListJamlovchi(0)).Rows(4).Find(What:=Wanted, LookIn:=xlValues).Column
If Err = 0 Then AlreadyDone = True
Err.Clear
On Error GoTo ErrHandler

If (bListlar(0) Or bListlar(1)) And AlreadyDone Then lAllCnt = 3

Call Show_PrBar_Or_No(lAllCnt, "Bajarilmoqda...")
Start = Timer
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Rem Manba balansdan ma'lumotlarni olib olamiz
For i = 0 To 1
    Call MyProgresBar
    Set acShJam = wbJamlovchi.Sheets(ListJamlovchi(i))
    With acShJam
        Rem Asosiy va boshqa kerak schotlarning pozitsiyalarini aniqlab olamiz
        Wanted = OldingiKun: OldingiKunClm(i) = .Rows(4).Find(What:=Wanted, LookIn:=xlValues).Column
        Wanted = SofFoyda: OxirgiQtr(i) = .Columns(2).Find(What:=Wanted, LookIn:=xlFormulas, SearchOrder:=xlByColumns).Row + 3
        ReDim arDoimiySch(1 To OxirgiQtr(i) - 8, 1 To 1)   'Jamlovchi balansdan doimiy schotlarni olib olamiz
        arDoimiySch = .Range(.Cells(8, 1), .Cells(OxirgiQtr(i) - 1, 1))
    End With

    If bListlar(i) Then
        ErrManba(i) = True
        Set acShMan = wbManba.Sheets(ListManba(i))
        With acShMan
            Rem Manbadagi asosiy va boshqa kerak schotlarning pozitsiyalarini aniqlab olamiz
            LastRow = .UsedRange.Row + .UsedRange.Rows.count - 1
            With .Range(.Cells(1, 1), .Cells(LastRow, 3))
                Wanted = "10000": FirstRow = .Find(What:=Wanted, LookIn:=xlValues, SearchOrder:=xlByColumns).Row - 1
                Wanted = "40000": LastRow = .Find(What:=Wanted, LookIn:=xlValues, SearchOrder:=xlByColumns).Row - 1
                QatorSoni = LastRow - FirstRow     'Massivning satrlari soni
                Wanted = "30000": JamiKapital = .Find(What:=Wanted, LookIn:=xlValues, SearchOrder:=xlByColumns).Row - FirstRow
                Wanted = "31206": Schot31206 = .Find(What:=Wanted, LookIn:=xlValues, SearchOrder:=xlByColumns).Row - FirstRow
                Wanted = "31200": Schot31200 = .Find(What:=Wanted, LookIn:=xlValues, SearchOrder:=xlByColumns).Row - FirstRow
            End With

            ReDim arTblSch(1 To QatorSoni, 1 To 3)  'Manbada 3ta ustunda joylashgan schotlar
            ReDim arJoriyQld(1 To QatorSoni, 1 To 5)    'Manbada joylashgan schotlarning qoldiqlari
            ReDim arJoriySch(1 To QatorSoni, 1 To 6)    'Schotlar va ularning mos qoldiqlarini birlashtiruvchi massiv

            arTblSch = .Range(.Cells(7, 1), .Cells(LastRow, 3)).Value
            arJoriyQld = .Range(.Cells(7, 4), .Cells(LastRow, 8)).Value
        End With

        For m = 1 To QatorSoni
            arJoriySch(m, 1) = CLng(IIf(arTblSch(m, 3) <> 0, arTblSch(m, 3), IIf(arTblSch(m, 2) <> 0, arTblSch(m, 2), arTblSch(m, 1)))) '3ta ustundagi schotlarni bittaga olamiz
            For j = 2 To 6
                arJoriySch(m, j) = IIf(arJoriySch(m, 1) < 20000, (-1) * arJoriyQld(m, j - 1), arJoriyQld(m, j - 1)) 'Umumiy massivni shakllantiramiz, bunda aktivlarning ishorasini o'zgartirib ketamiz
            Next j
        Next m

        Rem Yangi schot qo'shilganligiga tekshirish
        On Error Resume Next

        ReDim bYangiSchot(1 To QatorSoni)
        Rem Yangi schotlarni aniqlab olamiz
        For m = 1 To QatorSoni
            bYangiSchot(m) = WorksheetFunction.VLookup(arJoriySch(m, 1), arDoimiySch, 1, 0)
            If Err <> 0 Then
                ReDim Preserve arYangiSchot(1 To y(i) + 1)
                ReDim Preserve BeforeNew(1 To y(i) + 1)
                sYangiSchot(i) = sYangiSchot(i) & Chr(10) & Chr(32) & Chr(149) & Chr(32) & arJoriySch(m, 1)
                arYangiSchot(y(i) + 1) = arJoriySch(m, 1)
                BeforeNew(y(i) + 1) = arJoriySch(m - 1, 1)
                y(i) = y(i) + 1
                Err.Clear
                bNewAcc(i) = True
            End If
        Next m
        Rem Yangi schotlar bo'lsa, Msgbox orqali nima qilish kerakligini aniqlashtirib olamiz
        If bNewAcc(i) Then
            sSchot(i) = MsgNewAcc(i) & IIf(y(i) = 1, "hisobvaraq", "hisobvaraqlar")
            MsgNewAcc(i) = Left(ListJamlovchi(i), 2) & " balansida quyidagi yangi " & sSchot(i) & " aniqlandi:"
            MsgNewAcc(i) = MsgNewAcc(i) & Chr(10) & sYangiSchot(i) & Chr(10)
            MsgNewAcc(i) = MsgNewAcc(i) & vbNewLine & "Mazkur " & sSchot(i) & Chr(32) & ListJamlovchi(i) & " listiga ham qo'shilsinmi?"
            arNewAcc(i) = arYangiSchot
            Select Case MsgBox(MsgNewAcc(i), vbYesNoCancel + vbDefaultButton1 + vbExclamation, "Hisobvaraqlarda o'zgarish")
                Case vbYes      'Yangi schotlarni mos listga qo'shamiz
                    For j = 1 To y(i)
                        ReDim Preserve arYangidanOldin(1 To j)
                        With acShJam
                            BeforeNewRow = .Columns(1).Find(What:=BeforeNew(j), LookIn:=xlFormulas, SearchOrder:=xlByColumns).Row - 7
                            Rem Sikl orqali Yangi schotlar qaysi schotlardan keyin qo'yilishi kerakligini aniqlab olamiz
                            Do
                                arYangidanOldin(j) = WorksheetFunction.VLookup(CLng(Trim(arDoimiySch(BeforeNewRow, 1))), arDoimiySch, 1, 0)
                                BeforeNewRow = BeforeNewRow + 1
                            Loop While arYangidanOldin(j) < arYangiSchot(j)
                            arYangidanOldin(j) = IIf(Len(Trim(arDoimiySch(BeforeNewRow - 2, 1))) <> 0, _
                                                        arDoimiySch(BeforeNewRow - 2, 1), arDoimiySch(BeforeNewRow - 3, 1))
                            t = IIf(StrComp(Right(arYangiSchot(j), 2), "00", vbBinaryCompare) <> 0, 1, 2)   'Yangi schot balans yoki subschotligini aniqlashtiramiz
                            YangiSatr = .Columns(1).Find(What:=arYangidanOldin(j), LookIn:=xlFormulas, SearchOrder:=xlByColumns).Row + t
                            .Rows(YangiSatr & ":" & YangiSatr + t - 1).Insert
                            .Cells(YangiSatr, 1) = arYangiSchot(j)
                            Set rnNewAcc = .Range(.Cells(YangiSatr, 1), .Cells(YangiSatr, OldingiKunClm(i) + 4))
                            Rem Formatlarni to'g'irlaymiz
                            With rnNewAcc
                                If t = 1 Then     'Subschot qo'shilgan bo'lsa
                                    .Font.ThemeColor = xlThemeColorLight1
                                    .Font.Bold = False
                                    .Cells(1, 1).HorizontalAlignment = xlRight
                                    .Borders(xlEdgeRight).Weight = xlThin
                                    .Borders(xlInsideVertical).Weight = xlThin
                                Else              'Balansoviy schot qo'shilgan bo'lsa
                                    .Cells(1, 1).HorizontalAlignment = xlCenter
                                    .Resize(rnNewAcc.Rows.count + 1, rnNewAcc.Columns.count).Borders(xlEdgeRight).Weight = xlThin
                                    .Resize(rnNewAcc.Rows.count + 1, rnNewAcc.Columns.count).Borders(xlInsideVertical).Weight = xlThin
                                    .Font.Color = -65536
                                    .Font.TintAndShade = 0
                                    .Font.Bold = True
                                End If
                            End With
                            Rem Doimiy schotlarni saqlovchi massivni yangilaymiz
                            OxirgiQtr(i) = .Columns(2).Find(What:=SofFoyda, LookIn:=xlFormulas, SearchOrder:=xlByColumns).Row + 3
                            ReDim arDoimiySch(1 To OxirgiQtr(i) - 8, 1 To 1)
                            arDoimiySch = .Range(.Cells(8, 1), .Cells(OxirgiQtr(i) - 1, 1))
                        End With
                    Next
                Case vbNo
                    Rem Bunda kod shunchaki davom etadi
                Case vbCancel       'Makros tugatiladi
                    Unload frmStatusBar
                    Fayl = Right(Fayl, Len(Fayl) - InStrRev(Fayl, Application.PathSeparator))
                    If IsBookOpen(Fayl) Then wbManba.Close False
                    Unload Balans
                    Application.Calculation = xlCalculationAutomatic
                    Application.ScreenUpdating = True
                    Exit Sub
            End Select
        End If
        On Error GoTo ErrHandler

        Rem Aktiv va passiv o'rtasidagi farqni topib, sof foyda va boshqa bog'liq schotlarga qo'shib yuboramiz
        Farq = Application.VLookup(10000, arJoriySch, 2, 0) - Application.VLookup(20000, arJoriySch, 2, 0) - Application.VLookup(30000, arJoriySch, 2, 0)
        For h = 2 To 3
            arJoriySch(Schot31206, h) = arJoriySch(Schot31206, h) + Farq
            arJoriySch(Schot31200, h) = arJoriySch(Schot31200, h) + Farq
            arJoriySch(JamiKapital, h) = arJoriySch(JamiKapital, h) + Farq
        Next h

        Rem Manbadan olingan qoldiqlarni Jamlovchi balansga qo'yish uchun tayyorlab olamiz (oldingi "Yordamchi" fayl qilgan ish)
        ReDim arQoldiqlar(1 To OxirgiQtr(i) - 8, 1 To 5)
        For n = 1 To OxirgiQtr(i) - 8
            For j = 1 To 5
                arQoldiqlar(n, j) = IIf(Len(WorksheetFunction.Trim(arDoimiySch(n, 1))) <> 0, _
                    WorksheetFunction.IfError(Application.VLookup(arDoimiySch(n, 1), arJoriySch, j + 1, 0), 0), "")
            Next j
        Next n
        ErrManba(i) = False

        Rem Kategoriyalar orasida va eng oxirda joylashgan jamilarni shakllantiramiz
        Wanted = "АКТИВЫ, ВСЕГО": JamiAktiv = acShJam.Columns(2).Find(What:=Wanted, LookIn:=xlFormulas, SearchOrder:=xlByColumns).Row - 7
        Wanted = "ПАССИВЫ, ВСЕГО": JamiPassiv = acShJam.Columns(2).Find(What:=Wanted, LookIn:=xlFormulas, SearchOrder:=xlByColumns).Row - 7
        Jamilar = Array(JamiAktiv, JamiPassiv, OxirgiQtr(i) - 8)
        For j = 1 To 5
            arQoldiqlar(Jamilar(0), j) = Application.VLookup(10000, arJoriySch, j + 1, 0)
            arQoldiqlar(Jamilar(1), j) = Application.VLookup(20000, arJoriySch, j + 1, 0) + Application.VLookup(30000, arJoriySch, j + 1, 0)
            arQoldiqlar(Jamilar(2), j) = arQoldiqlar(Jamilar(1), j)
        Next j
        ManbaQoldiqlar(i) = arQoldiqlar
    End If
Next i
wbManba.Close False

For i = 0 To 1
    Call MyProgresBar
    Rem Oldingi kun uchun joy ochib, ma'lumotlarni ko'chirib qo'yamiz
    Set acShJam = wbJamlovchi.Sheets(ListJamlovchi(i))
    With acShJam
        If AlreadyDone = 0 Then
            Set TempRange = .Range(.Cells(8, OldingiKunClm(i)), .Cells(OxirgiQtr(i) - 1, OldingiKunClm(i) + 4))
            ReDim TempArray(1 To OxirgiQtr(i) - 8, 1 To 5)
            TempRange.Borders(xlEdgeRight).Weight = xlMedium
            TempRange.Borders(xlInsideVertical).Weight = xlThin
            TempArray = TempRange.Value
            TempRange.EntireColumn.Insert Shift:=xlToLeft
            .Range(.Cells(1, OldingiKunClm(i) + 5), .Cells(6, OldingiKunClm(i) + 9)).Copy .Cells(1, OldingiKunClm(i))
            TempRange.Offset(0, -5).Value = TempArray
            TempRange.Offset(-7, -5).Columns.AutoFit
            TempRange.Offset(-1, -5).Borders(xlEdgeRight).Weight = xlMedium
            TempRange.Offset(-1, -5).Borders(xlInsideVertical).Weight = xlThin
            .Cells(4, OldingiKunClm(i) + 5).Value = JoriySana
            .Range(.Cells(8, OldingiKunClm(i) + 10), .Cells(OxirgiQtr(i) - 1, OldingiKunClm(i) + 14)).FormulaR1C1 = "=RC[-5]-RC[-10]"   'Kunlik o'zgarishlarni aniqlaymiz
        End If
        If bListlar(i) Then .Range(.Cells(8, OldingiKunClm(i) + 5), .Cells(OxirgiQtr(i) - 1, OldingiKunClm(i) + 9)).Value = ManbaQoldiqlar(i)    'Tayyor  ma'lumotni balansga qo'yamiz
        If bNewAcc(i) Then
            .Cells(OxirgiQtr(i) + 2, OldingiKunClm(i) + 5) = "Yangi " & sSchot(i) & Chr(58)
            .Range(.Cells(OxirgiQtr(i) + 3, OldingiKunClm(i) + 5), .Cells(OxirgiQtr(i) + 2 + y(i), OldingiKunClm(i) + 5)).Value = Application.Transpose(arNewAcc(i))
        End If
    End With
Next i

If AlreadyDone = 0 Then    'КБ ёки ЦБ дан биттаси кайта ишланган бўлса иш якунланади
    Rem Jamlovchi balansning qolgan listlari bo'ylab siklni davom etamiz
    For i = i To UBound(ListJamlovchi) - 4
        Call MyProgresBar
        If i <> 4 Then b = 1 Else b = 3     ' "КП" listiga 3ta ustun qo'shish uchun
        Set acShJam = wbJamlovchi.Sheets(ListJamlovchi(i))
        With acShJam
            Wanted = "За день"
            Set OxirgiKatak = .Cells.Find(What:=Wanted, LookIn:=xlFormulas, SearchOrder:=xlByRows)
            OxirgiUstun = OxirgiKatak.Column
            UzgarRow = OxirgiKatak.Row
            LastRow = .UsedRange.Row + .UsedRange.Rows.count - 1
            .Range(.Cells(1, OxirgiUstun - b), .Cells(1, OxirgiUstun - 1)).EntireColumn.Insert
            ReDim arFrm(1 To LastRow, 1 To b)  'Joriy sanadan oldingi sanaga formulalarni ko'chirish uchun massiv
            arFrm = .Range(.Cells(1, OxirgiUstun), .Cells(LastRow, OxirgiUstun + b - 1)).FormulaR1C1
            Call MyProgresBar
            Rem Olingan formulalarni o'tgan sanaga moslab chiqamiz
            For n = 1 To (0.5 * b + 0.5)   ' f(3)=2, f(1)=1; bu faqat "КП" listida 2ta ustunni o'zgartirish k-k bo'lgani uchun ishlatiladi
                For m = 1 To LastRow
                    If InStr(arFrm(m, n), "C[") <> 0 Then   'faqat formulasi o'zgartirilishi kerak bo'lgan elementlargina qayta ishlanadi
                        ForStart = 1
                        Do
                            ForStart = InStr(ForStart, arFrm(m, n), "C[") + 1
                            ForEnd = InStr(ForStart + 1, arFrm(m, n), "]") - 1
                            ForLen = ForEnd - ForStart
                            NewClm = CLng(Mid(arFrm(m, n), ForStart + 1, ForLen))       ' pastda: agar formula "КП" listidan olinsa bor for.dan 2 ayriladi
                            NewClm = IIf(NewClm > 100, NewClm + IIf(InStr(arFrm(m, n), "КП!") = 0, b - 5, -2), NewClm) 'consider rewriting: very hard coded instruction; f(3)=2, f(1)=4
                            arFrm(m, n) = Left(arFrm(m, n), ForStart) & NewClm & Right(arFrm(m, n), Len(arFrm(m, n)) - ForEnd)
                        Loop While InStr(ForEnd + 1, arFrm(m, n), "C[") <> 0    'bitta elementdagi o'zgartirilishi k-k hamma formulalar o'zgarmaguncha sikl davom etadi
                    End If
                Next m
            Next n

            .Range(.Cells(1, OxirgiUstun - b), .Cells(LastRow, OxirgiUstun - 1)).Value = arFrm     'tayyor bo'lgan formulalarni o'tgan sanaga qo'yamiz
            .Range(.Cells(1, OxirgiUstun - b * 2), .Cells(1, OxirgiUstun - b - 1)).EntireColumn.Hidden = True      ' o'tgan sanadan oldingi ustunni yashiramiz

            If .Name Like "СТРУК_Норкулов" Then
                .Cells(2, OxirgiUstun - 1).Value = ""
                .Cells(3, OxirgiUstun).Value = "на" & Chr(10) & ExceptDaysOffPlus(JoriySana + 1) & "г."  'Chr(10)= "^p"
            End If

            If (.Name Like "Таблица 5") Or (.Name Like "Таблица 9 (Раис)") Then
                If .Name Like "Таблица 5" Then f = 1 Else f = 2
                ReDim arFoizUzg(1 To LastRow - UzgarRow - 1, 1 To 1)
                arFoizUzg = .Range(.Cells(UzgarRow + f, OxirgiUstun + 2), .Cells(LastRow, OxirgiUstun + 2)).FormulaR1C1
                For m = 1 To LastRow - UzgarRow - 1
                    arFoizUzg(m, 1) = IIf(InStr(arFoizUzg(m, 1), "=") <> 0, "=RC[-2]/RC[-3]-1", arFoizUzg(m, 1))
                Next m
                .Range(.Cells(UzgarRow + f, OxirgiUstun + 2), .Cells(LastRow, OxirgiUstun + 2)).Value = arFoizUzg
            End If

            If i >= 9 And i <= 15 Then
                RowN = .Cells(UzgarRow - 1, OxirgiUstun - 1).End(xlDown).Row - 1
                .Range(.Cells(UzgarRow - 1, OxirgiUstun - 1), .Cells(RowN, OxirgiUstun - 1)).Merge
            End If
            
            If .Name Like "Таблица 3" Then
                RowN = .Cells(UzgarRow - 1 + 16, OxirgiUstun - 1).End(xlDown).Row - 1
                .Range(.Cells(UzgarRow - 1 + 16, OxirgiUstun - 1), .Cells(RowN, OxirgiUstun - 1)).Merge
            End If
            
            If .Name Like "КЭ_конс (2)" Then .Cells(2, OxirgiUstun - 1).Value = ""

            ReDim Uzgarish(1 To LastRow - UzgarRow - 1, 1 To 1)     'kunlik o'zgarishlarni qo'yish uchun massiv
            Uzgarish = .Range(.Cells(UzgarRow + 1, OxirgiUstun + b), .Cells(LastRow, OxirgiUstun + b)).FormulaR1C1
            For m = 1 To LastRow - UzgarRow - 1
                Uzgarish(m, 1) = IIf(InStr(Uzgarish(m, 1), "=") <> 0, "=RC[-" & (0.5 * b + 0.5) & "]-RC[-" & (1.5 * b + 0.5) & "]", Uzgarish(m, 1)) 'f(3)=2, f(1)=1; f(3)=5, f(1)=2
            Next m                  'Len(WorksheetFunction.Trim(Uzgarish(m, 1))) <> 0
            .Range(.Cells(UzgarRow + 1, OxirgiUstun + b), .Cells(LastRow, OxirgiUstun + b)).Value = Uzgarish
        End With
    Next i
    Call MyProgresBar
End If
Unload frmStatusBar
Finish = Timer

'If OyUzgarishi Then
'    MsgBox "Bugun oyning birinchi ish kuni uchun balans tayyorlandi. " & vbNewLine & _
'      "Oylik o'zgarishlarni to'g'rilab qo'yishni unitmang!", vbInformation, "Eslatma"
'End If

With Application
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
    .StatusBar = Format(Finish - Start, "0.00") & " soniyada muvaffaqiyatli bajarildi!"
    .OnTime Now + TimeValue("00:00:06"), "KillStatBar"
End With
Exit Sub

ErrHandler:
    ErrMsg = "Xatolik sodir bo'ldi!" & Chr(10) & vbNewLine
    Fayl = Right(Fayl, Len(Fayl) - InStrRev(Fayl, Application.PathSeparator))
    If Not ErrManba(i) Then
        XatoList = ListJamlovchi(i)
    Else
        ErrMsg = ErrMsg & "Fayl: " & Chr(34) & wbManba.Name & Chr(34) & Chr(10)
        XatoList = ListManba(i)
    End If
    ErrMsg = ErrMsg & "List: " & Chr(34) & XatoList & Chr(34) & Chr(10)
    Select Case Err
        Case 9
            ErrMsg = ErrMsg & "List o'chirilgan yoki nomi o'zgartirilgan" & Chr(10)
        Case 91
            ErrMsg = ErrMsg & "Topilmadi: " & Chr(34) & Wanted & Chr(34) & " so'zi" & Chr(10)
    End Select
    ErrMsg = ErrMsg & Chr(10) & "Xatolik tavsifi: " & Chr(10) & Err.Description & Chr(10) & vbNewLine
    If Not IsBookOpen(Fayl) Then
        ErrMsg = ErrMsg & "Hozir kitob saqlanmasdan yopilib ochiladi. " & Chr(10) & "Barcha ma'lumotlar to'g'ri "
        ErrMsg = ErrMsg & "kiritilganligiga va kitobda hech qanday tarkibiy o'zgartirishlar bo'lmaganligiga "
        ErrMsg = ErrMsg & "ishonch hosil qilib, keyin Makrosni qaytadan ishga tushiring!" & Chr(10) & Chr(10)
        ErrMsg = ErrMsg & "Kitobning yopilib ochilishi biroz vaqt olishi mumkin. " & Chr(10)
        ErrMsg = ErrMsg & "Iltimos ozgina kutib turing!"
        MsgBox ErrMsg, vbCritical + vbMsgBoxHelpButton + vbDefaultButton1, "Xatolik", Err.HelpFile, Err.HelpContext
        On Error Resume Next
        Unload frmStatusBar
        Unload Balans
        ufrmKuting.Show 0
        JamFayl = wbJamlovchi.Path & Application.PathSeparator & wbJamlovchi.Name
        wbJamlovchi.Close False
        Workbooks.Open JamFayl, UpdateLinks:=False ', Password:="242"
        Unload ufrmKuting
    Else
        MsgBox ErrMsg, vbCritical + vbMsgBoxHelpButton + vbDefaultButton1, "Xatolik", Err.HelpFile, Err.HelpContext
        wbManba.Close False
        On Error Resume Next
        Unload frmStatusBar
        Unload Balans
    End If
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .StatusBar = "Makros oxirigacha bajarila olmadi. Faylda nazarada tutilmagan o'zgartirish yoki xatolik mavjud."
        .OnTime Now + TimeValue("00:00:25"), "KillStatBar"
    End With
End Sub

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

 Private Sub KillStatBar()
     Application.StatusBar = False
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







