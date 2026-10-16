Attribute VB_Name = "Konvertor"
'==========================================================================
'  Lotin-Kirill konvertor (Microsoft Excel uchun)
'  O'zbek tili: lotin <-> kirill yozuvi
'
'  Muallif / Author: Begzodbek Akramjonov
'  Uzbek Latin <-> Cyrillic converter for Microsoft Excel
'
'  - Belgilangan kataklardagi matnni o'giradi (formulalar va sonlar
'    o'zgarmaydi).
'  - =KIRILL(A1) va =LOTIN(A1) formulalari.
'  - "Ortga qaytarish" (Ctrl+Z ham ishlaydi).
'
'  Barcha kirill harflari ChrW() orqali yozilgan, shuning uchun kod
'  Windows'ning har qanday til sozlamasida bir xil ishlaydi.
'==========================================================================
Option Explicit

Private Const SARLAVHA As String = "Lotin-Kirill konvertor"
Private Const MUALLIF As String = "Begzodbek Akramjonov"
Private Const ILOVA As String = "UzLotinKirill"

' --- Ortga qaytarish ma'lumotlari ---
Private mUndoWs As Object
Private mUndoAddr() As String
Private mUndoOld() As String
Private mUndoNew() As String
Private mUndoPre() As Boolean
Private mUndoCount As Long

' --- Matn yig'uvchi bufer ---
Private mBuf As String
Private mLen As Long

' --- Apostrof sozlamasi ---
Private mAposYuklandi As Boolean
Private mOddiy As Boolean
Private mOg As String
Private mTq As String

' --- Istisno so'zlar ---
Private mExKey() As String
Private mExVal() As String
Private mExWhole() As Boolean
Private mExCount As Long

' --- Katta lug'at (qo'shimchaning yashirin varag'ida) ---
Private mDict As Collection
Private mDictPre As Collection
Private mDictCyr() As String
Private mDictWhole() As Boolean
Private mDictYuklandi As Boolean

'==========================================================================
'  LENTA (Ribbon) VA MENYU TUGMALARI
'==========================================================================
Public Sub Rb_LotinKirill(control As Object)
    LotinKirill
End Sub

Public Sub Rb_KirillLotin(control As Object)
    KirillLotin
End Sub

Public Sub Rb_Qaytar(control As Object)
    OrtgaQaytar
End Sub

Public Sub Rb_Yordam(control As Object)
    Yordam
End Sub

Public Sub Rb_AposBos(control As Object, pressed As Boolean)
    mOddiy = pressed
    BelgilarniSozla
    mAposYuklandi = True
    On Error Resume Next
    If pressed Then
        SaveSetting ILOVA, "Sozlama", "OddiyApostrof", "1"
    Else
        SaveSetting ILOVA, "Sozlama", "OddiyApostrof", "0"
    End If
End Sub

Public Sub Rb_AposHolat(control As Object, ByRef returnedVal)
    ApostrofYukla
    returnedVal = mOddiy
End Sub

'==========================================================================
'  ASOSIY BUYRUQLAR
'==========================================================================
Public Sub LotinKirill()
    KonvertQil True
End Sub

Public Sub KirillLotin()
    KonvertQil False
End Sub

Public Sub Yordam()
    MsgBox "Lotin-Kirill konvertor" & vbCrLf & _
        "Muallif / Author: " & MUALLIF & vbCrLf & vbCrLf & _
        "Kerakli kataklarni belgilang (butun varaq uchun Ctrl+A)," & vbCrLf & _
        "keyin 'Lotin -> Kirill' yoki 'Kirill -> Lotin' tugmasini bosing." & vbCrLf & _
        "Bu tugmalar sichqonchaning o'ng tugmasi menyusida ham bor." & vbCrLf & vbCrLf & _
        "Dastur faqat matnni o'zgartiradi. Formulalar, sonlar va" & vbCrLf & _
        "sanalar o'z holicha qoladi." & vbCrLf & vbCrLf & _
        "Adashib bosib yuborsangiz, 'Ortga qaytarish' tugmasi" & vbCrLf & _
        "yoki Ctrl+Z hammasini avvalgi holiga qaytaradi." & vbCrLf & vbCrLf & _
        "Formula bilan ham bo'ladi: =KIRILL(A1) yoki =LOTIN(A1).", _
        vbInformation, SARLAVHA
End Sub

Private Sub KonvertQil(ByVal toCyr As Boolean)
    Dim ws As Object, rng As Object, tr As Object, ar As Object, c As Object
    Dim v As Variant, f As Variant
    Dim r As Long, k As Long, qadam As Long
    Dim eski As String, yangi As String, nomi As String, xabar As String
    Dim jami As Double, bajarildi As Double
    Dim ozgardi As Long, otkazildi As Long
    Dim tekshir As Boolean, ok As Boolean, pre As Boolean, spXato As Boolean
    Dim calcEski As Long, scrEski As Boolean, evEski As Boolean, saqlandi As Boolean

    If toCyr Then nomi = "Lotin -> Kirill" Else nomi = "Kirill -> Lotin"

    If ActiveWorkbook Is Nothing Then
        MsgBox "Avval Excel faylini oching.", vbInformation, SARLAVHA
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        If ShakllarniOgir(toCyr) Then Exit Sub
        MsgBox "Avval o'zgartiriladigan kataklarni belgilang.", vbInformation, SARLAVHA
        Exit Sub
    End If

    Set ws = Selection.Parent
    Set rng = Selection

    ' Bitta bo'sh (yoki matnsiz) katak belgilangan bo'lsa - butun varaqni taklif qilamiz
    If rng.Cells.CountLarge = 1 Then
        If rng.HasFormula Or VarType(rng.Value) <> vbString Then
            If MsgBox("Faqat bitta katak belgilangan va unda matn yo'q." & vbCrLf & _
                      "Butun varaqdagi matnlarni o'zgartiraymi?", _
                      vbQuestion + vbYesNo, SARLAVHA) <> vbYes Then Exit Sub
            Set rng = ws.UsedRange
        End If
    End If

    If rng.Cells.CountLarge > 1 Then Set rng = Application.Intersect(rng, ws.UsedRange)
    If rng Is Nothing Then
        MsgBox "Belgilangan joyda matn topilmadi.", vbInformation, SARLAVHA
        Exit Sub
    End If

    On Error GoTo Xato
    scrEski = Application.ScreenUpdating
    evEski = Application.EnableEvents
    calcEski = Application.Calculation
    saqlandi = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Matnli kataklarni topish
    Set tr = Nothing
    If rng.Cells.CountLarge = 1 Then
        Set tr = rng
        tekshir = True
    Else
        On Error Resume Next
        Set tr = rng.SpecialCells(xlCellTypeConstants, xlTextValues)
        spXato = (Err.Number <> 0)
        Err.Clear
        On Error GoTo Xato
        If spXato Then
            Set tr = Nothing
            ' Himoyalangan varaqda SpecialCells ishlamasligi mumkin - har bir katakni tekshiramiz
            If ws.ProtectContents Then
                Set tr = rng
                tekshir = True
            End If
        End If
    End If

    UndoBoshla ws

    If Not tr Is Nothing Then
        jami = tr.Cells.CountLarge
        If jami < 1 Then jami = 1
        For Each ar In tr.Areas
            If ar.Cells.CountLarge = 1 Then
                ReDim v(1 To 1, 1 To 1)
                v(1, 1) = ar.Value
                If tekshir Then
                    ReDim f(1 To 1, 1 To 1)
                    f(1, 1) = ar.Formula
                End If
            Else
                v = ar.Value
                If tekshir Then f = ar.Formula
            End If
            For r = 1 To UBound(v, 1)
                For k = 1 To UBound(v, 2)
                    bajarildi = bajarildi + 1
                    If VarType(v(r, k)) = vbString Then
                        ok = True
                        If tekshir Then ok = (Left$(CStr(f(r, k)), 1) <> "=")
                        If ok Then
                            eski = v(r, k)
                            If toCyr Then
                                yangi = LotinDanKirillga(eski)
                            Else
                                yangi = KirillDanLotinga(eski)
                            End If
                            If yangi <> eski Then
                                Set c = ar.Cells(r, k)
                                pre = PrefiksBor(c)
                                If YozKatak(c, yangi, pre) Then
                                    UndoQosh c.Address(False, False), eski, yangi, pre
                                    ozgardi = ozgardi + 1
                                Else
                                    otkazildi = otkazildi + 1
                                End If
                            End If
                        End If
                    End If
                    qadam = qadam + 1
                    If qadam >= 500 Then
                        qadam = 0
                        Application.StatusBar = nomi & ": " & Format(bajarildi / jami, "0%")
                    End If
                Next k
            Next r
        Next ar
    End If

    Application.Calculation = calcEski
    Application.EnableEvents = evEski
    Application.ScreenUpdating = scrEski
    Application.StatusBar = False
    saqlandi = False

    If otkazildi > 0 Then
        xabar = otkazildi & " ta katakni o'zgartirib bo'lmadi (varaq yoki katak himoyalangan bo'lishi mumkin)."
        If ozgardi > 0 Then xabar = ozgardi & " ta katak o'zgartirildi." & vbCrLf & xabar
        MsgBox xabar, vbExclamation, SARLAVHA
    ElseIf ozgardi = 0 Then
        MsgBox "Belgilangan kataklarda o'zgartiriladigan matn topilmadi.", vbInformation, SARLAVHA
    End If

    If ozgardi > 0 Then
        Application.StatusBar = nomi & ": " & ozgardi & " ta katak o'zgartirildi."
        StatusniKeyinTozala
        On Error Resume Next
        Application.OnUndo "Lotin-Kirill: ortga qaytarish", "'" & ThisWorkbook.Name & "'!OrtgaQaytar"
    End If
    Exit Sub

Xato:
    xabar = Err.Description
    On Error Resume Next
    If saqlandi Then
        Application.Calculation = calcEski
        Application.EnableEvents = evEski
        Application.ScreenUpdating = scrEski
    End If
    Application.StatusBar = False
    If mUndoCount > 0 Then
        Application.OnUndo "Lotin-Kirill: ortga qaytarish", "'" & ThisWorkbook.Name & "'!OrtgaQaytar"
    End If
    MsgBox "Kutilmagan xatolik: " & xabar & vbCrLf & _
           "O'zgartirilgan kataklarni 'Ortga qaytarish' bilan tiklash mumkin.", vbExclamation, SARLAVHA
End Sub

Private Function PrefiksBor(ByVal c As Object) As Boolean
    On Error Resume Next
    PrefiksBor = (Len(c.PrefixCharacter) > 0)
End Function

' Katakka matn yozish: matn son, sana, formula yoki TRUE/FALSE bo'lib
' qolmasligi uchun kerak bo'lsa oldiga apostrof qo'yiladi.
Private Function YozKatak(ByVal c As Object, ByVal t As String, ByVal oldPrefix As Boolean) As Boolean
    On Error GoTo Xato
    If CStr(c.NumberFormat) = "@" Then
        c.Value = t
    ElseIf oldPrefix Or PrefiksKerak(t) Then
        c.Value = "'" & t
    Else
        c.Value = t
    End If
    YozKatak = True
    Exit Function
Xato:
    YozKatak = False
End Function

Private Function PrefiksKerak(ByVal t As String) As Boolean
    Dim b As String
    If Len(t) = 0 Then Exit Function
    b = Left$(t, 1)
    If b = "=" Or b = "+" Or b = "-" Or b = "@" Or b = "'" Then
        PrefiksKerak = True
        Exit Function
    End If
    On Error Resume Next
    If IsNumeric(t) Then
        PrefiksKerak = True
    ElseIf IsDate(t) Then
        PrefiksKerak = True
    ElseIf UCase$(Trim$(t)) = "TRUE" Or UCase$(Trim$(t)) = "FALSE" Then
        PrefiksKerak = True
    ElseIf Right$(t, 1) = "%" And Len(t) > 1 Then
        If IsNumeric(Left$(t, Len(t) - 1)) Then PrefiksKerak = True
    End If
End Function

Private Function ShakllarniOgir(ByVal toCyr As Boolean) As Boolean
    Dim shp As Object, n As Long, t As String
    On Error GoTo Tamom
    For Each shp In Selection.ShapeRange
        If shp.HasTextFrame Then
            If shp.TextFrame2.HasText Then
                t = shp.TextFrame2.TextRange.Text
                If toCyr Then
                    shp.TextFrame2.TextRange.Text = LotinDanKirillga(t)
                Else
                    shp.TextFrame2.TextRange.Text = KirillDanLotinga(t)
                End If
                n = n + 1
            End If
        End If
    Next shp
Tamom:
    ShakllarniOgir = (n > 0)
End Function

'==========================================================================
'  ORTGA QAYTARISH
'==========================================================================
Private Sub UndoBoshla(ByVal ws As Object)
    Set mUndoWs = ws
    mUndoCount = 0
    ReDim mUndoAddr(1 To 256)
    ReDim mUndoOld(1 To 256)
    ReDim mUndoNew(1 To 256)
    ReDim mUndoPre(1 To 256)
End Sub

Private Sub UndoQosh(ByVal addr As String, ByVal eski As String, ByVal yangi As String, ByVal pre As Boolean)
    Dim m As Long
    mUndoCount = mUndoCount + 1
    If mUndoCount > UBound(mUndoAddr) Then
        m = UBound(mUndoAddr) * 2
        ReDim Preserve mUndoAddr(1 To m)
        ReDim Preserve mUndoOld(1 To m)
        ReDim Preserve mUndoNew(1 To m)
        ReDim Preserve mUndoPre(1 To m)
    End If
    mUndoAddr(mUndoCount) = addr
    mUndoOld(mUndoCount) = eski
    mUndoNew(mUndoCount) = yangi
    mUndoPre(mUndoCount) = pre
End Sub

Public Sub OrtgaQaytar()
    Dim i As Long, n As Long, c As Object, nm As String, xabar As String
    Dim calcEski As Long, scrEski As Boolean, evEski As Boolean, saqlandi As Boolean

    If mUndoCount = 0 Or mUndoWs Is Nothing Then
        MsgBox "Ortga qaytariladigan konvertatsiya yo'q.", vbInformation, SARLAVHA
        Exit Sub
    End If

    On Error Resume Next
    nm = mUndoWs.Name
    If Err.Number <> 0 Then
        On Error GoTo 0
        mUndoCount = 0
        Set mUndoWs = Nothing
        MsgBox "Varaq yopilgan yoki o'chirilgan, ortga qaytarib bo'lmaydi.", vbInformation, SARLAVHA
        Exit Sub
    End If

    On Error GoTo Xato
    scrEski = Application.ScreenUpdating
    evEski = Application.EnableEvents
    calcEski = Application.Calculation
    saqlandi = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For i = mUndoCount To 1 Step -1
        Set c = mUndoWs.Range(mUndoAddr(i))
        ' Faqat konvertatsiyadan keyin qo'lda o'zgartirilmagan kataklar tiklanadi
        If Not c.HasFormula Then
            If VarType(c.Value) = vbString Then
                If c.Value = mUndoNew(i) Then
                    If YozKatak(c, mUndoOld(i), mUndoPre(i)) Then n = n + 1
                End If
            End If
        End If
    Next i

    Application.Calculation = calcEski
    Application.EnableEvents = evEski
    Application.ScreenUpdating = scrEski
    saqlandi = False
    mUndoCount = 0
    Set mUndoWs = Nothing
    Application.StatusBar = n & " ta katak avvalgi holatiga qaytarildi."
    StatusniKeyinTozala
    Exit Sub

Xato:
    xabar = Err.Description
    On Error Resume Next
    If saqlandi Then
        Application.Calculation = calcEski
        Application.EnableEvents = evEski
        Application.ScreenUpdating = scrEski
    End If
    MsgBox "Ortga qaytarishda xatolik: " & xabar, vbExclamation, SARLAVHA
End Sub

Private Sub StatusniKeyinTozala()
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 6), "'" & ThisWorkbook.Name & "'!StatusTozala"
End Sub

Public Sub StatusTozala()
    On Error Resume Next
    Application.StatusBar = False
End Sub

'==========================================================================
'  EXCEL FORMULALARI:  =KIRILL(A1)   =LOTIN(A1)
'==========================================================================
Public Function KIRILL(Matn As Variant) As Variant
    KIRILL = UdfOgir(Matn, True)
End Function

Public Function LOTIN(Matn As Variant) As Variant
    LOTIN = UdfOgir(Matn, False)
End Function

Private Function UdfOgir(ByVal v As Variant, ByVal toCyr As Boolean) As Variant
    Dim a As Variant, r As Long, k As Long
    On Error GoTo Xato
    a = v                       ' katak (Range) berilsa, uning qiymati olinadi
    If IsArray(a) Then
        If OlchamSoni(a) >= 2 Then
            For r = LBound(a, 1) To UBound(a, 1)
                For k = LBound(a, 2) To UBound(a, 2)
                    a(r, k) = BittaOgir(a(r, k), toCyr)
                Next k
            Next r
        Else
            For r = LBound(a) To UBound(a)
                a(r) = BittaOgir(a(r), toCyr)
            Next r
        End If
        UdfOgir = a
    Else
        UdfOgir = BittaOgir(a, toCyr)
    End If
    Exit Function
Xato:
    UdfOgir = CVErr(xlErrValue)
End Function

Private Function OlchamSoni(a As Variant) As Long
    Dim n As Long, d As Long
    On Error GoTo Tamom
    Do
        d = UBound(a, n + 1)
        n = n + 1
    Loop While n < 60
Tamom:
    OlchamSoni = n
End Function

Private Function BittaOgir(ByVal x As Variant, ByVal toCyr As Boolean) As Variant
    If IsError(x) Then
        BittaOgir = x
    ElseIf IsEmpty(x) Then
        BittaOgir = ""
    ElseIf VarType(x) = vbString Then
        If toCyr Then
            BittaOgir = LotinDanKirillga(CStr(x))
        Else
            BittaOgir = KirillDanLotinga(CStr(x))
        End If
    Else
        BittaOgir = x
    End If
End Function

'==========================================================================
'  KIRILL -> LOTIN
'==========================================================================
Public Function KirillDanLotinga(ByVal s As String) As String
    Dim n As Long, i As Long, c As Long, lc As Long
    Dim pv As Long, nx As Long, t As String
    n = Len(s)
    If n = 0 Then Exit Function
    ApostrofYukla
    BufStart n
    For i = 1 To n
        c = KodAt(s, i)
        If IsCyr(c) Then
            lc = CyrLow(c)
            pv = CyrLow(KodAt(s, i - 1))
            nx = CyrLow(KodAt(s, i + 1))
            t = KirillHarf(c, lc, pv, nx)
            If lc <> c And Len(t) > 0 Then t = LotinShakl(t, s, i)
            BufAdd t
        Else
            BufAdd Mid$(s, i, 1)
        End If
    Next i
    KirillDanLotinga = BufText()
End Function

Private Function KirillHarf(ByVal c As Long, ByVal lc As Long, ByVal pv As Long, ByVal nx As Long) As String
    Dim t As String
    Select Case lc
        Case 1072: t = "a"
        Case 1073: t = "b"
        Case 1074: t = "v"
        Case 1075: t = "g"
        Case 1076: t = "d"
        Case 1077                                   ' ye / e
            If Not HarfMi(pv) Or KirUnli(pv) Or pv = 1098 Or pv = 1100 Then
                t = "ye"
            Else
                t = "e"
            End If
        Case 1105: t = "yo"
        Case 1078: t = "j"
        Case 1079: t = "z"
        Case 1080: t = "i"
        Case 1081: t = "y"
        Case 1082: t = "k"
        Case 1083: t = "l"
        Case 1084: t = "m"
        Case 1085: t = "n"
        Case 1086: t = "o"
        Case 1087: t = "p"
        Case 1088: t = "r"
        Case 1089: t = "s"
        Case 1090: t = "t"
        Case 1091: t = "u"
        Case 1092: t = "f"
        Case 1093: t = "x"
        Case 1094                                   ' ts / s
            If KirUnli(pv) And HarfMi(nx) Then t = "ts" Else t = "s"
        Case 1095: t = "ch"
        Case 1096, 1097: t = "sh"
        Case 1098                                   ' tutuq belgisi
            If pv = 1118 Then
                t = ""
            ElseIf (nx = 1077 Or nx = 1105 Or nx = 1102 Or nx = 1103) And KirUndosh(pv) Then
                t = ""
            Else
                t = mTq
            End If
        Case 1099: t = "i"
        Case 1100                                   ' yumshatish belgisi
            If nx = 1086 Then t = "y" Else t = ""
        Case 1101: t = "e"
        Case 1102: t = "yu"
        Case 1103: t = "ya"
        Case 1118: t = "o" & mOg
        Case 1171: t = "g" & mOg
        Case 1179: t = "q"
        Case 1203                                   ' h (s'h)
            If pv = 1089 Then t = mTq & "h" Else t = "h"
        Case Else
            t = ChrW(c)
    End Select
    KirillHarf = t
End Function

' Katta harf: so'z to'liq katta harfda bo'lsa SH, aks holda Sh
Private Function LotinShakl(ByVal t As String, ByRef s As String, ByVal i As Long) As String
    Dim nxr As Long, hammasi As Boolean
    If Len(t) = 1 Then
        LotinShakl = UCase$(t)
        Exit Function
    End If
    nxr = KodAt(s, i + 1)
    If HarfMi(nxr) Then
        hammasi = KattaMi(nxr)
    Else
        hammasi = KattaMi(KodAt(s, i - 1))
    End If
    If hammasi Then
        LotinShakl = UCase$(t)
    Else
        LotinShakl = BoshHarfKatta(t)
    End If
End Function

Private Function BoshHarfKatta(ByVal t As String) As String
    Dim i As Long, c As Long
    For i = 1 To Len(t)
        c = KodAt(t, i)
        If c >= 97 And c <= 122 Then
            BoshHarfKatta = Left$(t, i - 1) & UCase$(Mid$(t, i, 1)) & Mid$(t, i + 1)
            Exit Function
        End If
    Next i
    BoshHarfKatta = t
End Function

'==========================================================================
'  LOTIN -> KIRILL
'==========================================================================
Public Function LotinDanKirillga(ByVal s As String) As String
    Dim n As Long, i As Long, c As Long, used As Long
    Dim lastK As Integer      ' 0 = so'z boshi, 1 = unli, 2 = undosh
    n = Len(s)
    If n = 0 Then Exit Function
    IstisnoYukla
    BufStart n
    lastK = 0
    i = 1
    Do While i <= n
        c = KodAt(s, i)
        If IsLat(c) Then
            used = 0
            If lastK = 0 Then used = SozBoshi(s, i, n, lastK)
            If used = 0 Then used = LotinHarf(s, i, lastK)
            i = i + used
        ElseIf IsApos(c) Then
            If IsLat(KodAt(s, i - 1)) And IsLat(KodAt(s, i + 1)) Then
                ' tutuq belgisi
                If IsLatU(KodAt(s, i - 1)) And IsLatU(KodAt(s, i + 1)) Then
                    BufAdd ChrW(1066)
                Else
                    BufAdd ChrW(1098)
                End If
            Else
                BufAdd Mid$(s, i, 1)
                lastK = 0
            End If
            i = i + 1
        Else
            BufAdd Mid$(s, i, 1)
            If IsCyr(c) Then
                If KirUnli(CyrLow(c)) Then lastK = 1 Else lastK = 2
            Else
                lastK = 0
            End If
            i = i + 1
        End If
    Loop
    LotinDanKirillga = BufText()
End Function

Private Function LotinHarf(ByRef s As String, ByVal i As Long, ByRef lastK As Integer) As Long
    Dim c As Long, lc As Long, up As Boolean
    Dim n1 As Long, n2 As Long, p As Long, r As Long, used As Long, unli As Boolean
    c = KodAt(s, i)
    up = IsLatU(c)
    lc = LowLat(c)
    n1 = LowLat(KodAt(s, i + 1))
    n2 = LowLat(KodAt(s, i + 2))
    p = LowLat(KodAt(s, i - 1))
    used = 1
    r = 0
    Select Case lc
        Case 97: r = 1072: unli = True                      ' a
        Case 101                                            ' e
            ' undoshdan keyin va "ie" (klient, patsient) da - kirillcha e, so'z boshida va unlidan keyin - teskari e
            If lastK = 2 Or p = 105 Then r = 1077 Else r = 1101
            unli = True
        Case 105: r = 1080: unli = True                     ' i
        Case 111                                            ' o, o'
            unli = True
            If IsApos(n1) Then
                r = 1118: used = 2
            Else
                r = 1086
            End If
        Case 117: r = 1091: unli = True                     ' u
        Case 99                                             ' ch
            If n1 = 104 Then r = 1095: used = 2
        Case 103                                            ' g, g'
            If IsApos(n1) Then
                r = 1171: used = 2
            Else
                r = 1075
            End If
        Case 115                                            ' s, sh, s'h
            If n1 = 104 Then
                r = 1096: used = 2
            ElseIf IsApos(n1) And n2 = 104 Then
                Chiq 1089, up
                Chiq 1203, IsLatU(KodAt(s, i + 2))
                used = 3: r = -1
            ElseIf SniCgaOgir(s, i, p) Then
                r = 1094
            Else
                r = 1089
            End If
        Case 116                                            ' t, ts
            ' tsex, tsirk (so'z boshi), abzats, prints (so'z oxiri) - o'zbekcha so'zlar
            ' ts bilan boshlanmaydi va tugamaydi; MTS, ATS kabi qisqartmalar o'zgarmaydi
            If n1 = 115 And TsQoida(s, i + 2) Then
                r = 1094: used = 2
            ElseIf n1 = 115 And KodAt(s, i + 1) = 115 And (lastK = 0 Or Not HarfMi(KodAt(s, i + 2))) Then
                r = 1094: used = 2
            Else
                r = 1090
            End If
        Case 121                                            ' y, yo, yu, ya, ye
            If n1 = 111 And Not IsApos(n2) Then
                r = 1105: used = 2: unli = True
            ElseIf n1 = 117 Then
                r = 1102: used = 2: unli = True
            ElseIf n1 = 97 Then
                r = 1103: used = 2: unli = True
            ElseIf n1 = 101 Then
                used = 2: unli = True
                If lastK = 2 And Not IsApos(KodAt(s, i - 1)) Then
                    Chiq 1098, up
                    Chiq 1077, IsLatU(KodAt(s, i + 1))
                    r = -1
                Else
                    r = 1077
                End If
            Else
                r = 1081
            End If
        Case 119: r = 0                                     ' w - o'zgarmaydi
        Case Else
            r = OddiyKod(lc)
    End Select
    If r > 0 Then
        Chiq r, up
    ElseIf r = 0 Then
        BufAdd Mid$(s, i, used)
    End If
    If unli Then lastK = 1 Else lastK = 2
    LotinHarf = used
End Function

' aksiya, funksiya, korrupsiya, stansiya, potensial, distansion ... -> ts
Private Function SniCgaOgir(ByRef s As String, ByVal i As Long, ByVal p As Long) As Boolean
    Dim t As String
    t = LCase$(Mid$(s, i + 1, 3))
    If p = 107 Or p = 112 Then
        SniCgaOgir = (t = "iya" Or t = "ion")
    ElseIf p = 110 Then
        SniCgaOgir = (t = "iya" Or t = "ion" Or Left$(t, 2) = "ia")
    End If
End Function

' operatsiya, natsional, sotsial, sotsiolog, koeffitsient ... -> ts
' (ts dan keyin i va unli kelsa; ketsa, aytsin kabi o'zbekcha so'zlar o'zgarmaydi)
Private Function TsQoida(ByRef s As String, ByVal j As Long) As Boolean
    Dim t As String, b2 As String, b3 As String
    t = LCase$(Mid$(s, j, 3))
    If Len(t) < 2 Then Exit Function
    If Left$(t, 1) <> "i" Then Exit Function
    b2 = Mid$(t, 2, 1)
    If b2 = "a" Or b2 = "e" Or b2 = "o" Or b2 = "u" Then
        TsQoida = True
    ElseIf b2 = "y" And Len(t) = 3 Then
        b3 = Mid$(t, 3, 1)
        TsQoida = (b3 = "a" Or b3 = "e" Or b3 = "o" Or b3 = "u")
    End If
End Function

' Avtomobil raqamlari (40 757 DCA, 01 A 123 BC), raqamga yopishgan kodlar (A4, MP3)
' va xalqaro qisqartmalar (USD, PDF ...) lotincha qoladi
Private Function KodMi(ByRef s As String, ByVal i As Long, ByVal j As Long, ByVal n As Long, ByVal w As String) As Boolean
    Dim L As Long, p As Long, k As Long
    L = Len(w)
    If L = 0 Then Exit Function
    For k = 1 To L
        If Not IsLatU(KodAt(w, k)) Then Exit Function
    Next k
    If RaqamMi(KodAt(s, i - 1)) Or RaqamMi(KodAt(s, j)) Then
        KodMi = True
        Exit Function
    End If
    Select Case w
        Case "USD", "EUR", "RUB", "UZS", "GBP", "KZT", "PDF", "ID", "IT", "GPS", "SMS", _
             "USB", "PIN", "QR", "IP", "IBAN", "IMEI", "VIP", "HTML", "SIM"
            KodMi = True
            Exit Function
    End Select
    If AlnumMi(KodAt(s, j)) Then Exit Function
    p = i - 1
    If L = 3 Then
        If OrqagaRaqam(s, p, 3) Then
            If OrqagaRaqam(s, p, 2) Then KodMi = Not AlnumMi(KodAt(s, p))
        End If
    ElseIf L = 2 Then
        If OrqagaRaqam(s, p, 3) Then
            If OrqagaHarf(s, p, 1) Then
                If OrqagaRaqam(s, p, 2) Then KodMi = Not AlnumMi(KodAt(s, p))
            End If
        End If
    ElseIf L = 1 Then
        If OrqagaRaqam(s, p, 2) Then
            If Not AlnumMi(KodAt(s, p)) Then
                p = j
                If OldingaRaqam(s, p, 3) Then
                    If OldingaHarf(s, p, 2) Then KodMi = Not AlnumMi(KodAt(s, p))
                End If
            End If
        End If
    End If
End Function

' p dan chapga: bitta bo'shliqni o'tkazib, k ta raqam o'qiydi
Private Function OrqagaRaqam(ByRef s As String, ByRef p As Long, ByVal k As Long) As Boolean
    Dim m As Long
    If BoshliqMi(KodAt(s, p)) Then p = p - 1
    For m = 1 To k
        If Not RaqamMi(KodAt(s, p)) Then Exit Function
        p = p - 1
    Next m
    OrqagaRaqam = True
End Function

Private Function OrqagaHarf(ByRef s As String, ByRef p As Long, ByVal k As Long) As Boolean
    Dim m As Long
    If BoshliqMi(KodAt(s, p)) Then p = p - 1
    For m = 1 To k
        If Not IsLatU(KodAt(s, p)) Then Exit Function
        p = p - 1
    Next m
    OrqagaHarf = True
End Function

' p dan o'ngga: bitta bo'shliqni o'tkazib, k ta raqam / harf o'qiydi
Private Function OldingaRaqam(ByRef s As String, ByRef p As Long, ByVal k As Long) As Boolean
    Dim m As Long
    If BoshliqMi(KodAt(s, p)) Then p = p + 1
    For m = 1 To k
        If Not RaqamMi(KodAt(s, p)) Then Exit Function
        p = p + 1
    Next m
    OldingaRaqam = True
End Function

Private Function OldingaHarf(ByRef s As String, ByRef p As Long, ByVal k As Long) As Boolean
    Dim m As Long
    If BoshliqMi(KodAt(s, p)) Then p = p + 1
    For m = 1 To k
        If Not IsLatU(KodAt(s, p)) Then Exit Function
        p = p + 1
    Next m
    OldingaHarf = True
End Function

Private Function RaqamMi(ByVal c As Long) As Boolean
    RaqamMi = (c >= 48 And c <= 57)
End Function

Private Function BoshliqMi(ByVal c As Long) As Boolean
    BoshliqMi = (c = 32 Or c = 160)
End Function

Private Function AlnumMi(ByVal c As Long) As Boolean
    AlnumMi = RaqamMi(c) Or HarfMi(c)
End Function

' So'z boshidagi maxsus holatlar: internet manzil, rim raqami, istisno so'zlar
Private Function SozBoshi(ByRef s As String, ByVal i As Long, ByVal n As Long, ByRef lastK As Integer) As Long
    Dim j As Long, c As Long, k As Long, best As Long, bestLen As Long
    Dim tok As String, w As String, wl As String, key As String
    Dim dl As Long, di As Long, t As String

    j = i
    Do While j <= n
        c = KodAt(s, j)
        If c = 32 Or c = 9 Or c = 10 Or c = 13 Or c = 160 Then Exit Do
        j = j + 1
    Loop
    tok = Mid$(s, i, j - i)
    If InStr(tok, "@") > 0 Or InStr(tok, "://") > 0 Or LCase$(Left$(tok, 4)) = "www." Then
        BufAdd tok
        lastK = 0
        SozBoshi = j - i
        Exit Function
    End If

    j = i
    Do While j <= n
        c = KodAt(s, j)
        If IsLat(c) Then
            j = j + 1
        ElseIf IsApos(c) And j > i And j < n Then
            If IsLat(KodAt(s, j + 1)) Then j = j + 1 Else Exit Do
        Else
            Exit Do
        End If
    Loop
    ' so'z oxiridagi o' / g' ham so'zning harfi (yog', bog', bo') - so'zga qo'shamiz
    If j <= n And j > i Then
        If IsApos(KodAt(s, j)) Then
            c = LowLat(KodAt(s, j - 1))
            If c = 111 Or c = 103 Then j = j + 1
        End If
    End If
    w = Mid$(s, i, j - i)

    If KodMi(s, i, j, n, w) Then
        BufAdd w
        lastK = 2
        SozBoshi = Len(w)
        Exit Function
    End If

    If RimRaqami(w) Then
        BufAdd w
        lastK = 2
        SozBoshi = Len(w)
        Exit Function
    End If

    wl = NormalSoz(w)

    ' O'zbek alifbosida yo'q harflar (w, ch'siz c) bo'lsa - chet so'z, o'zgarmaydi
    If ChetSoz(wl) Then
        BufAdd w
        lastK = 2
        SozBoshi = Len(w)
        Exit Function
    End If

    For k = 1 To mExCount
        key = mExKey(k)
        If Len(key) > bestLen Then
            If mExWhole(k) Then
                If wl = key Then best = k: bestLen = Len(key)
            ElseIf Left$(wl, Len(key)) = key Then
                best = k: bestLen = Len(key)
            End If
        End If
    Next k
    ' Katta lug'at: so'z oxirida tutuq belgisi bo'lgan so'zlar (foji', voqe')
    If IsApos(KodAt(s, j)) Then
        If LugatQidir(wl & "'", di, True) > 0 Then
            BufAdd HarfShakli(mDictCyr(di), w)
            lastK = 2
            SozBoshi = Len(w) + 1
            Exit Function
        End If
    End If

    ' Katta lug'at: butun so'z yoki so'z o'zagi (eng uzun mos kelgani).
    ' Dasturning o'z istisnolar ro'yxati ustun turadi.
    If best = 0 Then dl = LugatQidir(wl, di, False)
    If dl > 0 Then
        t = mDictCyr(di)
        If dl < Len(wl) And Right$(t, 1) = ChrW(1100) Then t = Left$(t, Len(t) - 1)
        BufAdd HarfShakli(t, Left$(w, dl))
        If InStr("aeiou", Mid$(wl, dl, 1)) > 0 Then lastK = 1 Else lastK = 2
        SozBoshi = dl
        Exit Function
    End If

    If best > 0 Then
        BufAdd HarfShakli(BelgiliKirill(mExVal(best)), Left$(w, bestLen))
        If InStr("aeiou", Right$(mExKey(best), 1)) > 0 Then lastK = 1 Else lastK = 2
        SozBoshi = bestLen
    End If
End Function

' Lug'atni yashirin varaqdan bir marta o'qib olish
Private Sub LugatYukla()
    Dim ws As Object, v As Variant, n As Long, r As Long, k As String
    If mDictYuklandi Then
        If Not mDict Is Nothing Then Exit Sub
    End If
    mDictYuklandi = True
    Set mDict = New Collection
    Set mDictPre = New Collection
    On Error GoTo Tamom
    Set ws = ThisWorkbook.Worksheets(1)
    ' D1 katakda so'zlar soni, 2-qatordan boshlab: A = lotin, B = kirill, C = 1 (faqat butun so'z)
    n = CLng(Val(CStr(ws.Cells(1, 4).Value)))
    If n < 1 Then Exit Sub
    v = ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 3)).Value
    ReDim mDictCyr(1 To n)
    ReDim mDictWhole(1 To n)
    For r = 1 To n
        k = CStr(v(r, 1))
        If Len(k) >= 3 Then
            mDictCyr(r) = CStr(v(r, 2))
            mDictWhole(r) = (CStr(v(r, 3)) = "1")
            On Error Resume Next
            mDict.Add r, k
            mDictPre.Add 1, Left$(k, 3)
            Err.Clear
            On Error GoTo Tamom
        End If
    Next r
Tamom:
End Sub

' wl so'zining lug'atdagi eng uzun mos boshlanmasini topadi (uzunligini qaytaradi)
Private Function LugatQidir(ByVal wl As String, ByRef idx As Long, ByVal faqatButun As Boolean) As Long
    Dim L As Long, r As Long, bor As Boolean, eng As Long
    If Len(wl) < 3 Then Exit Function
    LugatYukla
    If mDict Is Nothing Then Exit Function
    If mDict.Count = 0 Then Exit Function
    On Error Resume Next
    r = 0
    r = mDictPre.Item(Left$(wl, 3))
    bor = (Err.Number = 0)
    Err.Clear
    If Not bor Then Exit Function
    ' butun so'z har doim tekshiriladi; o'zak sifatida esa kamida 4 harfli yozuv ishlatiladi
    If faqatButun Or Len(wl) <= 4 Then eng = Len(wl) Else eng = 4
    For L = Len(wl) To eng Step -1
        r = 0
        r = mDict.Item(Left$(wl, L))
        If Err.Number <> 0 Then
            Err.Clear
        ElseIf r > 0 Then
            If L = Len(wl) Or Not mDictWhole(r) Then
                idx = r
                LugatQidir = L
                Exit Function
            End If
        End If
    Next L
End Function

Private Function ChetSoz(ByVal wl As String) As Boolean
    Dim p As Long
    If InStr(wl, "w") > 0 Then
        ChetSoz = True
        Exit Function
    End If
    p = InStr(wl, "c")
    Do While p > 0
        If Mid$(wl, p + 1, 1) <> "h" Then
            ChetSoz = True
            Exit Function
        End If
        p = InStr(p + 1, wl, "c")
    Loop
End Function

Private Function RimRaqami(ByVal w As String) As Boolean
    Dim i As Long, v As Long, jami As Long, oldingi As Long
    If Len(w) = 0 Or Len(w) > 12 Then Exit Function
    For i = Len(w) To 1 Step -1
        Select Case Mid$(w, i, 1)
            Case "I": v = 1
            Case "V": v = 5
            Case "X": v = 10
            Case "L": v = 50
            Case "C": v = 100
            Case "D": v = 500
            Case "M": v = 1000
            Case Else: Exit Function
        End Select
        If v < oldingi Then
            jami = jami - v
        Else
            jami = jami + v
            oldingi = v
        End If
    Next i
    If jami < 1 Or jami > 3999 Then Exit Function
    RimRaqami = (RimYoz(jami) = w)
End Function

Private Function RimYoz(ByVal x As Long) As String
    Dim qiymat As Variant, belgi As Variant, k As Long, r As String
    qiymat = Array(1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1)
    belgi = Array("M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I")
    For k = 0 To 12
        Do While x >= qiymat(k)
            r = r & belgi(k)
            x = x - qiymat(k)
        Loop
    Next k
    RimYoz = r
End Function

Private Function NormalSoz(ByVal w As String) As String
    Dim i As Long, r As String
    r = LCase$(w)
    For i = 1 To Len(r)
        If IsApos(KodAt(r, i)) Then Mid$(r, i, 1) = "'"
    Next i
    NormalSoz = r
End Function

' Istisno so'zning kirillchasini manba so'zning harf shakliga moslash
Private Function HarfShakli(ByVal t As String, ByVal src As String) As String
    Dim i As Long, c As Long, nU As Long, nL As Long
    For i = 1 To Len(src)
        c = KodAt(src, i)
        If c >= 65 And c <= 90 Then
            nU = nU + 1
        ElseIf c >= 97 And c <= 122 Then
            nL = nL + 1
        End If
    Next i
    If nU >= 2 And nL = 0 Then
        HarfShakli = KattaKirill(t)
    ElseIf IsLatU(KodAt(src, 1)) Then
        HarfShakli = KattaKirill(Left$(t, 1)) & Mid$(t, 2)
    Else
        HarfShakli = t
    End If
End Function

Private Function KattaKirill(ByVal t As String) As String
    Dim i As Long, r As String
    For i = 1 To Len(t)
        r = r & ChrW(CyrUp(KodAt(t, i)))
    Next i
    KattaKirill = r
End Function

' Istisno ro'yxatidagi yozuvni kirillga o'girish.
' Maxsus belgilar: C = ts, ~ = yumshatish belgisi, ^ = tutuq belgisi,
' @ = e (so'z boshidagi e), # = e (ye). Qolgan harflar odatiy qoidada.
Private Function BelgiliKirill(ByVal m As String) As String
    Dim i As Long, n As Long, ch As String, ikki As String, r As String
    n = Len(m)
    i = 1
    Do While i <= n
        ch = Mid$(m, i, 1)
        ikki = Mid$(m, i, 2)
        Select Case ch
            Case "C": r = r & ChrW(1094): i = i + 1
            Case "~": r = r & ChrW(1100): i = i + 1
            Case "^": r = r & ChrW(1098): i = i + 1
            Case "@": r = r & ChrW(1101): i = i + 1
            Case "#": r = r & ChrW(1077): i = i + 1
            Case Else
                Select Case ikki
                    Case "sh": r = r & ChrW(1096): i = i + 2
                    Case "ch": r = r & ChrW(1095): i = i + 2
                    Case "yo": r = r & ChrW(1105): i = i + 2
                    Case "yu": r = r & ChrW(1102): i = i + 2
                    Case "ya": r = r & ChrW(1103): i = i + 2
                    Case "ye": r = r & ChrW(1077): i = i + 2
                    Case "o'": r = r & ChrW(1118): i = i + 2
                    Case "g'": r = r & ChrW(1171): i = i + 2
                    Case Else
                        r = r & ChrW(OddiyKod(AscW(ch)))
                        i = i + 1
                End Select
        End Select
    Loop
    BelgiliKirill = r
End Function

Private Function OddiyKod(ByVal lc As Long) As Long
    Select Case lc
        Case 97: OddiyKod = 1072    ' a
        Case 98: OddiyKod = 1073    ' b
        Case 100: OddiyKod = 1076   ' d
        Case 101: OddiyKod = 1077   ' e
        Case 102: OddiyKod = 1092   ' f
        Case 103: OddiyKod = 1075   ' g
        Case 104: OddiyKod = 1203   ' h
        Case 105: OddiyKod = 1080   ' i
        Case 106: OddiyKod = 1078   ' j
        Case 107: OddiyKod = 1082   ' k
        Case 108: OddiyKod = 1083   ' l
        Case 109: OddiyKod = 1084   ' m
        Case 110: OddiyKod = 1085   ' n
        Case 111: OddiyKod = 1086   ' o
        Case 112: OddiyKod = 1087   ' p
        Case 113: OddiyKod = 1179   ' q
        Case 114: OddiyKod = 1088   ' r
        Case 115: OddiyKod = 1089   ' s
        Case 116: OddiyKod = 1090   ' t
        Case 117: OddiyKod = 1091   ' u
        Case 118: OddiyKod = 1074   ' v
        Case 120: OddiyKod = 1093   ' x
        Case 121: OddiyKod = 1081   ' y
        Case 122: OddiyKod = 1079   ' z
        Case Else: OddiyKod = lc
    End Select
End Function

Private Sub IstisnoYukla()
    If mExCount > 0 Then Exit Sub
    ReDim mExKey(1 To 100)
    ReDim mExVal(1 To 100)
    ReDim mExWhole(1 To 100)
    ' --- ts li so'zlar ---
    Ex "abzas", "abzaC"
    Ex "aksiz", "akCiz"
    Ex "appenditsit", "appendiCit"
    Ex "vaksina", "vakCina"
    Ex "glitserin", "gliCerin"
    Ex "dotse", "doCe"
    Ex "inersiya", "inerCiya"
    Ex "kvars", "kvarC"
    Ex "konsern", "konCern"
    Ex "konsert", "konCert"
    Ex "konsep", "konCep"
    Ex "konsentr", "konCentr"
    Ex "kotangens", "kotangenC"
    Ex "tangens", "tangenC"
    Ex "litse", "liCe"
    Ex "medits", "mediC"
    Ex "mototsikl", "motoCikl"
    Ex "ofitse", "ofiCe"
    Ex "porsiya", "porCiya"
    Ex "prinsip", "prinCip"
    Ex "pritse", "priCe"
    Ex "protse", "proCe"
    Ex "retse", "reCe"
    Ex "sellofan", "Cellofan"
    Ex "sement", "Cement"
    Ex "senzura", "Cenzura"
    Ex "sentner", "Centner"
    Ex "sentr", "Centr"
    Ex "sex", "Cex"
    Ex "sikl", "Cikl"
    Ex "silindr", "Cilindr"
    Ex "sirkul", "Cirkul"
    Ex "sisterna", "Cisterna"
    Ex "sitata", "Citata"
    Ex "sitrus", "Citrus"
    Ex "sivil", "Civil"
    Ex "shpris", "shpriC"
    Ex "spets", "speC"
    Ex "fransu", "franCu"
    Ex "fransiya", "franCiya"
    Ex "ensiklopedi", "@nCiklopedi"
    Ex "sirk", "Cirk", True
    Ex "sink", "Cink", True
    Ex "ssen", "sCen"
    Ex "aksent", "akCent"
    Ex "aksept", "akCept"
    Ex "shveytsar", "shveyCar"
    Ex "shveysar", "shveyCar"
    Ex "kalsiy", "kal~Ciy"
    Ex "selsiy", "Cel~siy"
    ' --- s qoladigan so'zlar ---
    Ex "pansion", "pansion"
    Ex "pensi", "pensi"
    Ex "ekspansi", "@kspansi"
    Ex "refleks", "refleks"
    Ex "epilepsi", "@pilepsi"
    Ex "biopsi", "biopsi"
    ' --- yumshatish va tutuq belgili so'zlar ---
    Ex "kompyuter", "komp~yuter"
    Ex "intervyu", "interv~yu"
    Ex "pyesa", "p~#sa"
    Ex "premyer", "prem~#r"
    Ex "kuryer", "kur~#r"
    Ex "batalyon", "batal~on"
    Ex "pavilyon", "pavil~on"
    Ex "medalyon", "medal~on"
    Ex "mo'jiza", "mo'^jiza"
    Ex "mo'tabar", "mo'^tabar"
    Ex "mo'tadil", "mo'^tadil"
    Ex "nyu", "n~yu", True
    Ex "ges", "g@s", True
    Ex "mer", "m@r", True
    Ex "duel", "du@l~", True
    Ex "pochtalyon", "pochtal~on"
    Ex "kompanyon", "kompan~on"
    Ex "bulyon", "bul~on"
    Ex "obyom", "ob^yom"
    Ex "konyunktur", "kon^yunktur"
    Ex "adyutant", "ad^yutant"
    Ex "vyetnam", "v~#tnam"
    Ex "atelye", "atel~#"
    Ex "portyer", "port~#r"
    ' so'z ichida yumshatish belgisi bor so'zlar (qo'shimcha bilan ham)
    Ex "film", "fil~m"
    Ex "multfilm", "mul~tfil~m"
    Ex "mult", "mul~t"
    Ex "filtr", "fil~tr"
    Ex "alternativ", "al~ternativ"
    Ex "albom", "al~bom"
    Ex "alpinist", "al~pinist"
    Ex "asfalt", "asfal~t"
    Ex "bulvar", "bul~var"
    Ex "konsult", "konsul~t"
    Ex "kult", "kul~t"
    Ex "saldo", "sal~do"
    Ex "rels", "rel~s"
    Ex "tyulpan", "tyul~pan"
    Ex "feldsher", "fel~dsher"
    Ex "folga", "fol~ga"
    Ex "delfin", "del~fin"
    Ex "delta", "del~ta"
    Ex "bilyard", "bil~yard"
    Ex "emulsiya", "@mul~siya"
    ' oxiri yumshatish belgisi bilan tugaydigan so'zlar (faqat butun so'z)
    Ex "avtomobil", "avtomobil~", True
    Ex "gastrol", "gastrol~", True
    Ex "gospital", "gospital~", True
    Ex "dizel", "dizel~", True
    Ex "kalendar", "kalendar~", True
    Ex "mebel", "mebel~", True
    Ex "medal", "medal~", True
    Ex "model", "model~", True
    Ex "neft", "neft~", True
    Ex "otel", "otel~", True
    Ex "parol", "parol~", True
    Ex "portfel", "portfel~", True
    Ex "pult", "pul~t", True
    Ex "puls", "pul~s", True
    Ex "rol", "rol~", True
    Ex "rubl", "rubl~", True
    Ex "sekretar", "sekretar~", True
    Ex "inventar", "inventar~", True
    Ex "spektakl", "spektakl~", True
    Ex "stil", "stil~", True
    Ex "shampun", "shampun~", True
    Ex "golf", "gol~f", True
    Ex "byulleten", "byulleten~", True
    Ex "veksel", "veksel~", True
    Ex "kafel", "kafel~", True
    Ex "kontrol", "kontrol~", True
    Ex "lager", "lager~", True
    Ex "volt", "vol~t", True
    Ex "belarus", "belarus~", True
    ' --- oylar ---
    Ex "yanvar", "yanvar~", True
    Ex "fevral", "fevral~", True
    Ex "aprel", "aprel~", True
    Ex "iyun", "iyun~", True
    Ex "iyul", "iyul~", True
    Ex "sentyabr", "sentyabr~", True
    Ex "oktyabr", "oktyabr~", True
    Ex "noyabr", "noyabr~", True
    Ex "dekabr", "dekabr~", True
End Sub

Private Sub Ex(ByVal k As String, ByVal v As String, Optional ByVal butun As Boolean = False)
    mExCount = mExCount + 1
    If mExCount > UBound(mExKey) Then
        ReDim Preserve mExKey(1 To mExCount + 50)
        ReDim Preserve mExVal(1 To mExCount + 50)
        ReDim Preserve mExWhole(1 To mExCount + 50)
    End If
    mExKey(mExCount) = k
    mExVal(mExCount) = v
    mExWhole(mExCount) = butun
End Sub

'==========================================================================
'  YORDAMCHI FUNKSIYALAR
'==========================================================================
Private Sub ApostrofYukla()
    Dim v As String
    If mAposYuklandi Then Exit Sub
    v = "0"
    On Error Resume Next
    v = GetSetting(ILOVA, "Sozlama", "OddiyApostrof", "0")
    On Error GoTo 0
    mOddiy = (v = "1")
    BelgilarniSozla
    mAposYuklandi = True
End Sub

Private Sub BelgilarniSozla()
    If mOddiy Then
        mOg = "'"
        mTq = "'"
    Else
        mOg = ChrW(8216)
        mTq = ChrW(8217)
    End If
End Sub

Private Sub BufStart(ByVal n As Long)
    mBuf = Space$(n * 2 + 16)
    mLen = 0
End Sub

Private Sub BufAdd(ByVal t As String)
    Dim k As Long
    k = Len(t)
    If k = 0 Then Exit Sub
    If mLen + k > Len(mBuf) Then mBuf = mBuf & Space$(Len(mBuf) + k)
    Mid$(mBuf, mLen + 1, k) = t
    mLen = mLen + k
End Sub

Private Function BufText() As String
    BufText = Left$(mBuf, mLen)
End Function

Private Sub Chiq(ByVal code As Long, ByVal up As Boolean)
    If up Then
        BufAdd ChrW(CyrUp(code))
    Else
        BufAdd ChrW(code)
    End If
End Sub

Private Function KodAt(ByRef s As String, ByVal i As Long) As Long
    Dim c As Long
    If i < 1 Or i > Len(s) Then
        KodAt = 0
        Exit Function
    End If
    c = AscW(Mid$(s, i, 1))
    If c < 0 Then c = c + 65536
    KodAt = c
End Function

Private Function IsLat(ByVal c As Long) As Boolean
    IsLat = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

Private Function IsLatU(ByVal c As Long) As Boolean
    IsLatU = (c >= 65 And c <= 90)
End Function

Private Function LowLat(ByVal c As Long) As Long
    If c >= 65 And c <= 90 Then LowLat = c + 32 Else LowLat = c
End Function

Private Function IsCyr(ByVal c As Long) As Boolean
    IsCyr = (c >= 1024 And c <= 1119) Or (c >= 1162 And c <= 1215)
End Function

Private Function HarfMi(ByVal c As Long) As Boolean
    HarfMi = IsLat(c) Or IsCyr(c)
End Function

Private Function KattaMi(ByVal c As Long) As Boolean
    If c >= 65 And c <= 90 Then
        KattaMi = True
    ElseIf c >= 1024 And c <= 1071 Then
        KattaMi = True
    ElseIf c >= 1162 And c <= 1215 Then
        KattaMi = ((c Mod 2) = 0)
    End If
End Function

Private Function CyrLow(ByVal c As Long) As Long
    If c >= 1040 And c <= 1071 Then
        CyrLow = c + 32
    ElseIf c >= 1024 And c <= 1039 Then
        CyrLow = c + 80
    ElseIf c >= 1162 And c <= 1215 And (c Mod 2) = 0 Then
        CyrLow = c + 1
    Else
        CyrLow = c
    End If
End Function

Private Function CyrUp(ByVal c As Long) As Long
    If c >= 1072 And c <= 1103 Then
        CyrUp = c - 32
    ElseIf c >= 1104 And c <= 1119 Then
        CyrUp = c - 80
    ElseIf c >= 1163 And c <= 1215 And (c Mod 2) = 1 Then
        CyrUp = c - 1
    Else
        CyrUp = c
    End If
End Function

Private Function KirUnli(ByVal c As Long) As Boolean
    Select Case c
        Case 1072, 1077, 1105, 1080, 1086, 1091, 1118, 1101, 1102, 1103, 1099
            KirUnli = True
    End Select
End Function

Private Function KirUndosh(ByVal c As Long) As Boolean
    KirUndosh = IsCyr(c) And Not KirUnli(c) And c <> 1098 And c <> 1100
End Function

Private Function IsApos(ByVal c As Long) As Boolean
    Select Case c
        Case 39, 96, 180, 697, 699, 700, 702, 703, 8216, 8217, 8219, 8242
            IsApos = True
    End Select
End Function
