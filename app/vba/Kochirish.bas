Attribute VB_Name = "Kochirish"
Option Explicit

' Filiallararo dori ko'chirish ro'yxati ("Ko'chirish" va "Skladdan" varaqlari).
' Python etaloni: bot papkasidagi tools\kochirish_ref.py (qadamlar [0]..[9] bir xil raqamlangan).
'
' Turi (K): SHOSHILINCH / ASSORTIMENT / ODDIY. Manfiy sotuv (qaytarilgan tovar) = 0.
' q = Int(yaroqli qoldiq); ceil(x) = -Int(-x). Qonunlar: beruvchi hech qachon 0 da qolmaydi; oluvchiga faqat kerakli miqdor.
'   ehtiyoj(ol)   = max(0, ceil(s(ol) * KOCH_COVER) - q(ol))
'   qoldiradi(b)  = max(MIN_QOLDIQ, ceil(s(b) * KOCH_COVER));  bera_oladi(b) = q(b) - qoldiradi(b)
' SHOSHILINCH: q(ol) = 0, s(ol) > 0                            -> min(bera_oladi, ehtiyoj)
' ASSORTIMENT: q(ol) = 0, s(ol) = 0, boshqa filialda sotuv bor -> min(ASSORT_SONI, q(b) - MIN_QOLDIQ)
' ODDIY : s(b) = 0 va "30+ kun" >= TURGAN_KUN, s(ol) > 0, q(ol) > 0 -> min(bera_oladi, ehtiyoj); summa < MIN_SUMMA_ODDIY - yo'q
' Himoya: beruvchida qoladigan < MIN_QOLDIQ yoki berilgan > ehtiyoj bo'lsa qator yozilmaydi -> yashirin "KochLog" varag'i.
' SKLADda qoldiq bo'lsa: filiallardan ko'chirilmaydi, "Skladdan" varag'iga ehtiyojcha (sklad Int(K) gacha, MIN_QOLDIQ yo'q).

' [0] Sozlamalar
Private Const KOCH_COVER As Double = 1         ' oluvchiga necha davrlik sotuv (1 = davr sotuvicha)
Private Const MIN_QOLDIQ As Long = 1           ' beruvchida kamida shuncha qoladi (1 dan kam bo'lsa 1)
Private Const TURGAN_KUN As Double = 30        ' oddiy ko'chirish: beruvchida kamida shuncha kun turgan ("30+ kun")
Private Const MIN_SUMMA_ODDIY As Double = 50000
Private Const MIN_SUMMA_SHOSH As Double = 0    ' SHOSHILINCH va ASSORTIMENT uchun minimal summa
Private Const ASSORT_SONI As Long = 1          ' assortiment bo'shlig'i (Z2) uchun yuboriladigan qadoq
Private Const SROK_MIN_KUN As Double = 0       ' partiyaning qolgan kuni bundan kam bo'lsa ko'chirilmaydi

Private Const SH_KOCH As String = "Ko'chirish"
Private Const SH_SKLAD As String = "Skladdan"
Private Const SH_LOG As String = "KochLog"     ' yashirin: himoya tekshiruvidan o'tmagan qatorlar (har ishga tushishda tozalanadi)
Private Const NF As Long = 3                   ' filiallar soni (E6, G6, I6)
Private Const T_Z1 As Long = 1
Private Const T_Z2 As Long = 2
Private Const T_OD As Long = 3
Private Const SABAB_Z1 As String = "Tugagan, sotuv bor"
Private Const SABAB_Z2 As String = "Filialda yo'q, boshqa filialda sotilyapti"
Private Const SABAB_SKLAD_OD As String = "Skladda bor, oluvchida yetmaydi"

Private mFound As Object, mValid As Object, mNear As Object, mSPrice As Object, mTurgan As Object, mNarx As Object
Private mLog As Worksheet, mLogN As Long

Public Sub KochirishRoyxati()
    Dim natija As String
    On Error GoTo Fail
    Application.ScreenUpdating = False
    natija = BuildKochirish()
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(SH_KOCH).Activate
    MsgBox natija
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Xato: " & Err.Description
End Sub

Public Function BuildKochirish() As String
    Dim ws As Worksheet, d As Variant, lastRow As Long, r As Long, i As Long, j As Long, g As Long, x As Long
    Dim fil(1 To NF) As String, nm As String, k As String, fk As String
    Dim q(1 To NF) As Double, raw(1 To NF) As Double, s(1 To NF) As Double, v(1 To NF) As Double
    Dim cap(1 To NF) As Double, capZ2(1 To NF) As Double, tk(1 To NF) As Double, given(1 To NF) As Double
    Dim rTp(1 To NF) As Long, rNeed(1 To NF) As Double, rOrd(1 To NF) As Long, nr As Long
    Dim gOrd(1 To NF) As Long, gAv(1 To NF) As Double, ng As Long, t As Long
    Dim anySale As Boolean, keep As Double, price As Double, skq As Double, skcap As Double, mq As Long
    Dim remN As Double, n As Double, summa As Double, eh As Double, tp As Long, avail As Double, leftQ As Double
    Dim kOut() As Variant, sOut() As Variant, nk As Long, ns As Long, jami As Double
    Dim nc(1 To 3) As Long
    Dim donNames As String, donQty As Double

    Set ws = ThisWorkbook.Worksheets(1)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 8 Then
        BuildKochirish = "Asosiy varaqda ma'lumot yo'q."
        Exit Function
    End If
    d = ws.Range("A8:L" & lastRow).Value
    fil(1) = Norm(ws.Range("E6").Value)
    fil(2) = Norm(ws.Range("G6").Value)
    fil(3) = Norm(ws.Range("I6").Value)
    mq = MIN_QOLDIQ
    If mq < 1 Then mq = 1
    PrepareLog

    LoadSrok                                                          ' [1]
    LoadTurgan                                                        ' [2]
    LoadNarx                                                          ' [3]

    ReDim kOut(1 To 2 * NF * UBound(d), 1 To 15)   ' 15-ustun: saralash uchun vaqtinchalik toifa tartibi (keyin o'chiriladi)
    ReDim sOut(1 To NF * UBound(d), 1 To 14)

    For r = 1 To UBound(d)
        nm = Norm(d(r, 1))
        If Len(nm) = 0 Then GoTo NextRow
        k = nm & "|" & Norm(d(r, 2))

        ' [4] har filial: qoldiq q = Int(yaroqli), sotuv, bera oladi (SHOSHILINCH/ODDIY va ASSORTIMENT uchun alohida)
        anySale = False
        For i = 1 To NF
            raw(i) = Num(d(r, 3 + 2 * i))
            s(i) = Num(d(r, 4 + 2 * i))
            If s(i) < 0 Then s(i) = 0               ' manfiy sotuv (qaytarilgan tovar) = 0
            If s(i) > 0 Then anySale = True
            fk = fil(i) & "|" & k
            If mFound.Exists(fk) Then
                v(i) = 0
                If mValid.Exists(fk) Then v(i) = mValid(fk)
                If v(i) > raw(i) Then v(i) = raw(i)
            Else
                v(i) = raw(i)
            End If
            ' Round(.., 6): suzuvchi nuqta xatosini yo'qotadi; kasr qadoq berilmaydi
            q(i) = Int(Round(v(i), 6))
            If q(i) < 0 Then q(i) = 0
            ' beruvchi hech qachon 0 da qolmaydi: o'zida max(MIN_QOLDIQ, o'z davr sotuvi) qoladi
            keep = Ceil6(s(i) * KOCH_COVER)
            If keep < mq Then keep = mq
            cap(i) = q(i) - keep
            If cap(i) < 0 Then cap(i) = 0
            ' ISTISNO (ASSORTIMENT): beruvchining o'z sotuvi hisobga olinmaydi, faqat MIN_QOLDIQ o'zida qoladi.
            ' Masalan q=6, s=6: o'ziga 6 kerak bo'lsa ham assortiment uchun 1 qadoq beradi (biznes talabi).
            capZ2(i) = q(i) - mq
            If capZ2(i) < 0 Then capZ2(i) = 0
            tk(i) = 0
            If mTurgan.Exists(fk) Then tk(i) = mTurgan(fk)
            given(i) = 0
        Next i
        price = 0
        If mNarx.Exists(k) Then price = mNarx(k)
        If price <= 0 Then
            For i = 1 To NF
                If mSPrice.Exists(fil(i) & "|" & k) Then
                    If mSPrice(fil(i) & "|" & k) > price Then price = mSPrice(fil(i) & "|" & k)
                End If
            Next i
        End If

        ' [5] oluvchilar: turi va ehtiyoj (faqat kerakli miqdor: davr sotuvi - qoldig'i)
        nr = 0
        For j = 1 To NF
            tp = 0
            eh = Ceil6(s(j) * KOCH_COVER) - q(j)
            If eh < 0 Then eh = 0
            If q(j) = 0 And s(j) > 0 Then
                tp = T_Z1: n = eh
            ElseIf q(j) = 0 And s(j) = 0 And anySale Then
                tp = T_Z2: n = ASSORT_SONI
            ElseIf q(j) > 0 And s(j) > 0 And eh >= 1 Then
                tp = T_OD: n = eh
            End If
            If tp > 0 Then
                nr = nr + 1
                rOrd(nr) = j: rTp(j) = tp: rNeed(j) = n
            End If
        Next j
        ' [6] tartib: SHOSHILINCH -> ASSORTIMENT -> ODDIY, ichida ehtiyoji kattasi, teng bo'lsa filial nomi (alifbo)
        For i = 1 To nr - 1
            For x = i + 1 To nr
                If RecvBefore(rOrd(x), rOrd(i), rTp, rNeed, fil) Then
                    t = rOrd(i): rOrd(i) = rOrd(x): rOrd(x) = t
                End If
            Next x
        Next i

        skq = Num(d(r, 11))
        If skq > 0 Then
            ' [7] SKLAD: filiallardan ko'chirilmaydi, sklad Int(K) gacha ehtiyojcha beradi (MIN_QOLDIQ va "30+ kun" yo'q)
            skcap = Int(skq)
            donNames = "": donQty = 0
            For i = 1 To NF
                If raw(i) > 0 And s(i) <= 0 Then
                    If Len(donNames) > 0 Then donNames = donNames & ", "
                    donNames = donNames & fil(i)
                    donQty = donQty + raw(i)
                End If
            Next i
            For x = 1 To nr
                If skcap < 1 Then Exit For
                j = rOrd(x): tp = rTp(j)
                n = Int(Min2(rNeed(j), skcap))
                summa = n * price
                If (tp = T_OD And summa < MIN_SUMMA_ODDIY) Or (tp <> T_OD And summa < MIN_SUMMA_SHOSH) Then GoTo NextSk
                ns = ns + 1
                sOut(ns, 1) = fil(j)
                sOut(ns, 2) = d(r, 1)
                sOut(ns, 3) = d(r, 2)
                sOut(ns, 4) = skq
                sOut(ns, 5) = s(j)
                sOut(ns, 6) = raw(j)
                sOut(ns, 7) = donNames
                sOut(ns, 8) = donQty
                sOut(ns, 9) = n
                sOut(ns, 10) = TuriNomi(tp)
                sOut(ns, 11) = SababNomi(tp, 0, True)
                sOut(ns, 12) = price
                sOut(ns, 13) = summa
                sOut(ns, 14) = tp
                skcap = skcap - n
NextSk:
            Next x
            GoTo NextRow
        End If

        ' [8] filiallararo: har oluvchiga beruvchilar - ko'p bera oladigani birinchi (teng: ko'proq turgani, filial nomi)
        For x = 1 To nr
            j = rOrd(x): tp = rTp(j)
            remN = rNeed(j)
            ng = 0
            For g = 1 To NF
                If g <> j Then
                    If tp = T_Z2 Then avail = capZ2(g) - given(g) Else avail = cap(g) - given(g)
                    If avail >= 1 Then
                        If tp <> T_OD Or (s(g) <= 0 And tk(g) >= TURGAN_KUN) Then
                            ng = ng + 1: gOrd(ng) = g: gAv(g) = avail
                        End If
                    End If
                End If
            Next g
            For i = 1 To ng - 1
                For g = i + 1 To ng
                    If GiverBefore(gOrd(g), gOrd(i), tk, gAv, fil) Then
                        t = gOrd(i): gOrd(i) = gOrd(g): gOrd(g) = t
                    End If
                Next g
            Next i
            For i = 1 To ng
                g = gOrd(i)
                n = Int(Min2(remN, gAv(g)))
                If n < 1 Then GoTo NextGiver
                summa = n * price
                If (tp = T_OD And summa < MIN_SUMMA_ODDIY) Or (tp <> T_OD And summa < MIN_SUMMA_SHOSH) Then GoTo NextGiver
                ' himoya: beruvchi 0 da qolmasin, oluvchiga ehtiyojdan ko'p emas (ASSORTIMENT: ASSORT_SONI dan ko'p emas)
                leftQ = q(g) - given(g) - n
                If leftQ < mq Or n > rNeed(j) Then
                    LogBad fil(g), fil(j), d(r, 1), d(r, 2), n, TuriNomi(tp), q(g), leftQ, rNeed(j)
                    GoTo NextGiver
                End If
                nk = nk + 1
                kOut(nk, 1) = fil(g)
                kOut(nk, 2) = fil(j)
                kOut(nk, 3) = d(r, 1)
                kOut(nk, 4) = d(r, 2)
                kOut(nk, 5) = n
                kOut(nk, 6) = s(j)
                kOut(nk, 7) = raw(j)
                kOut(nk, 8) = price
                kOut(nk, 9) = summa
                If mNear.Exists(fil(g) & "|" & k) Then kOut(nk, 10) = mNear(fil(g) & "|" & k)
                kOut(nk, 11) = TuriNomi(tp)
                kOut(nk, 12) = SababNomi(tp, tk(g), False)
                kOut(nk, 13) = s(g)
                kOut(nk, 14) = raw(g)
                kOut(nk, 15) = tp
                nc(tp) = nc(tp) + 1
                jami = jami + summa
                given(g) = given(g) + n
                remN = remN - n
                If remN < 1 Then Exit For
NextGiver:
            Next i
        Next x
NextRow:
    Next r

    WriteSheet SH_KOCH, Array("Qayerdan", "Qayerga", "Dori", "Ishlab chiqaruvchi", "Ko'chirish soni", _
        "Oluvchida sotuv", "Oluvchida qoldiq", "Kirim narxi", "Summa (kirim)", "Eng yaqin muddat", _
        "Turi", "Sabab", "Beruvchida sotuv", "Beruvchida qoldiq", "_tartib"), _
        kOut, nk, Array(20, 20, 40, 26, 11, 11, 11, 12, 14, 13, 13, 36, 11, 11, 6)
    With ThisWorkbook.Worksheets(SH_KOCH)
        .Columns("E").NumberFormat = "0"
        .Columns("H:I").NumberFormat = "#,##0"
        .Columns("J").NumberFormat = "dd.mm.yyyy"
        ' [9] tartib: Zudlik oldin ("Zudlik" > "Oddiy"), eng yaqin muddat (bo'shi oxirida), summa kattasi
        If nk > 1 Then .Range("A1").Resize(nk + 1, 15).Sort Key1:=.Range("O2"), Order1:=xlAscending, _
            Key2:=.Range("J2"), Order2:=xlAscending, Key3:=.Range("I2"), Order3:=xlDescending, Header:=xlYes
        If .AutoFilterMode Then .AutoFilterMode = False
        .Columns(15).Delete
        .Range("A1").Resize(nk + 1, 14).AutoFilter
    End With

    WriteSheet SH_SKLAD, Array("Qayerga", "Dori", "Ishlab chiqaruvchi", "Skladda qoldiq", "Sotuv", _
        "Filialda qoldiq", "Sotilmayotgan filial(lar)", "Ulardagi qoldiq", "Soni", "Turi", "Sabab", _
        "Kirim narxi", "Summa (kirim)", "_tartib"), _
        sOut, ns, Array(20, 40, 26, 12, 10, 12, 40, 12, 9, 13, 36, 12, 14, 6)
    With ThisWorkbook.Worksheets(SH_SKLAD)
        .Columns("H").NumberFormat = "0.##"
        .Columns("I").NumberFormat = "0"
        .Columns("L:M").NumberFormat = "#,##0"
        If ns > 1 Then .Range("A1").Resize(ns + 1, 14).Sort Key1:=.Range("N2"), Order1:=xlAscending, _
            Key2:=.Range("M2"), Order2:=xlDescending, Header:=xlYes
        If .AutoFilterMode Then .AutoFilterMode = False
        .Columns(14).Delete
        .Range("A1").Resize(ns + 1, 13).AutoFilter
    End With

    BuildKochirish = "SHOSHILINCH: " & nc(T_Z1) & " " & ChrW(183) & " ASSORTIMENT: " & nc(T_Z2) & " " & ChrW(183) & _
        " ODDIY: " & nc(T_OD) & " " & ChrW(183) & " jami " & Format(jami, "#,##0") & " so'm " & ChrW(183) & _
        " Skladdan: " & ns
End Function

' [6] oluvchi a oluvchi b dan oldin turadimi: toifa, ehtiyoj kattasi, filial nomi (alifbo), filial tartibi
Private Function RecvBefore(ByVal a As Long, ByVal b As Long, rTp() As Long, rNeed() As Double, fil() As String) As Boolean
    Dim ra As Long, rb As Long
    ra = rTp(a)                                  ' SHOSHILINCH (1) -> ASSORTIMENT (2) -> ODDIY (3)
    rb = rTp(b)
    If ra <> rb Then
        RecvBefore = (ra < rb)
    ElseIf rNeed(a) <> rNeed(b) Then
        RecvBefore = (rNeed(a) > rNeed(b))
    ElseIf fil(a) <> fil(b) Then
        RecvBefore = (StrComp(fil(a), fil(b), vbBinaryCompare) < 0)
    Else
        RecvBefore = (a < b)
    End If
End Function

' [8] beruvchi a beruvchi b dan oldin turadimi: ko'p bera oladigani, ko'proq turgani, filial nomi, filial tartibi
Private Function GiverBefore(ByVal a As Long, ByVal b As Long, tk() As Double, gAv() As Double, fil() As String) As Boolean
    If gAv(a) <> gAv(b) Then
        GiverBefore = (gAv(a) > gAv(b))
    ElseIf tk(a) <> tk(b) Then
        GiverBefore = (tk(a) > tk(b))
    ElseIf fil(a) <> fil(b) Then
        GiverBefore = (StrComp(fil(a), fil(b), vbBinaryCompare) < 0)
    Else
        GiverBefore = (a < b)
    End If
End Function

' ceil(x) = -Int(-x); Round(.., 6) suzuvchi nuqta xatosini yo'qotadi (0,1*3 = 0,30000000000000004)
Private Function Ceil6(ByVal x As Double) As Double
    Ceil6 = -Int(-Round(x, 6))
End Function

' Yashirin "KochLog" varag'i: har ishga tushishda tozalanadi (refresh_status.txt ga yozilmaydi)
Private Sub PrepareLog()
    On Error Resume Next
    Set mLog = Nothing
    Set mLog = ThisWorkbook.Worksheets(SH_LOG)
    On Error GoTo 0
    If mLog Is Nothing Then
        Set mLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        mLog.Name = SH_LOG
    End If
    mLog.Cells.Clear
    mLog.Range("A1:J1").Value = Array("Qayerdan", "Qayerga", "Dori", "Ishlab chiqaruvchi", "Soni", "Turi", _
        "Beruvchi qoldig'i", "Qolar edi", "Ehtiyoj", "Vaqt")
    mLog.Visible = xlSheetHidden
    mLogN = 1
End Sub

Private Sub LogBad(ByVal src As String, ByVal dst As String, ByVal nm As Variant, ByVal pr As Variant, ByVal n As Double, _
                   ByVal turi As String, ByVal qol As Double, ByVal leftQ As Double, ByVal need As Double)
    If mLog Is Nothing Then Exit Sub
    mLogN = mLogN + 1
    mLog.Range("A" & mLogN & ":J" & mLogN).Value = Array(src, dst, nm, pr, n, turi, qol, leftQ, need, Now)
End Sub

Private Function TuriNomi(ByVal tp As Long) As String
    Select Case tp
        Case T_Z1: TuriNomi = "SHOSHILINCH"
        Case T_Z2: TuriNomi = "ASSORTIMENT"
        Case Else: TuriNomi = "ODDIY"
    End Select
End Function

Private Function SababNomi(ByVal tp As Long, ByVal turgan As Double, ByVal sklad As Boolean) As String
    Select Case tp
        Case T_Z1: SababNomi = SABAB_Z1
        Case T_Z2: SababNomi = SABAB_Z2
        Case Else
            If sklad Then
                SababNomi = SABAB_SKLAD_OD
            Else
                SababNomi = Int(turgan) & " kun turibdi, oluvchida yetmaydi"
            End If
    End Select
End Function

Private Function Min2(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Min2 = a Else Min2 = b
End Function

' [1] "Srok": filial+dori bo'yicha yaroqli qoldiq, eng yaqin yaroqli muddat, kirim narxi
Private Sub LoadSrok()
    Dim w As Worksheet, a As Variant, lr As Long, r As Long, fk As String, ost As Double, kun As Double
    Set mFound = CreateObject("Scripting.Dictionary")
    Set mValid = CreateObject("Scripting.Dictionary")
    Set mNear = CreateObject("Scripting.Dictionary")
    Set mSPrice = CreateObject("Scripting.Dictionary")
    Set w = ThisWorkbook.Worksheets("Srok")
    lr = w.Cells(w.Rows.Count, 1).End(xlUp).Row
    If lr < 5 Then Exit Sub
    a = w.Range("A5:H" & lr).Value
    For r = 1 To UBound(a)
        If Len(Norm(a(r, 1))) > 0 And Len(Norm(a(r, 2))) > 0 Then
            ost = Num(a(r, 7))
            If ost > 0 Then
                fk = Norm(a(r, 1)) & "|" & Norm(a(r, 2)) & "|" & Norm(a(r, 3))
                mFound(fk) = True
                If VarType(a(r, 5)) = vbDate Then
                    kun = CLng(Int(a(r, 5))) - CLng(Date)
                Else
                    kun = Num(a(r, 6))
                End If
                If kun >= 0 And kun >= SROK_MIN_KUN Then     ' muddati o'tgan partiya doim chiqariladi
                    mValid(fk) = mValid(fk) + ost
                    If VarType(a(r, 5)) = vbDate Then
                        If Not mNear.Exists(fk) Then
                            mNear(fk) = a(r, 5)
                        ElseIf a(r, 5) < mNear(fk) Then
                            mNear(fk) = a(r, 5)
                        End If
                    End If
                End If
                If Not mSPrice.Exists(fk) And Num(a(r, 8)) > 0 Then mSPrice(fk) = Num(a(r, 8))
            End If
        End If
    Next r
End Sub

' [2] "30+ kun": filial+dori necha kun turibdi (eng kattasi)
Private Sub LoadTurgan()
    Dim w As Worksheet, a As Variant, lr As Long, r As Long, fk As String
    Set mTurgan = CreateObject("Scripting.Dictionary")
    Set w = ThisWorkbook.Worksheets("30+ kun")
    lr = w.Cells(w.Rows.Count, 1).End(xlUp).Row
    If lr < 5 Then Exit Sub
    a = w.Range("A5:F" & lr).Value
    For r = 1 To UBound(a)
        fk = Norm(a(r, 1)) & "|" & Norm(a(r, 2)) & "|" & Norm(a(r, 3))
        If Not mTurgan.Exists(fk) Then
            mTurgan(fk) = Num(a(r, 6))
        ElseIf Num(a(r, 6)) > mTurgan(fk) Then
            mTurgan(fk) = Num(a(r, 6))
        End If
    Next r
End Sub

' [3] "Narx": birinchi uchragan kirim narxi
Private Sub LoadNarx()
    Dim w As Worksheet, a As Variant, lr As Long, r As Long, k As String
    Set mNarx = CreateObject("Scripting.Dictionary")
    Set w = ThisWorkbook.Worksheets("Narx")
    lr = w.Cells(w.Rows.Count, 1).End(xlUp).Row
    If lr < 5 Then Exit Sub
    a = w.Range("A5:D" & lr).Value
    For r = 1 To UBound(a)
        If Len(Norm(a(r, 1))) > 0 Then
            k = Norm(a(r, 1)) & "|" & Norm(a(r, 2))
            If Not mNarx.Exists(k) Then mNarx(k) = Num(a(r, 4))
        End If
    Next r
End Sub

Private Sub WriteSheet(ByVal shName As String, ByVal headers As Variant, data As Variant, _
                       ByVal n As Long, ByVal widths As Variant)
    Dim w As Worksheet, c As Long
    On Error Resume Next
    Set w = ThisWorkbook.Worksheets(shName)
    On Error GoTo 0
    If w Is Nothing Then
        Set w = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        w.Name = shName
    End If
    If w.AutoFilterMode Then w.AutoFilterMode = False
    w.Cells.Clear

    With w.Range("A1").Resize(1, UBound(headers) + 1)
        .Value = headers
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(48, 84, 150)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    For c = 0 To UBound(widths)
        w.Columns(c + 1).ColumnWidth = widths(c)
    Next c
    If n > 0 Then
        w.Range("A2").Resize(n, UBound(headers) + 1).Value = TopRows(data, n, UBound(headers) + 1)
    End If
    w.Range("A1").Resize(n + 1, UBound(headers) + 1).AutoFilter
End Sub

Private Function TopRows(data As Variant, ByVal n As Long, ByVal cols As Long) As Variant
    Dim res() As Variant, r As Long, c As Long
    ReDim res(1 To n, 1 To cols)
    For r = 1 To n
        For c = 1 To cols
            res(r, c) = data(r, c)
        Next c
    Next r
    TopRows = res
End Function

' Filial/dori nomini solishtirish uchun: ortiqcha probellar birlashadi, katta harf
' ("ÄÎÐÈÕÎÍÀ  ÝÐÊÈÍ" = "ÄÎÐÈÕÎÍÀ ÝÐÊÈÍ")
Private Function Norm(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Norm = UCase$(Application.WorksheetFunction.Trim(CStr(v)))
    ' apostroflar bittasiga (U+02BC, U+2019, U+2018, `) -> '
    Norm = Replace(Replace(Replace(Replace(Norm, ChrW(700), "'"), ChrW(8217), "'"), ChrW(8216), "'"), "`", "'")
End Function

Private Function Num(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then Num = CDbl(v)
End Function
