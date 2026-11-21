Attribute VB_Name = "Module1"
Option Explicit

' ===== Ulanish =====
' Tugma bosilganda: SQL Server login oynasi chiqadi, xodim o'z login-parolini kiritadi.
'   Parol hech qayerda saqlanmaydi.
' Botdan (silent) ishga tushganda: login-parol shu fayldan olinadi
'   (faqat bot ishlaydigan kompyuterda bo'ladi):
'   C:\ProgramData\BazaReport\db.ini
Private Const INI_PATH As String = "C:\ProgramData\BazaReport\db.ini"
Private Const DEF_SERVER As String = "SERVER01"
Private Const DEF_DB As String = "FOMBAZA"
Private mKochInfo As String

Private Function IniValue(ByVal key As String) As String
    Dim f As Integer, s As String, p As Long
    If Dir(INI_PATH) = "" Then Exit Function
    f = FreeFile
    Open INI_PATH For Input As #f
    Do While Not EOF(f)
        Line Input #f, s
        p = InStr(s, "=")
        If p > 0 Then
            If LCase$(Trim$(Left$(s, p - 1))) = LCase$(key) Then
                IniValue = Trim$(Mid$(s, p + 1))
                Exit Do
            End If
        End If
    Loop
    Close #f
End Function

Private Function BaseConn() As String
    Dim srv As String, db As String
    srv = IniValue("Server"): If srv = "" Then srv = DEF_SERVER
    db = IniValue("Database"): If db = "" Then db = DEF_DB
    BaseConn = "Driver={SQL Server};Server=" & srv & ";Database=" & db & ";"
End Function

' Bot uchun: login-parol db.ini dan
Private Function BotConn() As String
    If Dir(INI_PATH) = "" Then
        Err.Raise vbObjectError + 1, , "Sozlama fayli topilmadi: " & INI_PATH
    End If
    If IniValue("Uid") = "" Then
        BotConn = BaseConn() & "Trusted_Connection=yes;"
    Else
        BotConn = BaseConn() & "Uid=" & IniValue("Uid") & ";Pwd=" & IniValue("Pwd") & ";"
    End If
End Function

Private Function StatusPath() As String
    StatusPath = ThisWorkbook.Path & "\refresh_status.txt"
End Function

Private Sub LogLine(ByVal s As String)
    Dim f As Integer
    On Error Resume Next
    s = Replace(Replace(s, vbCr, " "), vbLf, " ")
    f = FreeFile
    Open StatusPath() For Append As #f
    Print #f, s
    Close #f
End Sub

Public Sub RefreshAll(Optional ByVal silent As Boolean = False)
    Dim cn As Object, rs As Object, ws As Worksheet
    Dim b As String, e As String, lastRow As Long, errs As String, n As Long

    Set ws = ThisWorkbook.Worksheets(1)

    On Error Resume Next
    Kill StatusPath()
    On Error GoTo 0
    LogLine "START " & Format(Now, "yyyy-mm-dd hh:nn:ss")

    ' Botdan ishga tushganda sana oralig'i bugunga suriladi (davr uzunligi saqlanadi)
    If silent Then
        n = CLng(ws.Range("DateEnd").Value) - CLng(ws.Range("DateBegin").Value)
        If n <= 0 Then n = 30
        ws.Range("DateEnd").Value = Date
        ws.Range("DateBegin").Value = Date - n
    End If

    b = Format(ws.Range("DateBegin").Value, "yyyymmdd")
    e = Format(ws.Range("DateEnd").Value, "yyyymmdd")

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set cn = CreateObject("ADODB.Connection")
    cn.CommandTimeout = 300
    If silent Then
        cn.ConnectionString = BotConn()
    Else
        ' Login/parolsiz ulanish satri + so'rash rejimi = standart SQL Server login oynasi
        cn.ConnectionString = BaseConn()
        cn.Properties("Prompt") = 2   ' adPromptComplete
    End If
    cn.Open

    ' ---- 1) SALE (asosiy varaq) ----
    Set rs = cn.Execute("SET NOCOUNT ON; EXEC dbo.XLS_OST_SALE3 '" & b & "','" & e & "'")
    ws.Range("A8:L" & ws.Rows.Count).ClearContents
    ws.Range("M9:S" & ws.Rows.Count).ClearContents
    ws.Range("A8").CopyFromRecordset rs
    rs.Close
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 8 Then lastRow = 8
    ws.Range("M8").Formula = "=E8+G8+I8+K8"
    ws.Range("N8").Formula = "=F8+H8+J8"
    ws.Range("O8").Formula = "=M8-N8"
    ws.Range("P8").Formula = "=N8/3"
    ws.Range("Q8").Formula = "=M8-P8"
    ws.Range("R8").Formula = "=N8/2"
    ws.Range("S8").Formula = "=M8-R8"
    ws.Range("M8:S" & lastRow).FillDown
    ws.Range("A8:S8").Copy
    ws.Range("A9:S" & lastRow).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    ws.Range("L8:L" & lastRow).NumberFormat = "dd.mm.yyyy"
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A7:S" & lastRow).AutoFilter
    LogLine "OK Sale|" & (lastRow - 7)

    ' ---- 2) qolgan varaqlar (biri xato bersa, boshqalari davom etadi) ----
    errs = ""
    errs = errs & LoadSheet(cn, "Srok", "EXEC dbo.XLS_BAZA_SROK", "J")
    errs = errs & LoadSheet(cn, "30+ kun", "EXEC dbo.XLS_BAZA_OBOROT 30", "I")
    errs = errs & LoadSheet(cn, "Narx", "EXEC dbo.XLS_BAZA_NARX", "F")
    errs = errs & LoadSheet(cn, "Hisob", "EXEC dbo.XLS_BAZA_HISOB", "D")
    cn.Close

    ' ---- 3) Ko'chirish va Skladdan varaqlari (bazasiz, yuqoridagi varaqlardan hisoblanadi) ----
    errs = errs & RunKochirish()

    Application.ScreenUpdating = True
    If silent Then
        ThisWorkbook.Save
    Else
        ws.Activate
    End If
    LogLine "DONE"
    If Not silent Then
        If errs = "" Then
            MsgBox "Tayyor! Sale: " & (lastRow - 7) & " qator. Barcha varaqlar yangilandi." & vbCrLf & mKochInfo
        Else
            MsgBox "Sale: " & (lastRow - 7) & " qator yangilandi." & vbCrLf & "Xatolar:" & vbCrLf & errs
        End If
    End If
    Exit Sub
Fail:
    Dim msg As String
    msg = Err.Description
    Application.ScreenUpdating = True
    LogLine "FAIL " & msg
    On Error Resume Next
    If Not cn Is Nothing Then If cn.State <> 0 Then cn.Close
    On Error GoTo 0
    If Not silent Then MsgBox "Xato: " & msg
End Sub

Private Function LoadSheet(cn As Object, shName As String, sqlText As String, lastCol As String) As String
    Dim w As Worksheet, rs As Object, lr As Long, cnt As Long
    On Error GoTo Bad
    Set w = ThisWorkbook.Worksheets(shName)
    Set rs = cn.Execute("SET NOCOUNT ON; " & sqlText)
    w.Range("A5:" & lastCol & w.Rows.Count).ClearContents
    w.Range("A5").CopyFromRecordset rs
    rs.Close
    lr = w.Cells(w.Rows.Count, 1).End(xlUp).Row
    If lr < 5 Then lr = 5
    If w.AutoFilterMode Then w.AutoFilterMode = False
    w.Range("A4:" & lastCol & lr).AutoFilter
    w.Range("B2").Value = Now
    cnt = 0
    If Len(CStr(w.Range("A5").Value)) > 0 Then cnt = lr - 4
    LogLine "OK " & shName & "|" & cnt
    LoadSheet = ""
    Exit Function
Bad:
    LogLine "ERR " & shName & "|" & Err.Description
    LoadSheet = " - " & shName & ": " & Err.Description & vbCrLf
End Function

Private Function RunKochirish() As String
    Dim w As Worksheet, lr As Long, cnt As Long, d As String
    On Error GoTo Bad
    mKochInfo = BuildKochirish()
    Set w = ThisWorkbook.Worksheets("Ko'chirish")
    lr = w.Cells(w.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then cnt = lr - 1
    LogLine "OK Ko'chirish|" & cnt
    RunKochirish = ""
    Exit Function
Bad:
    d = Err.Description
    mKochInfo = ""
    LogLine "ERR Ko'chirish|" & d
    RunKochirish = " - Ko'chirish: " & d & vbCrLf
End Function
