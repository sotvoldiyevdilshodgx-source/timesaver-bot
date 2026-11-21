' TimeSaver bot: Excel faylni yashirin ochib, RefreshAll makrosini ishga tushiradi.
' Ishlatish: cscript //nologo refresh.vbs "E:\Hisobot\baza_yangi.xlsm" [kunlar]
'   kunlar (ixtiyoriy) - sotuv davri: DateBegin = bugun - kunlar, DateEnd = bugun (makros davrni shundan oladi).
Option Explicit
Dim path, xl, wb, days
If WScript.Arguments.Count < 1 Then
  WScript.Echo "Fayl yo'li berilmadi"
  WScript.Quit 1
End If
path = WScript.Arguments(0)
Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AutomationSecurity = 1
On Error Resume Next
Set wb = xl.Workbooks.Open(path, 0, False)
If Err.Number <> 0 Then
  WScript.Echo "OPEN: " & Err.Description
  xl.Quit
  WScript.Quit 2
End If
If wb.ReadOnly Then
  WScript.Echo "OPEN: fayl faqat o'qish rejimida (band)"
  wb.Close False
  xl.Quit
  WScript.Quit 3
End If
Err.Clear
days = 0
If WScript.Arguments.Count >= 2 Then
  If IsNumeric(WScript.Arguments(1)) Then days = CLng(WScript.Arguments(1))
End If
If days > 0 And days <= 3650 Then
  wb.Worksheets(1).Range("DateEnd").Value = Date
  wb.Worksheets(1).Range("DateBegin").Value = Date - days
  If Err.Number <> 0 Then
    WScript.Echo "DAYS: " & Err.Description
    Err.Clear
  End If
End If
xl.Run "'" & wb.Name & "'!RefreshAll", True
If Err.Number <> 0 Then
  WScript.Echo "RUN: " & Err.Description
  wb.Close False
  xl.Quit
  WScript.Quit 4
End If
wb.Close False
xl.Quit
WScript.Quit 0
