Attribute VB_Name = "MISO8601"
Option Explicit

Private Const FormatDTAlphabet  As String = "YMwDThmsf"
Private Const FormatDTSeparator As String = "PTW-+:,."
Private Const FormatTSSeparator As String = "YMWDTHMS.,/"

Public Function Dec2(ByVal Value As Long) As String
    Dim sign As Long: sign = Sgn(Value): Value = Abs(Value)
    Dim s As String: s = CStr(Value)
    Dim n As Long: n = Len(s)
    If n < 3 Then
        Dec2 = IIf(sign < 0, "-", "") & String(2 - n, "0") & s
    End If
End Function

Public Function Dec3(ByVal Value As Long) As String
    Dim sign As Long: sign = Sgn(Value): Value = Abs(Value)
    Dim s As String: s = CStr(Value)
    Dim n As Long: n = Len(s)
    If n < 4 Then
        Dec3 = IIf(sign < 0, "-", "") & String(3 - n, "0") & s
    End If
End Function

Public Function Dec4(ByVal Value As Long) As String
    Dim sign As Long: sign = Sgn(Value): Value = Abs(Value)
    Dim s As String: s = CStr(Value)
    Dim n As Long: n = Len(s)
    If n < 5 Then
        Dec4 = IIf(sign < 0, "-", "") & String(4 - n, "0") & s
    End If
End Function

Public Function FormatSystemTime(this As SYSTEMTIME, ByVal Format As String) As String
    
End Function
'    Dim dt As Date
'    dt = CDate("31.05.2024")
'    Debug.Print Format(dt, "YYYYMMDD")
'    Debug.Print Format(dt, "YYYY-MM-DD")
'    Dim tm As Date: tm = TimeSerial(16, 56, 12)
'    Dim dt2 As Date: dt2 = dt + tm
'    Debug.Print dt2
'    Debug.Print Format(dt2, "YYYY-MM-DDThh:mm:ss")
'20240531
'2024-05-31
'31.05.2024 16:56:12
'2024-05-31T16:56:12

Private Function NextState(ByVal oldstate As Long, ByVal ch As Long) As Long
    Select Case state
    Case 0
        Select Case ch
        Case 89:   state = 10: n = 1 '"Y"
        Case 77:   state = 20: n = 1 '"M"
        Case 68:   state = 30: n = 1 '"D"
        Case 104:  state = 40: n = 1 '"h"
        Case 109:  state = 50: n = 1 '"m"
        Case 115:  state = 60: n = 1 '"s"
        Case Else: state = -1: n = 0 'errorstate
        End Select
    Case 10
        Select Case ch
        Case 89:     n = n + 1
        Case Else: state = -1: n = 0  'errorstate
        End Select
    Case 4
    End Select
End Function

