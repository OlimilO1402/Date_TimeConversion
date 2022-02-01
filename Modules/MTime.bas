Attribute VB_Name = "MTime"
Option Explicit

Private Const HoursPerDay      As Long = 24&
Private Const MinutesPerHour   As Long = 60&
Private Const SecondsPerMinute As Long = 60&
Private Const SecondsPerDay    As Long = HoursPerDay * MinutesPerHour * SecondsPerMinute '86400

' Date:
' Enth�lt IEEE-64-Bit(8-Byte)-Werte, die Datumsangaben im Bereich vom 1. Januar des Jahres 0001 bis zum 31. Dezember
' des Jahres 9999 und Uhrzeiten von 00:00:00 Uhr (Mitternacht) bis 23:59:59.9999999 Uhr darstellen.
' Jedes Inkrement stellt 100 Nanosekunden verstrichener Zeit seit Beginn des 1. Januar des Jahres 1 im gregorianischen
' Kalender dar. Der maximale Wert stellt 100 Nanosekunden vor Beginn des 1. Januar des Jahres 10000 dar.
' Verwenden Sie den Date-Datentyp, um Datumswerte, Uhrzeitwerte oder Datums-und Uhrzeitwerte einzuschlie�en.
' Der Standardwert von Date ist 0:00:00 (Mitternacht) am 1. Januar 0001.
' Sie erhalten das aktuelle Datum und die aktuelle Uhrzeit aus der DateAndTime-Klasse. (VBA.DateTime)

'typedef struct _FILETIME {
'  DWORD dwLowDateTime;
'  DWORD dwHighDateTime;
'} FILETIME, *PFILETIME;
Public Type FILETIME
    dwLowDateTime  As Long ' 4
    dwHighDateTime As Long ' 4
End Type               'Sum: 8

'typedef struct _SYSTEMTIME {
'  WORD wYear;
'  WORD wMonth;
'  WORD wDayOfWeek;
'  WORD wDay;
'  WORD wHour;
'  WORD wMinute;
'  WORD wSecond;
'  WORD wMilliseconds;
'} SYSTEMTIME, *PSYSTEMTIME;
Public Type SYSTEMTIME
    wYear         As Integer ' 2
    wMonth        As Integer ' 2
    wDayOfWeek    As Integer ' 2
    wDay          As Integer ' 2
    wHour         As Integer ' 2
    wMinute       As Integer ' 2
    wSecond       As Integer ' 2
    wMilliseconds As Integer ' 2
End Type               ' Sum: 16
Public Type DOSTIME
'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-dosdatetimetofiletime
    ' Bits    Description
    ' 0 - 4   Day of the month (1�31)
    ' 5 - 8   Month (1 = January, 2 = February, and so on)
    ' 9 -15   Year offset from 1980 (add 1980 to get actual year)
    wDate As Integer
    
    ' Bits Description
    ' 0 - 4   Second divided by 2
    ' 5 -10   Minute (0�59)
    '11 -15   Hour (0�23 on a 24-hour clock)
    wTime As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias         As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type
Private Declare Sub GetSystemTime Lib "kernel32" ( _
    lpSysTime As SYSTEMTIME)

Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
    lpFilTime As FILETIME, lpSysTime As SYSTEMTIME) As Long

Private Declare Function SystemTimeToFileTime Lib "kernel32" ( _
    lpSysTime As SYSTEMTIME, lpFilTime As FILETIME) As Long

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" ( _
    lpFilTime As FILETIME, lpLocFilTime As FILETIME) As Long

Private Declare Function FileTimeToDosDateTime Lib "kernel32" ( _
    lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long

Private Declare Function DosDateTimeToFileTime Lib "kernel32" ( _
    ByVal wFatDate As Long, ByVal wFatTime As Long, lpFilTime As FILETIME) As Long

' ############################## '    DateTimeStamp    ' ############################## '
'can e.g. be found in executable files, exe, dll
Public Function DateTimeStamp_ToStr(ByVal DTStamp As Long) As String
    Dim l0  As Long:  l0 = DTStamp \ SecondsPerDay
    Dim l1  As Long:  l1 = DTStamp - l0 * SecondsPerDay
    Dim l2  As Long:  l2 = DateSerial(1970, 1, 2)
    Dim gmt As Date: gmt = l0 + Sgn(l1) + l1 / SecondsPerDay + l2
    DateTimeStamp_ToStr = Format$(gmt, "dd.mm.yyyy - hh:mm:ss")
End Function
Public Function DateTimeStamp_ToDate(ByVal DTStamp As Long) As Date
    Dim l0  As Long:  l0 = DTStamp \ SecondsPerDay
    Dim l1  As Long:  l1 = DTStamp - l0 * SecondsPerDay
    Dim l2  As Long:  l2 = DateSerial(1970, 1, 2)
    DateTimeStamp_ToDate = l0 + Sgn(l1) + l1 / SecondsPerDay + l2
End Function

' ############################## '        Date         ' ############################## '
Public Property Get Date_Now() As Date
    Date_Now = Now
End Property
Public Function Date_ToSystemTime(aDate As Date) As SYSTEMTIME
    With Date_ToSystemTime
        .wYear = Year(aDate)
        .wMonth = Month(aDate)
        .wDayOfWeek = Weekday(aDate, vbUseSystemDayOfWeek)
        .wDay = Day(aDate)
        .wHour = Hour(aDate)
        .wMinute = Minute(aDate)
        .wSecond = Second(aDate)
        '.wMilliseconds = millisecond(aDate) 'nope
    End With
End Function
Public Function Date_ToFileTime(aDate As Date) As FILETIME
    SystemTimeToFileTime Date_ToSystemTime(aDate), Date_ToFileTime
End Function
Public Function Date_ToUnixTime(aDate As Date) As Double
    Date_ToUnixTime = DateDiff("s", DateSerial(1970, 1, 1), aDate) - GetSummerTimeCorrector
End Function
Public Function Date_ToDosTime(aDate As Date) As DOSTIME
    Date_ToDosTime = FileTime_ToDosTime(Date_ToFileTime(aDate))
End Function
Public Function Date_ToDateTimeStamp(aDate As Date) As Long
    Date_ToDateTimeStamp = DateDiff("s", aDate, DateSerial(1970, 1, 2))
End Function
Public Function Date_ToStr(aDate As Date) As String
    Date_ToStr = FormatDateTime(aDate, VbDateTimeFormat.vbLongDate) & " " & FormatDateTime(aDate, VbDateTimeFormat.vbLongTime)
End Function
Public Function GetSummerTimeCorrector() As Double
    GetSummerTimeCorrector = DateDiff("s", SystemTime_ToDate(SystemTime_Now), Now)
End Function
Public Function Date_Equals(aDate As Date, other As Date) As Boolean
    Date_Equals = aDate = other
End Function

' ############################## '     SystemTime      ' ############################## '
Public Property Get SystemTime_Now() As SYSTEMTIME
    GetSystemTime SystemTime_Now
End Property
Public Function SystemTime_ToDate(aSt As SYSTEMTIME) As Date
    With aSt
        SystemTime_ToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function
Public Function SystemTime_ToFileTime(aSt As SYSTEMTIME) As FILETIME
    SystemTimeToFileTime aSt, SystemTime_ToFileTime
End Function

Public Function SystemTime_ToUnixTime(aSt As SYSTEMTIME) As Double
    SystemTime_ToUnixTime = Date_ToUnixTime(SystemTime_ToDate(aSt))
End Function
Public Function SystemTime_ToStr(aSt As SYSTEMTIME) As String
    With aSt
        SystemTime_ToStr = "y: " & CStr(.wYear) & "; m: " & CStr(.wMonth) & "; dow: " & CStr(.wDayOfWeek) & "; d: " & CStr(.wDay) & _
                         "; h: " & CStr(.wHour) & "; min: " & CStr(.wMinute) & "; s: " & CStr(.wSecond) & "; ms: " & CStr(.wMilliseconds)
    End With
End Function
Public Function SystemTime_Equals(aSt As SYSTEMTIME, other As SYSTEMTIME) As Boolean
    Dim b As Boolean
    With aSt
        b = .wYear = other.wYear:                 If Not b Then Exit Function
        b = .wMonth = other.wMonth:               If Not b Then Exit Function
        b = .wDay = other.wDay:                   If Not b Then Exit Function
        b = .wDayOfWeek = other.wDayOfWeek:       If Not b Then Exit Function
        b = .wHour = other.wHour:                 If Not b Then Exit Function
        b = .wMinute = other.wMinute:             If Not b Then Exit Function
        b = .wSecond = other.wSecond:             If Not b Then Exit Function
        b = .wMilliseconds = other.wMilliseconds ': If Not b Then Exit Function
    End With
    SystemTime_Equals = b
End Function

' ############################## '      FileTime       ' ############################## '
Public Property Get FileTime_Now() As FILETIME
    FileTime_Now = SystemTime_ToFileTime(SystemTime_Now)
End Property
Public Function FileTime_ToLocalFileTime(aFt As FILETIME) As FILETIME
    FileTimeToLocalFileTime aFt, FileTime_ToLocalFileTime
End Function

Public Property Get FileTime_ToDosTime(aFt As FILETIME) As DOSTIME
    Dim pdt As Long: pdt = VarPtr(FileTime_ToDosTime)
    FileTimeToDosDateTime aFt, pdt, pdt + 2
End Property

Public Function FileTime_ToDate(aFt As FILETIME) As Date
    With FileTime_ToSystemTime(aFt) 'st
        FileTime_ToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function
Public Function FileTime_ToSystemTime(aFt As FILETIME) As SYSTEMTIME
    FileTimeToSystemTime aFt, FileTime_ToSystemTime
End Function
Public Function FileTime_ToUnixTime(aFt As FILETIME) As Double
    FileTime_ToUnixTime = Date_ToUnixTime(FileTime_ToDate(aFt))
End Function
Public Function FileTime_ToStr(aFt As FILETIME) As String
    With aFt
        FileTime_ToStr = "lo: &&H" & Hex(.dwLowDateTime) & "; hi: &&H" & Hex(.dwHighDateTime)
    End With
End Function
Public Function FileTime_Equals(aFt As FILETIME, other As FILETIME) As Boolean
    Dim b As Boolean
    With aFt
        b = .dwHighDateTime = other.dwHighDateTime: If Not b Then Exit Function
        b = .dwLowDateTime = other.dwLowDateTime ':   If Not b Then Exit Function
    End With
    FileTime_Equals = b
End Function

' ############################## '      UnixTime       ' ############################## '
'In Unix und Linux werden Datumsangaben intern immer als die Anzahl der Sekunden seit
'dem 1. Januar 1970 um 00:00 Greenwhich Mean Time (GMT, heute UTC) dargestellt.
'Dieses Urdatum wird manchmal auch "The Epoch" genannt. In manchen Situationen muss
'man in Shellskripten die Unix-Zeit in ein normales Datum umrechnen und umgekehrt.
Public Property Get UnixTime_Now() As Double
    UnixTime_Now = Date_ToUnixTime(Date_Now)
End Property
Public Function UnixTime_ToDate(ByVal uts As Double) As Date
    UnixTime_ToDate = DateAdd("s", uts + GetSummerTimeCorrector, DateSerial(1970, 1, 1))
End Function
Public Function UnixTime_ToSystemTime(ByVal uts As Double) As SYSTEMTIME
    UnixTime_ToSystemTime = Date_ToSystemTime(UnixTime_ToDate(uts))
End Function
Public Function UnixTime_ToFileTime(ByVal uts As Double) As FILETIME
    UnixTime_ToFileTime = Date_ToFileTime(UnixTime_ToDate(uts))
End Function
Public Function UnixTime_ToStr(ByVal uts As Double) As String
    UnixTime_ToStr = CStr(uts)
End Function
Public Function UnixTime_Equals(uts As Double, other As Double) As Boolean
    UnixTime_Equals = uts = other
End Function

' ############################## '       DosTime       ' ############################## '
' oder auch FAT-Time also die Zeit die unter DOS in der FAT der Festplatte gespeichert wird
Public Function DosTime_Now() As DOSTIME
    DosTime_Now = FileTime_ToDosTime(Date_ToFileTime(Date_Now))
End Function
Public Property Get DosTime_ToFileTime(aDosTime As DOSTIME) As FILETIME
    DosDateTimeToFileTime aDosTime.wDate, aDosTime.wTime, DosTime_ToFileTime
End Property

Public Function DosTime_ToStr(aDt As DOSTIME) As String
    ' Bits    Description
    ' 0 - 4   Day of the month (1�31)
    ' 5 - 8   Month (1 = January, 2 = February, and so on)
    ' 9 -15   Year offset from 1980 (add 1980 to get actual year)
    'wDate As Integer
    ' Bits Description
    ' 0 - 4   Second divided by 2
    ' 5 -10   Minute (0�59)
    '11 -15   Hour (0�23 on a 24-hour clock)
    'wTime As Integer
    'DosTime_ToStr = SystemTime_ToStr(FileTime_ToSystemTime(DosTime_ToFileTime(aDosTime)))
'    Dim dm As Byte:    dm = (aDt.wDate And &H1F)
'    Dim mo As Byte:    mo = (aDt.wDate And &H1E0) \ 32
'    Dim ye As Integer: ye = (aDt.wDate And &H7F00) \ 512 + 1980
'
'    Dim se As Byte: se = (aDt.wTime And &H1F)
'    Dim mi As Byte: mi = (aDt.wTime And &H7E0) \ 32
'    Dim ho As Byte: ho = (aDt.wTime And &H7C00) \ 2048
'    DosTime_ToStr = Str2(dm) & "." & Str2(mo) & "." & CStr(ye) & " " & Str2(ho) & ":" & Str2(mi) & ":" & Str2(se)
    DosTime_ToStr = SystemTime_ToStr(FileTime_ToSystemTime(DosTime_ToFileTime(aDt)))
End Function
'Private Function Str2(ByVal by As Byte) As String
'    Str2 = CStr(by): If Len(Str2) < 2 Then Str2 = "0" & Str2
'End Function
Public Function DosTime_Equals(aDosTime As DOSTIME, other As DOSTIME) As Boolean
    Dim b As Boolean
    With aDosTime
        b = .wDate = other.wDate: If Not b Then Exit Function
        b = .wTime = other.wTime ': If Not b Then Exit Function
    End With
    DosTime_Equals = b
End Function

' ############################## '       CyTime        ' ############################## '
'contains time in milliseconds in a Currency, with maximum 4 digits
Public Function CyTime_FromSng(ms As Single) As Currency
    CyTime_FromSng = CCur(ms)
End Function
Public Function CyTime_FromDbl(ms As Double) As Currency
    CyTime_FromDbl = CCur(ms)
End Function
'Function CyTime_FromDate(aDate As Date) As Currency
'    'hmm m��te eigentlich sein Date_ToCyTime
'End Function
Public Function CyTime_ToStr(t As Currency) As String
    '1 day = 24 h =
    Dim yy As Integer
    Dim mo As Integer
    Dim dd As Integer
    Dim hh As Integer
    Dim mm As Integer
    Dim ss As Double
    CyTime_ToStr = CStr(ms)
End Function

' ############################## '       StrTime       ' ############################## '
' "hh:mm:ss.mls"
Public Function StrTime_ToCyTime(t As String) As Currency
    Dim sa() As String: sa = Split(t, ":")
    'Dim h  As Integer:  h = sa(0)
    'Dim m  As Integer:  m = sa(1)
    'Dim s  As Integer:  s = Int(Val(sa(2)))
    'Dim ms As Integer: ms = Split(sa(2), ".")(1)
    StrTime_ToCyTime = CLng(Split(sa(2), ".")(1)) + Int(Val(sa(2))) * 1000 + CLng(sa(1)) * 60 * 1000 + CLng(sa(0)) * 60 * 60 * 1000
End Function

Public Function StrTime_ToSYSTEMTIME(t As String) As SYSTEMTIME
    Dim sa() As String: sa = Split(t, ":")
    With StrTime_ToSYSTEMTIME
        .wYear = Year(Now)
        .wMonth = Month(Now)
        .wDay = Day(Now)
        .wHour = sa(0)
        .wMinute = sa(1)
        .wSecond = Int(Val(sa(2)))
        .wMilliseconds = Split(sa(2), ".")(1)
    End With
    'Debug.Print SystemTime_ToStr(MyTime_ToSYSTEMTIME)
End Function

