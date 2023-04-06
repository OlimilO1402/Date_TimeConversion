Attribute VB_Name = "MTime"
Option Explicit

Public Const HoursPerDay      As Long = 24&
Public Const MinutesPerHour   As Long = 60&
Public Const SecondsPerMinute As Long = 60&
Public Const SecondsPerHour   As Long = SecondsPerMinute * MinutesPerHour '  3600
Public Const SecondsPerDay    As Long = SecondsPerHour * HoursPerDay      ' 86400
Public Const MillisecondsPerSecond     As Long = 1000&
Public Const MillisecondsPerMinute     As Long = MillisecondsPerSecond * SecondsPerMinute '   60000
Public Const MillisecondsPerHour       As Long = MillisecondsPerSecond * SecondsPerHour   ' 3600000
Public Const MillisecondsPerDay        As Long = MillisecondsPerHour * HoursPerDay
Public Const NanosecondsPerMillisecond As Long = 1000000    ' = 1 million
Public Const NanosecondsPerSecond      As Long = 1000000000 ' = 1 billion 'deutsch: 1 Milliarde
Public Const NanosecondsPerTick        As Long = 100&
Public Const TicksPerMillisecond       As Long = 10000&    'ten-thousand zehntausend
Public Const TicksPerSecond            As Long = 10000000 'MillisecondsPerSecond * TicksPerMillisecond ' = 1 000 * 10 000 = 10 000 000 ' = 10 millions

' Date:
' Enth‰lt IEEE-64-Bit(8-Byte)-Werte, die Datumsangaben im Bereich vom 1. Januar des Jahres 0001 bis zum 31. Dezember
' des Jahres 9999 und Uhrzeiten von 00:00:00 Uhr (Mitternacht) bis 23:59:59.9999999 Uhr darstellen.
' Jedes Inkrement stellt 100 Nanosekunden verstrichener Zeit seit Beginn des 1. Januar des Jahres 1 im gregorianischen
' Kalender dar. Der maximale Wert stellt 100 Nanosekunden vor Beginn des 1. Januar des Jahres 10000 dar.
' Verwenden Sie den Date-Datentyp, um Datumswerte, Uhrzeitwerte oder Datums-und Uhrzeitwerte einzuschlieﬂen.
' Der Standardwert von Date ist 0:00:00 (Mitternacht) am 1. Januar 0001.
' Sie erhalten das aktuelle Datum und die aktuelle Uhrzeit aus der DateAndTime-Klasse. (VBA.DateTime)


'typedef struct _FILETIME {
'  DWORD dwLowDateTime;
'  DWORD dwHighDateTime;
'} FILETIME, *PFILETIME;
'https://learn.microsoft.com/en-us/windows/win32/api/minwinbase/ns-minwinbase-filetime
'Contains a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (UTC).
Public Type FILETIME
    dwLowDateTime  As Long ' 4
    dwHighDateTime As Long ' 4
End Type               'Sum: 8

'https://learn.microsoft.com/de-de/uwp/api/windows.foundation.datetime?view=winrt-22621
'UniversalTime: A 64-bit signed integer that represents a point in time as the number of 100-nanosecond intervals prior
'to or after midnight on January 1, 1601 (according to the Gregorian Calendar).
Public Type WindowsFoundationDateTime
    UniversalTime As Currency ' 8
End Type                 ' Sum: 8

'https://docs.microsoft.com/en-us/windows/win32/api/minwinbase/ns-minwinbase-systemtime
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
    ' 0 - 4   Day of the month (1ñ31)
    ' 5 - 8   Month (1 = January, 2 = February, and so on)
    ' 9 -15   Year offset from 1980 (add 1980 to get actual year)
    wDate As Integer ' 2
    
    ' Bits Description
    ' 0 - 4   Second divided by 2
    ' 5 -10   Minute (0ñ59)
    '11 -15   Hour (0ñ23 on a 24-hour clock)
    wTime As Integer ' 2
End Type        ' Sum: 4

Private Type TIME_ZONE_INFORMATION
    Bias                  As Long
    StandardName(1 To 64) As Byte
    StandardDate          As SYSTEMTIME
    StandardBias          As Long
    DaylightName(1 To 64) As Byte
    DaylightDate          As SYSTEMTIME
    DaylightBias          As Long
End Type

Private Type THexLng
    Value As Long
End Type

Private Type THexDbl
    Value As Double
End Type

Private Type THexDat
    Value As Date
End Type

Private Type THexBytes
    Value(0 To 15) As Byte
End Type

Const TIME_ZONE_ID_UNKNOWN  As Long = &H0&
Const TIME_ZONE_ID_STANDARD As Long = &H1&
Const TIME_ZONE_ID_DAYLIGHT As Long = &H2&

Private m_TZI As TIME_ZONE_INFORMATION
Public IsSummerTime As Boolean

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    
Private Declare Sub GetSystemTime Lib "kernel32" (lpSysTime As SYSTEMTIME)

Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFilTime As FILETIME, lpSysTime As SYSTEMTIME) As Long

Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSysTime As SYSTEMTIME, lpFilTime As FILETIME) As Long

'Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFilTime As FILETIME, lpLocFilTime As FILETIME) As Long
'Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocFilTime As FILETIME, lpFilTime As FILETIME) As Long

Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION, ByRef lpUniversalTime As SYSTEMTIME, ByRef lpLocalTime As SYSTEMTIME) As Long

Private Declare Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION, ByRef lpLocalTime As SYSTEMTIME, ByRef lpUniversalTime As SYSTEMTIME) As Long

Private Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long

Private Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFilTime As FILETIME) As Long

'void GetSystemTimePreciseAsFileTime(
'  [out] LPFILETIME lpSystemTimeAsFileTime
');
Private Declare Sub GetSystemTimePreciseAsFileTime Lib "kernel32" (lpSystemTimeAsFileTime As FILETIME)
'Private Declare Sub GetSystemTimePreciseAsFileTimeCy Lib "kernel32" Alias "GetSystemTimePreciseAsFileTime" (lpSystemTimeAsFileTime As Currency)

Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount_out As Currency) As Long
'

Public Sub Init()
    Dim ret As Long: ret = GetTimeZoneInformation(m_TZI)
    IsSummerTime = ret = TIME_ZONE_ID_DAYLIGHT
    If IsSummerTime Or ret = TIME_ZONE_ID_STANDARD Or ret = TIME_ZONE_ID_UNKNOWN Then Exit Sub
    MsgBox "Error trying to get time-zone-info!"
    'Select Case ret
    'Case TIME_ZONE_ID_UNKNOWN:  IsSummerTime = False
    'Case TIME_ZONE_ID_STANDARD: IsSummerTime = False
    'Case TIME_ZONE_ID_DAYLIGHT: IsSummerTime = True
    'Case Else: MsgBox "Error trying to get time-zone-imfo!"
    'End Select
End Sub

'Public Function GetSystemUpTime() As Date
Public Function GetSystemUpTime() As String
    'Returns the timespan since the last new start of your pc
    Dim ms As Currency 'milliseconds since system new start
    QueryPerformanceCounter ms
    Dim d As Long: d = ms \ MillisecondsPerDay:     ms = ms - CCur(d) * CCur(MillisecondsPerDay)
    Dim h As Long: h = ms \ MillisecondsPerHour:    ms = ms - h * MillisecondsPerHour
    Dim m As Long: m = ms \ MillisecondsPerMinute:  ms = ms - m * MillisecondsPerMinute
    Dim s As Long: s = ms \ MillisecondsPerSecond:  ms = ms - s * MillisecondsPerSecond
    GetSystemUpTime = d & ":" & Format(h, "00") & ":" & Format(m, "00") & ":" & Format(s, "00") & "." & Format(ms, "000")
End Function

'    Dim d As Date ' empty date!
'    MsgBox FormatDateTime(d, VBA.VbDateTimeFormat.vbLongDate) & " " & FormatDateTime(d, VBA.VbDateTimeFormat.vbLongTime)
'    'Samstag, 30. Dezember 1899 00:00:00

Public Function GetPCStartTime() As Date
    Dim ms As Currency 'milliseconds since system new start
    QueryPerformanceCounter ms
    Dim d As Long: d = ms \ MillisecondsPerDay:     ms = ms - d * MillisecondsPerDay
    Dim h As Long: h = ms \ MillisecondsPerHour:    ms = ms - h * MillisecondsPerHour
    Dim m As Long: m = ms \ MillisecondsPerMinute:  ms = ms - m * MillisecondsPerMinute
    Dim s As Long: s = ms \ MillisecondsPerSecond:  ms = ms - s * MillisecondsPerSecond
    GetPCStartTime = VBA.DateTime.Now - DateSerial(1900, 1, d - 1) - TimeSerial(h, m, s)
End Function

Public Function TimeZoneInfo_ToStr() As String
    Dim s As String, s1 As String
    With m_TZI
        s = s & "Bias         : " & .Bias & vbCrLf
        
        s1 = Trim0(.StandardName)
        
        s = s & "StandardName : " & s1 & vbCrLf
        s = s & "StandardDate : " & SystemTime_ToDate(.StandardDate) & vbCrLf
        s = s & "StandardBias : " & .StandardBias & vbCrLf
        
        s1 = Trim0(.DaylightName)
        
        s = s & "DaylightName : " & s1 & vbCrLf
        s = s & "DaylightDate : " & SystemTime_ToDate(.DaylightDate) & vbCrLf
        s = s & "DaylightBias : " & .DaylightBias & vbCrLf
    End With
    TimeZoneInfo_ToStr = s
End Function

Function Trim0(ByVal s As String) As String
    Trim0 = Trim(s)
    If Right(Trim0, 1) = vbNullChar Then
        Trim0 = Left(Trim0, Len(Trim0) - 1)
        Trim0 = Trim0(Trim0)
    'Else
    '    Exit Function
    End If
End Function

' ############################## '    DateTimeStamp    ' ############################## '
' e.g. can be found in executable files, exe, dll
Public Function DateTimeStamp_Now() As Long
    DateTimeStamp_Now = Date_ToDateTimeStamp(Date_Now)
End Function

Public Function DateTimeStamp_ToDate(ByVal DtStamp As Long) As Date
    Dim l0  As Long:  l0 = DtStamp \ SecondsPerDay
    Dim l1  As Long:  l1 = DtStamp - l0 * SecondsPerDay
    Dim l2  As Long:  l2 = DateSerial(1970, 1, 2)
    DateTimeStamp_ToDate = l0 + Sgn(l1) + l1 / SecondsPerDay + l2
End Function

Public Function DateTimeStamp_ToSystemTime(ByVal DtStamp As Long) As SYSTEMTIME
    DateTimeStamp_ToSystemTime = Date_ToSystemTime(DateTimeStamp_ToDate(DtStamp))
End Function

Public Function DateTimeStamp_ToFileTime(ByVal DtStamp As Long) As FILETIME
    DateTimeStamp_ToFileTime = Date_ToFileTime(DateTimeStamp_ToDate(DtStamp))
End Function

Public Function DateTimeStamp_ToUnixTime(ByVal DtStamp As Long) As Double
    DateTimeStamp_ToUnixTime = Date_ToUnixTime(DateTimeStamp_ToDate(DtStamp))
End Function

Public Function DateTimeStamp_ToDosTime(ByVal DtStamp As Long) As DOSTIME
    DateTimeStamp_ToDosTime = Date_ToDosTime(DateTimeStamp_ToDate(DtStamp))
End Function

Public Function DateTimeStamp_ToWindowsFoundationDateTime(ByVal DtStamp As Long) As WindowsFoundationDateTime
    DateTimeStamp_ToWindowsFoundationDateTime = Date_ToWindowsFoundationDateTime(DateTimeStamp_ToDate(DtStamp))
End Function

Public Function DateTimeStamp_ToStr(ByVal DtStamp As Long) As String
    'Dim l0  As Long:  l0 = DTStamp \ SecondsPerDay
    'Dim l1  As Long:  l1 = DTStamp - l0 * SecondsPerDay
    'Dim l2  As Long:  l2 = DateSerial(1970, 1, 2)
    'Dim gmt As Date: gmt = l0 + Sgn(l1) + l1 / SecondsPerDay + l2
    'DateTimeStamp_ToStr = Format$(gmt, "yyyy.mm.dd - hh:mm:ss")
    DateTimeStamp_ToStr = Format$(DateTimeStamp_ToDate(DtStamp), "yyyy.mm.dd - hh:mm:ss")
    'DateTimeStamp_ToStr = "&&H" & Hex(DTStamp)
End Function

Public Function DateTimeStamp_ToHex(ByVal DtStamp As Long) As String
    Dim th As THexBytes, tl As THexLng: tl.Value = DtStamp: LSet th = tl
    DateTimeStamp_ToHex = THexBytes_ToStr(th)
End Function

Public Function DateTimeStamp_ToHexNStr(ByVal DtStamp As Long) As String
    DateTimeStamp_ToHexNStr = DateTimeStamp_ToHex(DtStamp) & "; " & DateTimeStamp_ToStr(DtStamp)
End Function

' ############################## '        Date         ' ############################## '
Public Property Get Date_Now() As Date
    Date_Now = VBA.DateTime.Now
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
    Date_ToDateTimeStamp = DateDiff("s", DateSerial(1970, 1, 2), aDate)
End Function

Public Function Date_ToWindowsFoundationDateTime(aDate As Date) As WindowsFoundationDateTime
    LSet Date_ToWindowsFoundationDateTime = Date_ToFileTime(aDate)
End Function

Public Function Date_ToStr(aDate As Date) As String
    Date_ToStr = FormatDateTime(aDate, VbDateTimeFormat.vbLongDate) & " - " & FormatDateTime(aDate, VbDateTimeFormat.vbLongTime)
End Function

Public Function GetSummerTimeCorrector() As Double
    GetSummerTimeCorrector = DateDiff("s", SystemTime_ToDate(SystemTime_Now), Now)
End Function

Public Function Date_Equals(aDate As Date, other As Date) As Boolean
    Date_Equals = aDate = other
End Function

Public Function Date_ToHex(ByVal aDate As Date) As String
    Dim th As THexBytes, td As THexDat: td.Value = aDate: LSet th = td
    Date_ToHex = THexBytes_ToStr(th)
End Function

Public Function Date_ToHexNStr(ByVal aDate As Date) As String
    Date_ToHexNStr = Date_ToHex(aDate) & "; " & Date_ToStr(aDate)
End Function

' ############################## '     SystemTime      ' ############################## '
Public Property Get SystemTime_Now() As SYSTEMTIME
    GetSystemTime SystemTime_Now
    SystemTime_Now = SystemTime_ToTzSpecificLocalTime(SystemTime_Now)
End Property

Public Function SystemTime_ToTzSpecificLocalTime(aSt As SYSTEMTIME) As SYSTEMTIME
    SystemTimeToTzSpecificLocalTime m_TZI, aSt, SystemTime_ToTzSpecificLocalTime
End Function
Public Function TzSpecificLocalTime_ToSystemTime(aSt As SYSTEMTIME) As SYSTEMTIME
    TzSpecificLocalTimeToSystemTime m_TZI, aSt, TzSpecificLocalTime_ToSystemTime
End Function

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

Public Function SystemTime_ToDosTime(aSt As SYSTEMTIME) As DOSTIME
    SystemTime_ToDosTime = Date_ToDosTime(SystemTime_ToDate(aSt))
End Function

Public Function SystemTime_ToWindowsFoundationDateTime(aSt As SYSTEMTIME) As WindowsFoundationDateTime
    LSet SystemTime_ToWindowsFoundationDateTime = SystemTime_ToFileTime(aSt)
End Function

Public Function SystemTime_ToDateTimeStamp(aSt As SYSTEMTIME) As Long
    SystemTime_ToDateTimeStamp = Date_ToDateTimeStamp(SystemTime_ToDate(aSt))
End Function

Public Function SystemTime_ToStr(aSt As SYSTEMTIME) As String
    With aSt
        SystemTime_ToStr = "y: " & CStr(.wYear) & "; m: " & CStr(.wMonth) & "; d: " & CStr(.wDay) & "; dow: " & CStr(.wDayOfWeek) & _
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

Public Function SystemTime_ToHex(aSt As SYSTEMTIME) As String
    Dim th As THexBytes: LSet th = aSt
    SystemTime_ToHex = THexBytes_ToStr(th)
End Function

Public Function SystemTime_ToHexNStr(aSt As SYSTEMTIME) As String
    SystemTime_ToHexNStr = SystemTime_ToHex(aSt) & "; " & SystemTime_ToStr(aSt)
End Function

' ############################## '      FileTime       ' ############################## '
Public Property Get FileTime_Now() As FILETIME
    'FileTime_Now = SystemTime_ToFileTime(SystemTime_Now)
    GetSystemTimePreciseAsFileTime FileTime_Now
    FileTime_Now = FileTime_ToLocalFileTime(FileTime_Now)
End Property

Public Function FileTime_ToLocalFileTime(aFt As FILETIME) As FILETIME
'    FileTimeToLocalFileTime aFt, FileTime_ToLocalFileTime
    Dim st_in As SYSTEMTIME: st_in = FileTime_ToSystemTime(aFt)
    Dim stout As SYSTEMTIME
    SystemTimeToTzSpecificLocalTime m_TZI, st_in, stout
    FileTime_ToLocalFileTime = SystemTime_ToFileTime(stout)
End Function

Public Function LocalFileTime_ToFileTime(aFt As FILETIME) As FILETIME
'    LocalFileTimeToFileTime aFt, LocalFileTime_ToFileTime
    Dim st_in As SYSTEMTIME: st_in = FileTime_ToSystemTime(aFt)
    Dim stout As SYSTEMTIME
    TzSpecificLocalTimeToSystemTime m_TZI, st_in, stout
    LocalFileTime_ToFileTime = SystemTime_ToFileTime(stout)
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

Public Function FileTime_ToWindowsFoundationDateTime(aFt As FILETIME) As WindowsFoundationDateTime
    LSet FileTime_ToWindowsFoundationDateTime = aFt
End Function

Public Function FileTime_ToDateTimeStamp(aFt As FILETIME) As Long
    FileTime_ToDateTimeStamp = Date_ToDateTimeStamp(FileTime_ToDate(aFt))
End Function

Public Function FileTime_ToStr(aFt As FILETIME) As String
    With aFt
        FileTime_ToStr = "lo: " & CStr(.dwLowDateTime) & "; hi: " & CStr(.dwHighDateTime)
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

Public Function FileTime_ToHex(aFt As FILETIME) As String
    Dim th As THexBytes: LSet th = aFt
    FileTime_ToHex = THexBytes_ToStr(th)
End Function

Public Function FileTime_ToHexNStr(aFt As FILETIME) As String
    FileTime_ToHexNStr = FileTime_ToHex(aFt) & "; " & FileTime_ToStr(aFt)
End Function

' ############################## '      UnixTime       ' ############################## '
' In Unix und Linux werden Datumsangaben intern immer als die Anzahl der Sekunden seit
' dem 1. Januar 1970 um 00:00 Greenwhich Mean Time (GMT, heute UTC) dargestellt.
' Dieses Urdatum wird manchmal auch "The Epoch" genannt. In manchen Situationen muss
' man in Shellskripten die Unix-Zeit in ein normales Datum umrechnen und umgekehrt.
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

Public Function UnixTime_ToDosTime(ByVal uts As Double) As DOSTIME
    UnixTime_ToDosTime = MTime.Date_ToDosTime(MTime.UnixTime_ToDate(uts))
End Function

Public Function UnixTime_ToWindowsFoundationDateTime(ByVal uts As Double) As WindowsFoundationDateTime
    LSet UnixTime_ToWindowsFoundationDateTime = UnixTime_ToFileTime(uts)
End Function

Public Function UnixTime_ToDateTimeStamp(ByVal uts As Double) As Long
    UnixTime_ToDateTimeStamp = Date_ToDateTimeStamp(UnixTime_ToDate(uts))
End Function

Public Function UnixTime_ToStr(ByVal uts As Double) As String
    UnixTime_ToStr = CStr(uts)
End Function

Public Function UnixTime_Equals(ByVal uts As Double, ByVal other As Double) As Boolean
    UnixTime_Equals = uts = other
End Function

Public Function UnixTime_ToHex(ByVal uts As Double) As String
    Dim th As THexBytes, td As THexDbl: td.Value = uts: LSet th = td
    UnixTime_ToHex = THexBytes_ToStr(th)
End Function

Public Function UnixTime_ToHexNStr(ByVal uts As Double) As String
    UnixTime_ToHexNStr = UnixTime_ToHex(uts) & "; " & UnixTime_ToStr(uts)
End Function

' ############################## '       DosTime       ' ############################## '
' oder auch FAT-Time also die Zeit die unter DOS in der FAT der Festplatte gespeichert wird
Public Function DosTime_Now() As DOSTIME
    DosTime_Now = FileTime_ToDosTime(Date_ToFileTime(Date_Now))
End Function

Public Function DosTime_ToDate(aDosTime As DOSTIME) As Date
    DosTime_ToDate = FileTime_ToDate(DosTime_ToFileTime(aDosTime))
End Function

Public Function DosTime_ToSystemTime(aDosTime As DOSTIME) As SYSTEMTIME
    DosTime_ToSystemTime = FileTime_ToSystemTime(DosTime_ToFileTime(aDosTime))
End Function

Public Property Get DosTime_ToFileTime(aDosTime As DOSTIME) As FILETIME
    DosDateTimeToFileTime aDosTime.wDate, aDosTime.wTime, DosTime_ToFileTime
End Property

Public Function DosTime_ToUnixTime(aDosTime As DOSTIME) As Double
    DosTime_ToUnixTime = Date_ToUnixTime(DosTime_ToDate(aDosTime))
End Function

Public Function DosTime_ToWindowsFoundationDateTime(aDosTime As DOSTIME) As WindowsFoundationDateTime
    LSet DosTime_ToWindowsFoundationDateTime = DosTime_ToFileTime(aDosTime)
End Function

Public Function DosTime_ToDateTimeStamp(aDosTime As DOSTIME) As Long
    DosTime_ToDateTimeStamp = Date_ToDateTimeStamp(DosTime_ToDate(aDosTime))
End Function

Public Function DosTime_ToStr(aDt As DOSTIME) As String
    ' Bits    Description
    ' 0 - 4   Day of the month (1ñ31)
    ' 5 - 8   Month (1 = January, 2 = February, and so on)
    ' 9 -15   Year offset from 1980 (add 1980 to get actual year)
    'wDate As Integer
    ' Bits Description
    ' 0 - 4   Second divided by 2
    ' 5 -10   Minute (0ñ59)
    '11 -15   Hour (0ñ23 on a 24-hour clock)
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
    DosTime_ToStr = "wDate: " & CStr(aDt.wDate) & "; wTime: " & CStr(aDt.wTime)
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

Public Function DosTime_ToHex(aDosTime As DOSTIME) As String
    Dim th As THexBytes: LSet th = aDosTime
    DosTime_ToHex = THexBytes_ToStr(th)
End Function

Public Function DosTime_ToHexNStr(aDosTime As DOSTIME) As String
    DosTime_ToHexNStr = DosTime_ToHex(aDosTime) & "; " & DosTime_ToStr(aDosTime)
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
'    'hmm m¸ﬂte eigentlich sein Date_ToCyTime
'End Function
Public Function CyTime_ToStr(aCyTime As Currency) As String
    '1 day = 24 h =
    Dim yy As Integer '=
    Dim mo As Integer
    Dim dd As Integer
    Dim hh As Integer
    Dim mm As Integer
    Dim ss As Double
    'CyTime_ToStr = CStr(ms)
End Function

' ############################## '       StrTime       ' ############################## '
' "hh:mm:ss.mls"
Public Function StrTime_ToCyTime(T As String) As Currency
    Dim sa() As String: sa = Split(T, ":")
    'Dim h  As Integer:  h = sa(0)
    'Dim m  As Integer:  m = sa(1)
    'Dim s  As Integer:  s = Int(Val(sa(2)))
    'Dim ms As Integer: ms = Split(sa(2), ".")(1)
    StrTime_ToCyTime = CLng(Split(sa(2), ".")(1)) + Int(Val(sa(2))) * 1000 + CLng(sa(1)) * 60 * 1000 + CLng(sa(0)) * 60 * 60 * 1000
End Function

Public Function StrTime_ToSYSTEMTIME(T As String) As SYSTEMTIME
    Dim sa() As String: sa = Split(T, ":")
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

' ############################## '       WindowsFoundationDateTime       ' ############################## '
'https://learn.microsoft.com/de-de/uwp/api/windows.foundation.datetime?view=winrt-22621
'UniversalTime: A 64-bit signed integer that represents a point in time as the number of 100-nanosecond intervals prior to or after midnight on January 1, 1601 (according to the Gregorian Calendar).

Public Function WindowsFoundationDateTime_Now() As WindowsFoundationDateTime
    LSet WindowsFoundationDateTime_Now = FileTime_Now
End Function

Public Function WindowsFoundationDateTime_ToFileTime(this As WindowsFoundationDateTime) As FILETIME
    LSet WindowsFoundationDateTime_ToFileTime = this
End Function

Public Function WindowsFoundationDateTime_ToSystemTime(this As WindowsFoundationDateTime) As SYSTEMTIME
    WindowsFoundationDateTime_ToSystemTime = FileTime_ToSystemTime(WindowsFoundationDateTime_ToFileTime(this))
End Function

Public Function WindowsFoundationDateTime_ToDate(this As WindowsFoundationDateTime) As Date
    WindowsFoundationDateTime_ToDate = FileTime_ToDate(WindowsFoundationDateTime_ToFileTime(this))
End Function

'11644473600
Public Function WindowsFoundationDateTime_ToUnixTime(this As WindowsFoundationDateTime) As Double
    WindowsFoundationDateTime_ToUnixTime = FileTime_ToUnixTime(WindowsFoundationDateTime_ToFileTime(this))
End Function

Public Function WindowsFoundationDateTime_ToDosTime(this As WindowsFoundationDateTime) As DOSTIME
    WindowsFoundationDateTime_ToDosTime = FileTime_ToDosTime(WindowsFoundationDateTime_ToFileTime(this))
End Function

Public Function WindowsFoundationDateTime_ToDateTimeStamp(this As WindowsFoundationDateTime) As Long
    WindowsFoundationDateTime_ToDateTimeStamp = Date_ToDateTimeStamp(WindowsFoundationDateTime_ToDate(this))
End Function

Public Function WindowsFoundationDateTime_ToStr(this As WindowsFoundationDateTime) As String
    WindowsFoundationDateTime_ToStr = "UniversalTime: " & CStr(this.UniversalTime)
End Function

Public Function WindowsFoundationDateTime_Equals(this As WindowsFoundationDateTime, other As WindowsFoundationDateTime) As Boolean
    WindowsFoundationDateTime_Equals = this.UniversalTime = other.UniversalTime
End Function

Public Function WindowsFoundationDateTime_ToHex(this As WindowsFoundationDateTime) As String
    Dim th As THexBytes: LSet th = this
    WindowsFoundationDateTime_ToHex = THexBytes_ToStr(th)
End Function

Public Function WindowsFoundationDateTime_ToHexNStr(this As WindowsFoundationDateTime) As String
    WindowsFoundationDateTime_ToHexNStr = WindowsFoundationDateTime_ToHex(this) & "; " & WindowsFoundationDateTime_ToStr(this)
End Function

' ############################## '       THexBytes       ' ############################## '
Private Function THexBytes_ToStr(this As THexBytes) As String
    Dim s As String: s = "&&H"
    Dim i As Long, u As Long: u = UBound(this.Value)
    For i = u To 0 Step -1
        s = s & Hex2(this.Value(i))
    Next
    THexBytes_ToStr = s
End Function
Private Function Hex2(ByVal b As Byte) As String
    Hex2 = Hex(b): If Len(Hex2) < 2 Then Hex2 = "0" & Hex2
End Function

