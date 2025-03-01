Attribute VB_Name = "MTime"
Option Explicit 'Lines: 1402 14.jun.2023; 1603 16.sep.2024

Public Enum ECalendar
    JulianCalendar
    GregorianCalendar
End Enum

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
' Enthält IEEE-64-Bit(8-Byte)-Werte, die Datumsangaben im Bereich vom 1. Januar des Jahres 0001 bis zum 31. Dezember
' des Jahres 9999 und Uhrzeiten von 00:00:00 Uhr (Mitternacht) bis 23:59:59.9999999 Uhr darstellen.
' Jedes Inkrement stellt 100 Nanosekunden verstrichener Zeit seit Beginn des 1. Januar des Jahres 1 im gregorianischen
' Kalender dar. Der maximale Wert stellt 100 Nanosekunden vor Beginn des 1. Januar des Jahres 10000 dar.
' Verwenden Sie den Date-Datentyp, um Datumswerte, Uhrzeitwerte oder Datums-und Uhrzeitwerte einzuschließen.
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
    ' 0 - 4   Day of the month (1–31)
    ' 5 - 8   Month (1 = January, 2 = February, and so on)
    ' 9 -15   Year offset from 1980 (add 1980 to get actual year)
    wDate As Integer ' 2
    
    ' Bits Description
    ' 0 - 4   Second divided by 2
    ' 5 -10   Minute (0–59)
    '11 -15   Hour (0–23 on a 24-hour clock)
    wTime As Integer ' 2
End Type        ' Sum: 4

Private Type TIME_ZONE_INFORMATION
    Bias                        As Long
    StandardName(1 To 64)       As Byte
    StandardDate                As SYSTEMTIME
    StandardBias                As Long
    DaylightName(1 To 64)       As Byte
    DaylightDate                As SYSTEMTIME
    DaylightBias                As Long
End Type

'typedef struct _TIME_DYNAMIC_ZONE_INFORMATION {
'  LONG       Bias;
'  WCHAR      StandardName[32];
'  SYSTEMTIME StandardDate;
'  LONG       StandardBias;
'  WCHAR      DaylightName[32];
'  SYSTEMTIME DaylightDate;
'  LONG       DaylightBias;
'  WCHAR      TimeZoneKeyName[128];
'  BOOLEAN    DynamicDaylightTimeDisabled;
'} DYNAMIC_TIME_ZONE_INFORMATION, *PDYNAMIC_TIME_ZONE_INFORMATION;

Private Type DYNAMIC_TIME_ZONE_INFORMATION
    TZI                         As TIME_ZONE_INFORMATION
    TimeZoneKeyName(1 To 256)   As Byte
    DynamicDaylightTimeDisabled As Long
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

Private m_TZI    As TIME_ZONE_INFORMATION
Private m_DynTZI As DYNAMIC_TIME_ZONE_INFORMATION
Public IsSummerTime As Boolean

#If VBA7 Then
    'https://learn.microsoft.com/en-us/windows/win32/api/timezoneapi/nf-timezoneapi-gettimezoneinformation
    Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    '
    Private Declare PtrSafe Function GetDynamicTimeZoneInformation Lib "kernel32" (pTimeZoneInformation As DYNAMIC_TIME_ZONE_INFORMATION) As Long
    Private Declare PtrSafe Function GetTimeZoneInformationForYear Lib "kernel32" (ByVal wYear As Integer, pdtzi As DYNAMIC_TIME_ZONE_INFORMATION, ptzi As TIME_ZONE_INFORMATION) As Long
    
    
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSysTime As SYSTEMTIME)
    
    Private Declare PtrSafe Function FileTimeToSystemTime Lib "kernel32" (lpFilTime As FILETIME, lpSysTime As SYSTEMTIME) As Long
    
    Private Declare PtrSafe Function SystemTimeToFileTime Lib "kernel32" (lpSysTime As SYSTEMTIME, lpFilTime As FILETIME) As Long
    
    'Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFilTime As FILETIME, lpLocFilTime As FILETIME) As Long
    'Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocFilTime As FILETIME, lpFilTime As FILETIME) As Long
    
    Private Declare PtrSafe Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long
    
    Private Declare PtrSafe Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpLocalTime As SYSTEMTIME, lpUniversalTime As SYSTEMTIME) As Long
    
    Private Declare PtrSafe Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, lpFatDate As Integer, lpFatTime As Integer) As Long
    
    Private Declare PtrSafe Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFilTime As FILETIME) As Long
    
    'void GetSystemTimePreciseAsFileTime(
    '  [out] LPFILETIME lpSystemTimeAsFileTime
    ');
    Private Declare PtrSafe Sub GetSystemTimePreciseAsFileTime Lib "kernel32" (lpSystemTimeAsFileTime As FILETIME)
    'Private Declare Sub GetSystemTimePreciseAsFileTimeCy Lib "kernel32" Alias "GetSystemTimePreciseAsFileTime" (lpSystemTimeAsFileTime As Currency)
    
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount_out As Currency) As Long
    
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency_out As Currency) As Long
#Else
    'https://learn.microsoft.com/en-us/windows/win32/api/timezoneapi/nf-timezoneapi-gettimezoneinformation
    Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    '
    Private Declare Function GetDynamicTimeZoneInformation Lib "kernel32" (pTimeZoneInformation As DYNAMIC_TIME_ZONE_INFORMATION) As Long
    Private Declare Function GetTimeZoneInformationForYear Lib "kernel32" (ByVal wYear As Integer, pdtzi As DYNAMIC_TIME_ZONE_INFORMATION, ptzi As TIME_ZONE_INFORMATION) As Long
    
    
    Private Declare Sub GetSystemTime Lib "kernel32" (lpSysTime As SYSTEMTIME)
    
    Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFilTime As FILETIME, lpSysTime As SYSTEMTIME) As Long
    
    Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSysTime As SYSTEMTIME, lpFilTime As FILETIME) As Long
    
    'Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFilTime As FILETIME, lpLocFilTime As FILETIME) As Long
    'Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocFilTime As FILETIME, lpFilTime As FILETIME) As Long
    
    Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long
    
    Private Declare Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpLocalTime As SYSTEMTIME, lpUniversalTime As SYSTEMTIME) As Long
    
    Private Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, lpFatDate As Integer, lpFatTime As Integer) As Long
    
    Private Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFilTime As FILETIME) As Long
    
    'void GetSystemTimePreciseAsFileTime(
    '  [out] LPFILETIME lpSystemTimeAsFileTime
    ');
    Private Declare Sub GetSystemTimePreciseAsFileTime Lib "kernel32" (lpSystemTimeAsFileTime As FILETIME)
    'Private Declare Sub GetSystemTimePreciseAsFileTimeCy Lib "kernel32" Alias "GetSystemTimePreciseAsFileTime" (lpSystemTimeAsFileTime As Currency)
    
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount_out As Currency) As Long
    
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency_out As Currency) As Long
#End If

Public Sub Init()
    
    Dim ret As Long
    'm_TZI.DaylightDate = SystemTime_Now
    'm_TZI.StandardDate = SystemTime_Now
    ret = GetTimeZoneInformation(m_TZI)
    IsSummerTime = ret = TIME_ZONE_ID_DAYLIGHT
    'Debug.Print "----------"
    'Debug.Print PTimeZoneInfo_ToStr(m_TZI)
    
    ret = GetDynamicTimeZoneInformation(m_DynTZI)
    IsSummerTime = ret = TIME_ZONE_ID_DAYLIGHT
    'Debug.Print "----------"
    'Debug.Print PDynTimeZoneInfo_ToStr(m_DynTZI)
    
    Dim y As Integer: y = DateTime.Year(Now)
    ret = GetTimeZoneInformationForYear(y, m_DynTZI, m_TZI)
    'Debug.Print "----------"
    'Debug.Print PTimeZoneInfo_ToStr(m_TZI)
    'Debug.Print PDynTimeZoneInfo_ToStr(m_DynTZI)
    
    If IsSummerTime Or ret = TIME_ZONE_ID_STANDARD Or ret = TIME_ZONE_ID_UNKNOWN Then Exit Sub
    MsgBox "Error trying to get time-zone-info!" & vbCrLf & ret
    
End Sub

' get accurate timer
Public Function GetTimer() As Double
    Dim f As Currency: QueryPerformanceFrequency f
    Dim n As Currency: QueryPerformanceCounter n
    GetTimer = n / f
End Function

Public Function TimeZoneInfo_ConvertTimeToUtc(ByVal this As Date) As Date
    TimeZoneInfo_ConvertTimeToUtc = SystemTime_ToDate(TzSpecificLocalTime_ToSystemTime(Date_ToSystemTime(this)))
End Function

Public Property Get TimeZoneInfo_Bias() As Long
    TimeZoneInfo_Bias = m_DynTZI.TZI.Bias
End Property
Public Property Get TimeZoneInfo_StandarBias() As Long
    TimeZoneInfo_StandarBias = m_DynTZI.TZI.StandardBias
End Property
Public Property Get TimeZoneInfo_DaylightBias() As Long
    TimeZoneInfo_DaylightBias = m_DynTZI.TZI.DaylightBias
End Property

Public Function TimeZoneInfo_ToStr() As String
    TimeZoneInfo_ToStr = PTimeZoneInfo_ToStr(m_TZI)
End Function

Private Function PTimeZoneInfo_ToStr(this As TIME_ZONE_INFORMATION) As String
    Dim S As String
    With this
        S = S & "Bias         : " & .Bias & vbCrLf
        S = S & "StandardName : " & Trim0(.StandardName) & vbCrLf
        S = S & "StandardDate : " & TimeZoneInfoSystemTime_ToDate(.StandardDate) & vbCrLf
        S = S & "StandardBias : " & .StandardBias & vbCrLf
        S = S & "DaylightName : " & Trim0(.DaylightName) & vbCrLf
        S = S & "DaylightDate : " & TimeZoneInfoSystemTime_ToDate(.DaylightDate) & vbCrLf
        S = S & "DaylightBias : " & .DaylightBias & vbCrLf
    End With
    PTimeZoneInfo_ToStr = S
End Function

Public Function DynTimeZoneInfo_ToStr() As String
    DynTimeZoneInfo_ToStr = PDynTimeZoneInfo_ToStr(m_DynTZI)
End Function

Private Function PDynTimeZoneInfo_ToStr(this As DYNAMIC_TIME_ZONE_INFORMATION) As String
    Dim S As String
    With this
        S = S & PTimeZoneInfo_ToStr(.TZI)
        S = S & "TimeZoneKeyName : " & Trim0(.TimeZoneKeyName) & vbCrLf
        S = S & "TimeZoneKeyName : " & .DynamicDaylightTimeDisabled & vbCrLf
        S = S & "IsSummerTime    : " & IsSummerTime & vbCrLf
    End With
    PDynTimeZoneInfo_ToStr = S
End Function

Function Trim0(ByVal S As String) As String
    Trim0 = Trim(S)
    If Right(Trim0, 1) = vbNullChar Then
        Trim0 = Left(Trim0, Len(Trim0) - 1)
        Trim0 = Trim0(Trim0)
    'Else
    '    Exit Function
    End If
End Function

'Public Function GetSystemUpTime() As Date
Public Function GetSystemUpTime() As String
    'Returns the timespan since the last new start of your pc
    Dim ms As Currency 'milliseconds since system new start
    QueryPerformanceCounter ms
    Dim d As Long: d = ms \ MillisecondsPerDay:     ms = ms - CCur(d) * CCur(MillisecondsPerDay)
    Dim H As Long: H = ms \ MillisecondsPerHour:    ms = ms - H * MillisecondsPerHour
    Dim m As Long: m = ms \ MillisecondsPerMinute:  ms = ms - m * MillisecondsPerMinute
    Dim S As Long: S = ms \ MillisecondsPerSecond:  ms = ms - S * MillisecondsPerSecond
    GetSystemUpTime = d & ":" & Format(H, "00") & ":" & Format(m, "00") & ":" & Format(S, "00") & "." & Format(ms, "000")
End Function

'    Dim d As Date ' empty date!
'    MsgBox FormatDateTime(d, VBA.VbDateTimeFormat.vbLongDate) & " " & FormatDateTime(d, VBA.VbDateTimeFormat.vbLongTime)
'    'Samstag, 30. Dezember 1899 00:00:00

Public Function GetPCStartTime() As Date
    Dim ms As Currency 'milliseconds since system new start
    QueryPerformanceCounter ms
    Dim d As Long: d = ms \ MillisecondsPerDay:     ms = ms - d * MillisecondsPerDay
    Dim H As Long: H = ms \ MillisecondsPerHour:    ms = ms - H * MillisecondsPerHour
    Dim m As Long: m = ms \ MillisecondsPerMinute:  ms = ms - m * MillisecondsPerMinute
    Dim S As Long: S = ms \ MillisecondsPerSecond:  ms = ms - S * MillisecondsPerSecond
    GetPCStartTime = VBA.DateTime.Now - DateSerial(1900, 1, d - 1) - TimeSerial(H, m, S)
End Function

' ############################## '    DateTimeStamp    ' ############################## '
' e.g. can be found in executable files, exe, dll
Public Function DateTimeStamp(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer) As Long
    DateTimeStamp = Date_ToDateTimeStamp(New_Date(Year, Month, Day, Hour, Minute, Second))
End Function

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

Public Function DateTimeStamp_ToUniversalTimeCoordinated(DtStamp As Long) As SYSTEMTIME
    Dim syt As SYSTEMTIME: syt = DateTimeStamp_ToSystemTime(DtStamp)
    DateTimeStamp_ToUniversalTimeCoordinated = MTime.SystemTime_ToUniversalTimeCoordinated(syt)
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

Public Function DateTimeStamp_ToStrISO8601(ByVal DtStamp As Long) As String
    DateTimeStamp_ToStrISO8601 = SystemTime_ToStrISO8601(DateTimeStamp_ToSystemTime(DtStamp))
End Function

' ############################## '        Date         ' ############################## '
Public Function New_Date(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer) As Date
    New_Date = DateSerial(Year, Month, Day) + TimeSerial(Hour, Minute, Second)
End Function

Public Property Get Date_Now() As Date
    Date_Now = VBA.DateTime.Now
End Property

Public Function Date_ToSystemTime(this As Date) As SYSTEMTIME
    With Date_ToSystemTime
        .wYear = Year(this)
        .wMonth = Month(this)
        .wDayOfWeek = Weekday(this, vbUseSystemDayOfWeek)
        .wDay = Day(this)
        .wHour = Hour(this)
        .wMinute = Minute(this)
        .wSecond = Second(this)
        '.wMilliseconds = millisecond(aDate) 'nope
    End With
End Function

Public Function Date_ToFileTime(this As Date) As FILETIME
    SystemTimeToFileTime Date_ToSystemTime(this), Date_ToFileTime
End Function

Public Function Date_ToUniversalTimeCoordinated(this As Date) As SYSTEMTIME
    'Dim dat As Date: dat = TimeZoneInfo_ConvertTimeToUtc(aDate)
    Date_ToUniversalTimeCoordinated = Date_ToSystemTime(TimeZoneInfo_ConvertTimeToUtc(this))
End Function

Public Function Date_ToUnixTime(this As Date) As Double
    Date_ToUnixTime = DateDiff("s", DateSerial(1970, 1, 1), this) - GetSummerTimeCorrector
End Function

Public Function Date_ToDosTime(this As Date) As DOSTIME
    Date_ToDosTime = FileTime_ToDosTime(Date_ToFileTime(this))
End Function

Public Function Date_ToDateTimeStamp(this As Date) As Long
    Date_ToDateTimeStamp = DateDiff("s", DateSerial(1970, 1, 2), this)
End Function

Public Function Date_ToWindowsFoundationDateTime(this As Date) As WindowsFoundationDateTime
    LSet Date_ToWindowsFoundationDateTime = Date_ToFileTime(this)
End Function

Public Function Date_ToStr(this As Date) As String
    Date_ToStr = FormatDateTime(this, VbDateTimeFormat.vbLongDate) & " - " & FormatDateTime(this, VbDateTimeFormat.vbLongTime)
End Function

Public Function GetSummerTimeCorrector() As Double
    GetSummerTimeCorrector = DateDiff("s", SystemTime_ToDate(SystemTime_Now), Now)
End Function

Public Function Date_BiasMinutesToUTC(ByVal this As Date) As Long
    Dim utc As Date: utc = MTime.TimeZoneInfo_ConvertTimeToUtc(this)
    Date_BiasMinutesToUTC = DateDiff("n", utc, this)
End Function

Public Function Date_Equals(this As Date, other As Date) As Boolean
    Date_Equals = this = other
End Function

Public Function Date_ToHex(ByVal this As Date) As String
    Dim th As THexBytes, td As THexDat: td.Value = this: LSet th = td
    Date_ToHex = THexBytes_ToStr(th)
End Function

Public Function Date_ToHexNStr(ByVal this As Date) As String
    Date_ToHexNStr = Date_ToHex(this) & "; " & Date_ToStr(this)
End Function

Public Function Date_FormatISO8601(ByVal this As Date, Optional doFormatDate As Boolean = True, Optional doFormatTime As Boolean = True, Optional ByVal DateSeparator As String = "-", Optional ByVal TimeSeparator As String = ":") As String
    Dim fmt As String
    If doFormatDate Then
        fmt = "YYYY" & DateSeparator & "MM" & DateSeparator & "DD" & IIf(doFormatTime, "T", "")
    End If
    If doFormatTime Then
        fmt = fmt & "hh" & TimeSeparator & "mm" & TimeSeparator & "ss"
    End If
    Date_FormatISO8601 = Format(this, fmt)
End Function

Public Function Date_Format(ByVal this As Date, ByVal FormatStr As String) As String
    Dim S As String, y As Integer
    Select Case FormatStr
    Case "YYYY-Www":   S = Year(this) & "-W" & WeekOfYear(this)
   'Case "YYYY-Www":   ' 2004-07-11  -YYYY-MM-DD     -0333-07-11 ' 2004-W28    - YYYY-Www - 0333-W28
    Case "YYYYWww":    S = Year(this) & "W" & WeekOfYear(this)
   'Case "YYYYWww":    ' 2004-07-11  -YYYY-MM-DD     -0333-07-11 ' 2004W28     - YYYYWww    -0333W28
    Case "YYYY-Www-D": S = Year(this) & "-W" & WeekOfYear(this) & "-" & DayOfWeek(Year(this), Month(this), Day(this))
                       ' 2004-07-11  -YYYY-MM-DD     -0333-07-11 ' 2004-W28-7  - YYYY-Www-D -0333-W28-7
    Case "YYYYWwwD":   S = Year(this) & "W" & WeekOfYear(this) & DayOfWeek(Year(this), Month(this), Day(this))
                       ' 2004-07-11  -YYYY-MM-DD     -0333-07-11 ' 2004W287    - YYYYWwwD   -0333W287
    Case "YYYY-DDD":   S = Year(this) & "-" & DayOfYear(this)
                       ' 2004-07-11  -YYYY-MM-DD     -0333-07-11 ' 2004-193    - YYYY-DDD   -0333-193
    Case "YYYYDDD":    S = Year(this) & DayOfYear(this)
                       ' 2004-07-11  -YYYY-MM-DD     -0333-07-11 ' 2004193     - YYYYDDD    - 333193
    Case Else:         S = Format(this, FormatStr)
    End Select
    Date_Format = S
End Function

Public Function Date_ParseFromISO8601(ByVal S As String) As Date
Try: On Error GoTo Catch
    S = Trim$(S)
    Dim ye As Integer, mo As Integer, da As Integer, woy As Integer
    Dim ho As Integer, mn As Integer, se As Integer
    Dim DatTimSep As String: DatTimSep = GetDateTimeSeparator(S)
    Dim sa() As String
    If Len(DatTimSep) Then
        sa = Split(S, DatTimSep)
        Dim u As Long: u = UBound(sa)
        Dim sDate As String: If u > 0 Then sDate = sa(0)
        Dim sTime As String: If u > 0 Then sTime = sa(1)
        Dim lDate As Long: lDate = Len(sDate)
        Dim lTime As Long: lTime = Len(sTime)
        'Dim dDate As Date: dDate = MTime.d
        'Dim dTime As Date
        If lDate Then
            Dim DatSep As String: DatSep = GetDateSeparator(sDate)
            If Len(DatSep) Then
                sa = Split(sDate, DatSep)
                Dim ud As Long: ud = UBound(sa)
                If ud > 0 Then ye = CInt(sa(0))
                If ud > 0 Then mo = CInt(sa(1))
                If ud > 1 Then da = CInt(sa(2))
            Else
                If Not IsNumeric(sDate) Then Exit Function
                Select Case lDate
                Case 8: ye = CInt(Left(sDate, 4))
                        mo = CInt(Mid(sDate, 5, 2))
                        da = CInt(Mid(sDate, 7, 2))
                Case 7: ye = CInt(Left(sDate, 4))
                        Dim doy As Integer: doy = CLng(Mid(sDate, 5))
                        If doy > 367 Then Exit Function
                        Dim tmp As Date: tmp = Date_FromDayOfYear(ye, doy): Exit Function
                        mo = Month(tmp)
                        da = Day(tmp)
                Case 6: ye = CLng(Left(sDate, 2))
                        ye = ye + IIf(ye < 35, 2000, 1900)
                        mo = CLng(Mid(sDate, 3, 2))
                        da = CLng(Mid(sDate, 5, 2))
                End Select
            End If
        End If
        If lTime Then
            Dim TimSep As String: TimSep = GetTimeSeparator(sTime)
            If Len(TimSep) Then
                sa = Split(sTime, TimSep)
                Dim ut As Long: ut = UBound(sa)
                If ut > 0 Then ho = CInt(sa(0))
                If ut > 0 Then mn = CInt(sa(1))
                If ut > 1 Then se = CInt(sa(2))
            Else
                If Not IsNumeric(sTime) Then Exit Function
                If lTime > 1 Then ho = CInt(Left(sTime, 2))
                If lTime > 3 Then mn = CInt(Mid(sTime, 3, 2))
                If lTime > 5 Then se = CInt(Mid(sTime, 5, 2))
            End If
        End If
    Else
        If Not IsNumeric(S) Then
            If Str_Contains(S, "W") Then
                sa = Split(S, "W")
                ye = sa(0)
                If UBound(sa) > 0 Then woy = sa(1)
                Date_ParseFromISO8601 = Date_FromWeekOfYear(ye, woy)
            Else
                If Len(S) = 7 And Str_Contains(S, "-") Then
                    sa = Split(S, "-")
                    ye = sa(0)
                    If UBound(sa) > 0 Then woy = sa(1)
                    Date_ParseFromISO8601 = Date_FromWeekOfYear(ye, woy)
                Else
                    Date_ParseFromISO8601 = CDate(S)
                End If
            End If
            Exit Function
        Else
            If Len(S) > 3 Then ye = CInt(Left(S, 4))
            If Len(S) > 5 Then mo = CInt(Mid(S, 5, 2))
            If Len(S) > 7 Then da = CInt(Mid(S, 7, 2))
            If Len(S) > 9 Then ho = CInt(Mid(S, 9, 2))
            If Len(S) > 11 Then mn = CInt(Mid(S, 11, 2))
            If Len(S) > 13 Then se = CInt(Mid(S, 13, 2))
        End If
    End If
    Date_ParseFromISO8601 = New_Date(ye, mo, da, ho, mn, se)
Catch:
End Function

Private Function GetDateTimeSeparator(S As String) As String
    GetDateTimeSeparator = " "
    If Str_Contains(S, GetDateTimeSeparator) Then Exit Function
    GetDateTimeSeparator = "T"
    If Str_Contains(S, GetDateTimeSeparator) Then Exit Function
    GetDateTimeSeparator = "P"
    If Str_Contains(S, GetDateTimeSeparator) Then Exit Function
    GetDateTimeSeparator = ""
End Function

Private Function GetDateSeparator(S As String) As String
    GetDateSeparator = "-"
    If Str_Contains(S, GetDateSeparator) Then Exit Function
    GetDateSeparator = "."
    If Str_Contains(S, GetDateSeparator) Then Exit Function
    GetDateSeparator = "/"
    If Str_Contains(S, GetDateSeparator) Then Exit Function
    GetDateSeparator = "\"
    If Str_Contains(S, GetDateSeparator) Then Exit Function
    GetDateSeparator = ""
    Dim i As Long, char As Long
    For i = 1 To Len(S)
        char = AscW(Mid(S, i, 1))
        Select Case char
        Case 48 To 57 '0-9
        Case Else: GetDateSeparator = ChrW(char): Exit Function
        End Select
    Next
End Function

Private Function GetTimeSeparator(S As String) As String
    GetTimeSeparator = ":"
    If Str_Contains(S, GetTimeSeparator) Then Exit Function
    GetTimeSeparator = "."
    If Str_Contains(S, GetTimeSeparator) Then Exit Function
    GetTimeSeparator = "/"
    If Str_Contains(S, GetTimeSeparator) Then Exit Function
    GetTimeSeparator = "\"
    If Str_Contains(S, GetTimeSeparator) Then Exit Function
    GetTimeSeparator = "-"
    If Str_Contains(S, GetTimeSeparator) Then Exit Function
    GetTimeSeparator = ""
    Dim i As Long, char As Long
    For i = 1 To Len(S)
        char = AscW(Mid(S, i, 1))
        Select Case char
        Case 48 To 57 '0-9
        Case Else: GetTimeSeparator = ChrW(char): Exit Function
        End Select
    Next
End Function

Private Function Str_Contains(this As String, ByVal Value As String) As Boolean
    Str_Contains = InStr(1, this, Value) > 0
End Function
' ############################## '       MDate       ' ############################## '

Public Function ECalendar_ToStr(e As ECalendar) As String
    Dim S As String
    Select Case e
    Case ECalendar.GregorianCalendar: S = "GregorianCalendar"
    Case ECalendar.JulianCalendar:    S = "JulianCalendar"
    End Select
    ECalendar_ToStr = S
End Function

Public Function ECalendar_Parse(S As String) As ECalendar
    Dim e As ECalendar
    Select Case S
    Case "GregorianCalendar": e = ECalendar.GregorianCalendar
    Case "JulianCalendar":    e = ECalendar.JulianCalendar
    End Select
    ECalendar_Parse = e
End Function

Public Function CalcEasterdateGauss1800(ByVal y As Long, Optional ByVal ecal As ECalendar = ECalendar.GregorianCalendar) As Date
    Dim A As Long: A = y Mod 19 'der Mondparameter
    Dim b As Long: b = y Mod 4
    Dim c As Long: c = y Mod 7
    Dim k As Long: k = y \ 100 'die Säkularzahl
    Dim p As Long
    Dim q As Long
    Dim m As Long 'die säkulare Mondschaltung
    Dim d As Long 'der Keim für den ersten Vollmond im Frühling
    Dim n As Long
    Dim e As Long
    Dim OS As Long 'das Datum des Ostersonntags als Märzdatum
    Dim EasterMonth As Long
    
    Select Case ecal
    Case ECalendar.JulianCalendar
        m = 15
    Case ECalendar.GregorianCalendar
        p = k \ 3
        q = k \ 4
        m = (15 + k - p - q) Mod 30
    End Select
    
    d = (19 * A + m) Mod 30
    
    Select Case ecal
    Case ECalendar.JulianCalendar
        n = 6
    Case ECalendar.GregorianCalendar
        n = (4 + k - q) Mod 7
    End Select
    
    e = (2 * b + 4 * c + 6 * d + n) Mod 7
    
    OS = (22 + d + e)
    EasterMonth = 3
    If OS > 31 Then
        OS = OS - 31
        EasterMonth = 4
    End If
    Dim easter As Date: easter = OS & "." & EasterMonth & "." & y
    CalcEasterdateGauss1800 = easter
End Function

Public Function CalcEasterdateGauss1816(ByVal y As Long, Optional ByVal ecal As ECalendar = ECalendar.GregorianCalendar) As Date
    Dim A As Long: A = y Mod 19 'der Mondparameter / Gaußsche Zykluszahl
    Dim b As Long: b = y Mod 4
    Dim c As Long: c = y Mod 7
    Dim k As Long: k = y \ 100 'die Säkularzahl
    Dim p As Long
    Dim q As Long
    Dim m As Long 'die säkulare Mondschaltung
    Dim d As Long 'der Keim für den ersten Vollmond im Frühling
    Dim n As Long
    Dim e As Long
    Dim OS As Long 'das Datum des Ostersonntags als Märzdatum
    Dim EasterMonth As Long
    
    Select Case ecal
    Case ECalendar.JulianCalendar
        m = 15
    Case ECalendar.GregorianCalendar
        p = (8 * k + 13) \ 25 'hier unterschiedlich zu 1800
        q = k \ 4
        m = (15 + k - p - q) Mod 30
    End Select
    
    d = (19 * A + m) Mod 30
    
    Select Case ecal
    Case ECalendar.JulianCalendar
        n = 6
    Case ECalendar.GregorianCalendar
        n = (4 + k - q) Mod 7
    End Select
    
    e = (2 * b + 4 * c + 6 * d + n) Mod 7
    
    OS = (22 + d + e)
    
    CalcEasterdateGauss1816 = CorrectOSDay(OS, y)
End Function

'Schritt     Bedeutung   Formel
'1.  die Säkularzahl                                    K(X) = X div 100
'2.  die säkulare Mondschaltung                         M(K) = 15 + (3K + 3) div 4 - (8K + 13) div 25
'3.  die säkulare Sonnenschaltung                       S(K) = 2 - (3K + 3) div 4
'4.  den Mondparameter                                  A(X) = X mod 19
'5.  den Keim für den ersten Vollmond im Frühling       D(A,M) = (19A + M) mod 30
'6.  die kalendarische Korrekturgröße                   R(D,A) = (D + A div 11) div 29[13]
'7.  die Ostergrenze                                    OG(D,R) = 21 + D - R
'8.  den ersten Sonntag im März                         SZ(X,S) = 7 - (X + X div 4 + S) mod 7
'9.  die Entfernung des Ostersonntags von der Ostergrenze
'    (Osterentfernung in Tagen)                         OE(OG,SZ) = 7 - (OG - SZ) mod 7
'10. das Datum des Ostersonntags als Märzdatum
'    (32. März = 1. April usw.)                         OS = OG + OE
Public Function CalcEasterdateGaussCorrected1900(ByVal y As Long, Optional ByVal ecal As ECalendar = ECalendar.GregorianCalendar) As Date
    Dim A As Long: A = y Mod 19 'der Mondparameter / Gaußsche Zykluszahl
    'Dim b As Long: b = y Mod 4
    'Dim c As Long: c = y Mod 7
    Dim k As Long: k = y \ 100 'die Säkularzahl
    Dim p As Long
    Dim q As Long
    Dim m As Long 'die säkulare Mondschaltung
    Dim S As Long 'die säkulare Sonnenschaltung
    Dim d As Long 'der Keim für den ersten Vollmond im Frühling
    Dim r As Long 'die kalendarische Korrekturgröße
    Dim OG As Long 'die Ostergrenze
    Dim SZ As Long 'der erste Sonntag im März
    Dim OE As Long 'die Entfernung des Ostersonntags von der Ostergrenze (Osterentfernung in Tagen)
    Dim OS As Long 'das Datum des Ostersonntags als Märzdatum
    Dim n As Long
    Dim e As Long
    Dim EasterMonth As Long
    
    Select Case ecal
    Case ECalendar.JulianCalendar
        m = 15
        S = 0
    Case ECalendar.GregorianCalendar
        p = (8 * k + 13) \ 25 'hier unterschiedlich zu 1800
        q = (3 * k + 3) \ 4
        m = 15 + q - p
        S = 2 - q
    End Select
    
    d = (19 * A + m) Mod 30
    r = (d + A \ 11) \ 29
    OG = 21 + d - r
    SZ = 7 - (y + y \ 4 + S) Mod 7
    OE = 7 - (OG - SZ) Mod 7
    
    OS = OG + OE
    
    CalcEasterdateGaussCorrected1900 = CorrectOSDay(OS, y)
End Function

Public Function CorrectOSDay(ByVal OS_Mrz As Long, ByVal y As Long) As Date
    Dim OSDay   As Long: OSDay = OS_Mrz + 31 * (OS_Mrz > 31)
    Dim OSMonth As Long: OSMonth = 3 - (OS_Mrz > 31)
    CorrectOSDay = DateSerial(y, OSMonth, OSDay)
End Function

Public Function OsternShort(ByVal y As Long, Optional ByVal ecal As ECalendar = ECalendar.GregorianCalendar) As Date
    'code taken from CalcEasterdateGaussCorrected1900 + CorrectOSDay
    'and then shortened
    Dim m As Long 'die säkulare Mondschaltung
    Dim S As Long 'die säkulare Sonnenschaltung
    Select Case ecal
    Case ECalendar.JulianCalendar
        m = 15
        S = 0
    Case ECalendar.GregorianCalendar
        Dim k As Long: k = y \ 100  'die Säkularzahl
        Dim p As Long: p = (8 * k + 13) \ 25 'hier unterschiedlich zu 1800
        Dim q As Long: q = (3 * k + 3) \ 4
        m = 15 + q - p
        S = 2 - q
    End Select
    
    Dim A       As Long:  A = y Mod 19                   'der Mondparameter / Gaußsche Zykluszahl
    Dim d       As Long:  d = (19 * A + m) Mod 30       'der Keim für den ersten Vollmond im Frühling
    Dim r       As Long:  r = (d + A \ 11) \ 29         'die kalendarische Korrekturgröße
    Dim OG      As Long: OG = 21 + d - r                'die Ostergrenze
    Dim SZ      As Long: SZ = 7 - (y + y \ 4 + S) Mod 7 'der erste Sonntag im März
    Dim OE      As Long: OE = 7 - (OG - SZ) Mod 7       'die Entfernung des Ostersonntags von der Ostergrenze (Osterentfernung in Tagen)
    Dim OS      As Long: OS = OG + OE                   'das Datum des Ostersonntags als Märzdatum
    Dim OS_Mrz  As Long: OS_Mrz = OS
    Dim OSDay   As Long: OSDay = OS_Mrz + 31 * (OS_Mrz > 31)
    Dim OSMonth As Long: OSMonth = 3 - (OS_Mrz > 31)
    OsternShort = DateSerial(y, OSMonth, OSDay)
End Function

Public Function OsternShort2(ByVal y As Long) As Date
    'let's say we only want to have GregorianCalendar
    'code taken from CalcEasterdateGaussCorrected1900 and CorrectOSDay and then shortened it
    Dim k  As Long:  k = y \ 100                                            'die Säkularzahl
                                                                            '(8 * k + 13) \ 25 'hier unterschiedlich zu 1800
    Dim q  As Long:  q = (3 * k + 3) \ 4
                                                                            '2 - q '= die säkulare Sonnenschaltung
    Dim A  As Long:  A = y Mod 19                                           'der Mondparameter / Gaußsche Zykluszahl
                                                                                      '15 + q - ((8 * k + 13) \ 25) '= die säkulare Mondschaltung
    Dim d  As Long:  d = (19 * A + (15 + q - ((8 * k + 13) \ 25))) Mod 30   'der Keim für den ersten Vollmond im Frühling
                                                                                      '(d + a \ 11) \ 29 'die kalendarische Korrekturgröße
    Dim OG As Long: OG = 21 + d - (d + A \ 11) \ 29                         'die Ostergrenze
                                                                                      '7 - (y + y \ 4 + (2 - q)) Mod 7  'der erste Sonntag im März
    Dim OE As Long: OE = 7 - (OG - (7 - (y + y \ 4 + (2 - q)) Mod 7)) Mod 7 'die Entfernung des Ostersonntags von der Ostergrenze (Osterentfernung in Tagen)
    Dim OS As Long: OS = OG + OE                                            'das Datum des Ostersonntags als Märzdatum
          OsternShort2 = DateSerial(y, (3 - (OS > 31)), (OS + 31 * (OS > 31)))
End Function

Public Function AdventSunday1(ByVal Year As Integer) As Date
    Dim Nov26 As Date: Nov26 = DateSerial(Year, 11, 26)
    Dim wd As VbDayOfWeek: wd = Weekday(Nov26, VbDayOfWeek.vbMonday)
    AdventSunday1 = Nov26 + 7 - wd
End Function

Public Function Mothersday(ByVal Year As Integer) As Date
    Dim May1 As Date: May1 = DateSerial(Year, 5, 1)
    Mothersday = May1 + 15 - Weekday(May1)
End Function

Public Function Date_FromDayOfYear(ByVal Year As Integer, ByVal DayOfYear As Integer) As Date
    Dim mds As Integer
    
    mds = 31
    If DayOfYear <= mds Then Date_FromDayOfYear = DateSerial(Year, 1, DayOfYear): Exit Function
    DayOfYear = DayOfYear - mds
    
    mds = 28 - CInt(IsLeapYear(Year))
    If DayOfYear <= mds Then Date_FromDayOfYear = DateSerial(Year, 2, DayOfYear): Exit Function
    DayOfYear = DayOfYear - mds
    
    mds = 31
    If DayOfYear <= mds Then Date_FromDayOfYear = DateSerial(Year, 3, DayOfYear): Exit Function
    DayOfYear = DayOfYear - mds
    
    mds = 30
    If DayOfYear <= mds Then Date_FromDayOfYear = DateSerial(Year, 4, DayOfYear): Exit Function
    DayOfYear = DayOfYear - mds
    
    mds = 31
    If DayOfYear <= mds Then Date_FromDayOfYear = DateSerial(Year, 5, DayOfYear): Exit Function
    DayOfYear = DayOfYear - mds
    
    mds = 30
    If DayOfYear <= mds Then Date_FromDayOfYear = DateSerial(Year, 6, DayOfYear): Exit Function
    DayOfYear = DayOfYear - mds
    
    mds = 31
    If DayOfYear <= mds Then Date_FromDayOfYear = DateSerial(Year, 7, DayOfYear): Exit Function
    DayOfYear = DayOfYear - mds
    
    mds = 31
    If DayOfYear <= mds Then Date_FromDayOfYear = DateSerial(Year, 8, DayOfYear): Exit Function
    DayOfYear = DayOfYear - mds
    
    mds = 30
    If DayOfYear <= mds Then Date_FromDayOfYear = DateSerial(Year, 9, DayOfYear): Exit Function
    DayOfYear = DayOfYear - mds
    
    mds = 31
    If DayOfYear <= mds Then Date_FromDayOfYear = DateSerial(Year, 10, DayOfYear): Exit Function
    DayOfYear = DayOfYear - mds
    
    mds = 30
    If DayOfYear <= mds Then Date_FromDayOfYear = DateSerial(Year, 11, DayOfYear): Exit Function
    DayOfYear = DayOfYear - mds
    
    mds = 31
    If DayOfYear <= mds Then Date_FromDayOfYear = DateSerial(Year, 12, DayOfYear): Exit Function
End Function

'Public Function Date_FromDayOfYear(ByVal Year As Integer, ByVal doy As Integer) As Date
'    Dim m As Integer, nd As Long, nd1 As Long
'    For m = 0 To 11
'        nd1 = nd + DaysInMonth(Year, m + 1)
'        If (nd < doy) And (doy <= nd1) Then Exit For
'        nd = nd1
'    Next
'    Dim d As Integer: d = doy - nd
'    Date_FromDayOfYear = DateSerial(Year, m + 1, d)
'End Function

Public Function Date_FromWeekOfYear(ByVal Year As Integer, ByVal woy As Integer) As Date
    Date_FromWeekOfYear = Date_FromDayOfYear(Year, 7 * woy)
End Function

Public Function Date_TryParse(ByVal S As String, ByRef out_date As Date) As Boolean
Try: On Error GoTo Catch
    If LCase(S) = "now" Or LCase(S) = "jetzt" Then out_date = Now: Exit Function
    out_date = CDate(S)
    Date_TryParse = True
    Exit Function
Catch:
    MsgBox Err.Number & " " & Err.Description
End Function

Public Function Time_TryParse(ByVal S As String, out_time As Date) As Boolean
Try: On Error GoTo Catch
    If Len(S) = 0 Then Exit Function
    If LCase(S) = "now" Or LCase(S) = "jetzt" Then S = Now
    Dim sa() As String: sa = Split(S, ":")
    Dim u As Long: u = UBound(sa)
    Dim hh As String: hh = sa(0)
    If u > 0 Then
        Dim mm As String: mm = sa(1)
        If u > 1 Then
            Dim ss As String: ss = sa(2)
            Dim hhh As Integer: hhh = CInt(hh)
            Dim mmm As Integer: mmm = CInt(mm)
            Dim sss As Integer: sss = CInt(ss)
            out_time = TimeSerial(hhh, mmm, sss)
            Time_TryParse = True
            Exit Function
        End If
    End If
    'out_date = CDate(s)
    Time_TryParse = True
    Exit Function
Catch:
    MsgBox Err.Number & " " & Err.Description
End Function

Public Function Date_JulianDay(ByVal dt As Date) As Double
    Dim dat As Date: dat = DateSerial(Year(dt), Month(dt), Day(dt))
    Dim tim As Date: tim = TimeSerial(Hour(dt), Minute(dt), Second(dt))
    Dim UtcOffset As Double: UtcOffset = Date_BiasMinutesToUTC(dt) / 60
    Date_JulianDay = dat + 2415018.5 + tim - UtcOffset / 24
End Function

Public Function Date_JulianCentury(ByVal dt As Date) As Double
    Dim jd As Double: jd = Date_JulianDay(dt)
    Date_JulianCentury = (jd - 2451545#) / 36525#
End Function

'Many thanks to idiv alias Chris for the following function
'unsigned int GetDayOfWeek(unsigned int Year, unsigned int Month, unsigned int Day)
'{
'    unsigned int y, c;
'
'    if ((Month > 0) && (Month <= 12))
'    {
'        if ((Day > 0) && (Day <= GetMonthDayCount(Year, Month)))
'        {
'            y = (Year % 100);
'            c = Year / 100;
'
'            if (Month > 2)
'                Month -= 2;
'            Else
'            {
'                Month += 10;
'                if (y > 0)
'                    y--;
'                Else
'                {
'                    y = 99;
'                    c--;
'                }
'            }
'
'            return static_cast<unsigned>(((static_cast<signed>(Day) + (26*static_cast<signed>(Month)-2) / 10 + static_cast<signed>(y) + static_cast<signed>(y)/4 + static_cast<signed>(c)/4 - 2*static_cast<signed>(c)) + 7000) % 7);
'        }
'    }
'
'    return 0;
'}
Public Function DayOfWeek(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long) As Long
    
    DayOfWeek = -1
    If (Month < 1) Or (12 < Month) Then Exit Function
    
    If (Day < 1) Or (DaysInMonth(Year, Month) < Day) Then Exit Function
    
    Dim y As Long: y = Year Mod 100
    Dim c As Long: c = Year \ 100
    
    If (Month > 2) Then
        Month = Month - 2
    Else
        Month = Month + 10
        If (y > 0) Then
            y = y - 1
        Else
            y = 99
            c = c - 1
        End If
    End If
    
    'return static_cast<unsigned>(((static_cast<signed>(Day) + (26*static_cast<signed>(Month)-2) / 10 + static_cast<signed>(y) + static_cast<signed>(y)/4 + static_cast<signed>(c)/4 - 2*static_cast<signed>(c)) + 7000) % 7);
    DayOfWeek = ((Day + (26 * Month - 2) \ 10 + y + y \ 4 + c \ 4 - 2 * c) + 7000) Mod 7
    
End Function

Public Function DayOfYear(ByVal d As Date) As Long
    Dim y As Long
    Dim i As Long
    y = Year(d)
    For i = 1 To Month(d) - 1
        DayOfYear = DayOfYear + DaysInMonth(y, i)
    Next
    DayOfYear = DayOfYear + Day(d) 'Day(d)=DayOfMonth
End Function

'https://www.aktuelle-kalenderwoche.org/
'
'Was ist die Kalenderwoche?
'Die Kalenderwoche ist eine Woche, die mit dem Montag beginnt und mit dem Sonntag endet. Das Jahr hat insgesamt meist 52 Kalenderwochen.
'
'Wann beginnt die erste Kalenderwoche eines Jahres?
'Die erste Kalenderwoche eines Jahres ist die Woche, die mindestens vier Tage des neuen Jahres beinhaltet.
'Fällt also beispielsweise der 1. Januar auf einen Dienstag, beginnt die erste Kalenderwoche mit Montag, den 31.12.,
'da diese Woche sechs Tage des neuen Jahres enthält (Dienstag, Mittwoch, Donnerstag, Freitag, Samstag und Sonntag).
'Fällt der 1. Januar hingegen auf den Freitag, dann beginnt die erste Kalenderwoche des neuen Jahres mit Montag, dem 04.01.,
'da die Vorwoche nur drei Tage des neuen Jahres enthält (Freitag, Samstag, Sonntag).
'
'Gibt es für die Kalenderwoche eine offizielle Regelung?
'Ja, die ISO 8601, die als ersten Wochentag den Montag festlegt.
'
'Gilt die Regelung für die Kalenderwoche international?
'Ja und nein. In vielen Ländern wird diese Vorgehensweise verwendet. In zahlreichen anderen Ländern, wie beispielsweise den USA, Kanada,
'Mexiko und Australien, beginnt die Woche jedoch mit dem Sonntag. Ausserdem beginnt dort die erste Kalenderwoche immer mit dem 01. Januar,
'egal auf welchen Wochentag dieser fällt. Im Mittleren Osten beginnt die Woche jedoch überwiegend mit dem Samstag. Ausserdem beginnt auch
'dort die erste Kalenderwoche immer mit dem 01. Januar, egal auf welchen Wochentag dieser fällt.
'
'Gibt es eine internationale Schreibweise für die Kalenderwochen?
'Ja, auch die gibt es nach ISO 8601. Demnach wird die Kalenderwoche wie folgt geschrieben:
'YYYY-Www oder YYYYWww
'YYYY-Www-D oder YYYYWwwD
'Demnach würde in Deutschland Montag, der 31.12.2012 wie folgt bezeichnet: 2013-W01-1
'In den USA hingegen wäre es 2013-W01-2
'Und im nahen Osten 2013-W01-3.

'Zwei Beispiele für die erste KW gemäß ISO 8601
'Fällt der 1. Januar auf einen Mittwoch, beginnt die erste Kalenderwoche mit Montag, dem 30. Dezember.
'Diese 1. Woche enthält dann zwei Tage des alten Jahres (Montag und Dienstag) und fünf Tage des neuen
'Jahres (Mittwoch bis Sonntag).
'Fällt der 1. Januar hingegen auf einen Freitag, beginnt die erste Kalenderwoche mit Montag, dem 4. Januar.
'Denn die Vorwoche hätte nur drei Tage des neuen Jahres enthalten (Freitag, Samstag, Sonntag)
'
'Jahr    KW                  Anz KW   Anz Tage Schaltjahr
'2021    Kalenderwochen 2021   52       365
'2022    Kalenderwochen 2022   52       365
'2023    Kalenderwochen 2023   52       365
'2024    Kalenderwochen 2024   52       366         Ja
'2025    Kalenderwochen 2025   52       365
'2026    Kalenderwochen 2026   53       365

'https://www.vbarchiv.net/tipps/tipp_995-kalenderwoche-korrekt-ermitteln.html

Public Function WeekOfYear(ByVal d As Date) As Integer
'OK wir möchten eine Zahl erreichen die glatt durch 7 teilbar ist, und die Zahl der Kalenderwoche ergibt
    Dim y   As Integer:   y = Year(d)
    Dim wd0 As Integer: wd0 = Weekday(DateSerial(y, 1, 1), vbMonday)
    Dim wd1 As Integer: wd1 = Weekday(d, vbMonday)
    Dim doy As Integer: doy = DayOfYear(d)
    WeekOfYear = (doy + wd0 + 6 - wd1) / 7
End Function

'https://www.vbarchiv.net/tipps/tipp_995-kalenderwoche-korrekt-ermitteln.html

' *** Ermittelt die Kalenderwoche mit dem zur Kalenderwoche gehörigem Jahr
' *** fConvKW wird in der Form YYYYWW zurückgegeben ***
'Fazit:
'Solange es sich bei dem Datum nicht um die letzte Woche eines Jahres handelt, gibt die Format-Funktion die korrekten Kalenderwoche zurück.
'Nachfolgender Tipp errechnet die Kalenderwoche auch dann korrekt, wenn es sich um die letzte Woche eines Kalenderjahres handelt.
'Das Problem wird umgangen, indem man das Datum des Dienstags in dieser Woche ermittle (also losgelöst vom Montag und vom Jahresletzten) und daran die Datepart-Funktion ausführt. Dann klappt's und das Ergebnis muss ggf. nur noch auf die jeweilige Wochendarstellung zurückgerechnet werden.
Public Function WeekOfYearISO(ByVal d As Date) As Integer
    ' PROBLEM:
    ' DatePart - in der uns üblichen Arbeitsweise
    ' DatePart("ww", Datum, vbMonday, vbFirstFourDays) -
    ' wirft für folgende Daten folgende Werte aus
    ' So, 28.12.2003 oder 30.12.2007 -> KW52 -> richtig
    ' Mo, 29.12.2003 oder 31.12.2007 -> KW53 -> QUATSCH
    ' Di, 30.12.2003 oder 01.01.2008 -> KW01 -> richtig
    ' deshalb bestimme ich vorsichtshalber die Kalenderwoche
    ' des Dienstags ?!?!?!?!?!
    If Weekday(d) = vbMonday Then d = d + 1
    WeekOfYearISO = DatePart("ww", d, VbDayOfWeek.vbMonday, VbFirstWeekOfYear.vbFirstFourDays)
    ' Anpassung des Jahres und Ergänzung der Null
    ' (in den Kalenderwochen 01, 52 und 53 kann das Jahr des Datums
    ' anders sein als das Jahr der zugehörigen Kalenderwoche.
    ' So liegt der 31.12.2001 in KW01/2002)
End Function

Public Function DaysInMonth(ByVal Year As Integer, ByVal Month As Integer) As Integer
    Select Case Month
    Case 1, 3, 5, 7, 8, 10, 12: DaysInMonth = 31
    Case 2: If IsLeapYear(Year) Then DaysInMonth = 29 Else DaysInMonth = 28
    Case 4, 6, 9, 11: DaysInMonth = 30
    End Select
End Function

Public Function Weekday_ToStr(ByVal dow As VbDayOfWeek, Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbMonday, Optional ByVal isShort As Boolean = False) As String
    If FirstDayOfWeek = vbMonday Then dow = dow + IIf(dow = 7, -6, 1)
    Dim S As String
    Select Case dow
    Case VbDayOfWeek.vbSunday:    S = "Sonntag"    ' 1
    Case VbDayOfWeek.vbMonday:    S = "Montag"     ' 2
    Case VbDayOfWeek.vbTuesday:   S = "Dienstag"   ' 3
    Case VbDayOfWeek.vbWednesday: S = "Mittwoch"   ' 4
    Case VbDayOfWeek.vbThursday:  S = "Donnerstag" ' 5
    Case VbDayOfWeek.vbFriday:    S = "Freitag"    ' 6
    Case VbDayOfWeek.vbSaturday:  S = "Samstag"    ' 7
    End Select
    If isShort Then S = Left(S, 2)
    Weekday_ToStr = S
End Function

Public Function IsLeapYear(ByVal y As Long) As Boolean
'Schaltjahr (LeapYear)
'a leap year is a year which is
'either (i.)
'    evenly divisible
'        by 4
'    and not
'        by 100
'or (ii.)
'    evenly divisible
'        by 400
    IsLeapYear = (((y Mod 4) = 0) And Not ((y Mod 100) = 0)) Or ((y Mod 400) = 0)
End Function

'https://de.wikipedia.org/wiki/ISO_8601

'https://docs.microsoft.com/de-de/dotnet/standard/base-types/standard-date-and-time-format-strings
'
'Formatbezeichner    Beschreibung    Beispiele
'"d"     Kurzes Datumsmuster.
'
'Weitere Informationen finden Sie unter Der Formatbezeichner für das kurze Datum („d“).  2009-06-15T13:45:30 -> 6/15/2009 (en-US)
'
'2009-06-15T13:45:30 -> 15/06/2009 (fr-FR)
'
'2009-06-15T13:45:30 -> 2009/06/15 (ja-JP)
'"D"     Langes Datumsmuster.
'
'Weitere Informationen finden Sie unter Der Formatbezeichner für das lange Datum („D“).  2009-06-15T13:45:30 -> Monday, June 15, 2009 (en-US)
'
'2009-06-15T13:45:30 -> 15 ???? 2009 ?. (ru-RU)
'
'2009-06-15T13:45:30 -> Montag, 15. Juni 2009 (de-DE)
'"f"     Vollständiges Datums-/Zeitmuster (kurze Zeit).
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für vollständiges Datum und kurze Zeit („f“).  2009-06-15T13:45:30 -> Monday, June 15, 2009 1:45 PM (en-US)
'
'2009-06-15T13:45:30 -> den 15 juni 2009 13:45 (sv-SE)
'
'2009-06-15T13:45:30 -> ?e?t??a, 15 ??????? 2009 1:45 µµ (el-GR)
'"F"     Vollständiges Datums-/Zeitmuster (lange Zeit).
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für vollständiges Datum und lange Zeit („F“).  2009-06-15T13:45:30 -> Monday, June 15, 2009 1:45:30 PM (en-US)
'
'2009-06-15T13:45:30 -> den 15 juni 2009 13:45:30 (sv-SE)
'
'2009-06-15T13:45:30 -> ?e?t??a, 15 ??????? 2009 1:45:30 µµ (el-GR)
'"g"     Allgemeines Datums-/Zeitmuster (kurze Zeit).
'
'Weitere Informationen finden Sie unter: Der allgemeine Formatbezeichner für Datum und kurze Zeit („g“).     2009-06-15T13:45:30 -> 6/15/2009 1:45 PM (en-US)
'
'2009-06-15T13:45:30 -> 15/06/2009 13:45 (es-ES)
'
'2009-06-15T13:45:30 -> 2009/6/15 13:45 (zh-CN)
'"G"     Allgemeines Datums-/Zeitmuster (lange Zeit).
'
'Weitere Informationen finden Sie unter: Der allgemeine Formatbezeichner für Datum und lange Zeit („G“).     2009-06-15T13:45:30 -> 6/15/2009 1:45:30 PM (en-US)
'
'2009-06-15T13:45:30 -> 15/06/2009 13:45:30 (es-ES)
'
'2009-06-15T13:45:30 -> 2009/6/15 13:45:30 (zh-CN)
'"M", "m"    Monatstagmuster.
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für den Monat („M“, „m“).  2009-06-15T13:45:30 -> June 15 (en-US)
'
'2009-06-15T13:45:30 -> 15. juni (da-DK)
'
'2009-06-15T13:45:30 -> 15 Juni (id-ID)
'"O", "o"    Datums-/Uhrzeitmuster für Roundtrip.
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für Roundtrips („O“, „o“).     DateTime-Werte sind:
'
'2009-06-15T13:45:30 (DateTimeKind.Local) --> 2009-06-15T13:45:30.0000000-07:00
'
'2009-06-15T13:45:30 (DateTimeKind.Utc) --> 2009-06-15T13:45:30.0000000Z
'
'2009-06-15T13:45:30 (DateTimeKind.Unspecified) --> 2009-06-15T13:45:30.0000000
'
'DateTimeOffset:
'
'2009-06-15T13:45:30-07:00 --> 2009-06-15T13:45:30.0000000-07:00
'"R", "r"    RFC1123-Muster.
'
'Weitere Informationen finden Sie unter: Der RFC1123-Formatbezeichner („R“, „r“).    2009-06-15T13:45:30 -> Mon, 15 Jun 2009 20:45:30 GMT
'"s"     Sortierbares Datums-/Zeitmuster.
'
'Weitere Informationen finden Sie unter: Der sortierbare Formatbezeichner („s“).     2009-06-15T13:45:30 (DateTimeKind.Local) -> 2009-06-15T13:45:30
'
'2009-06-15T13:45:30 (DateTimeKind.Utc) -> 2009-06-15T13:45:30
'"t"     Kurzes Zeitmuster.
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für kurze Zeit („t“).  2009-06-15T13:45:30 -> 1:45 PM (en-US)
'
'2009-06-15T13:45:30 -> 13:45 (hr-HR)
'
'2009-06-15T13:45:30 -> 01:45 ? (ar-EG)
'"T"     Langes Zeitmuster.
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für lange Zeit („T“).  2009-06-15T13:45:30 -> 1:45:30 PM (en-US)
'
'2009-06-15T13:45:30 -> 13:45:30 (hr-HR)
'
'2009-06-15T13:45:30 -> 01:45:30 ? (ar-EG)
'"u"     Universelles, sortierbares Datums-/Zeitmuster.
'
'Weitere Informationen finden Sie unter: Der universelle sortierbare Formatbezeichner („u“).     Mit einem DateTime-Wert: 2009-06-15T13:45:30 -> 2009-06-15 13:45:30Z
'
'Mit einem DateTimeOffset-Wert: 2009-06-15T13:45:30 -> 2009-06-15 20:45:30Z
'"U"     Universelles Datums-/Zeitmuster (Koordinierte Weltzeit).
'
'Weitere Informationen finden Sie unter: Der universelle vollständige Formatbezeichner („U“).    2009-06-15T13:45:30 -> Monday, June 15, 2009 8:45:30 PM (en-US)
'
'2009-06-15T13:45:30 -> den 15 juni 2009 20:45:30 (sv-SE)
'
'2009-06-15T13:45:30 -> ?e?t??a, 15 ??????? 2009 8:45:30 µµ (el-GR)
'"Y", "y"    Jahr-Monat-Muster.
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für Jahr-Monat („Y“).  2009-06-15T13:45:30 -> Juni 2009 (en-US)
'
'2009-06-15T13:45:30 -> juni 2009 (da-DK)
'
'2009-06-15T13:45:30 -> Juni 2009 (id-ID)
'Jedes andere einzelne Zeichen   Unbekannter Bezeichner.     Löst eine FormatException zur Laufzeit aus.


' ############################## '     SystemTime      ' ############################## '
'SystemTime here is timezonespecific local time, for Systemtime as UTC-time look next chapter “Universal Time Coordinated”
Public Function SYSTEMTIME(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer, ByVal Millisecond As Integer) As SYSTEMTIME
    With SYSTEMTIME
        .wYear = Year
        .wMonth = Month
        .wDay = Day
        .wHour = Hour
        .wMinute = Minute
        .wSecond = Second
        .wMilliseconds = Millisecond
    End With
End Function

Public Property Get SystemTime_SecondsAsSingle(this As SYSTEMTIME) As Single
    With this
        SystemTime_SecondsAsSingle = CSng(.wSecond) + CSng(.wMilliseconds) / 1000!
    End With
End Property

Public Property Let SystemTime_SecondsAsSingle(this As SYSTEMTIME, ByVal Value As Single)
    With this
        .wSecond = Int(Value)
        .wMilliseconds = CInt((Value - CSng(.wSecond)) * 1000!)
    End With
End Property

Public Property Get SystemTime_Now() As SYSTEMTIME
    GetSystemTime SystemTime_Now
    SystemTime_Now = SystemTime_ToTzSpecificLocalTime(SystemTime_Now)
End Property

Public Function SystemTime_ToTzSpecificLocalTime(this As SYSTEMTIME) As SYSTEMTIME
    'UTC to local time
    SystemTimeToTzSpecificLocalTime m_DynTZI.TZI, this, SystemTime_ToTzSpecificLocalTime
End Function
Public Function TzSpecificLocalTime_ToSystemTime(this As SYSTEMTIME) As SYSTEMTIME
    'local time to UTC
    TzSpecificLocalTimeToSystemTime m_DynTZI.TZI, this, TzSpecificLocalTime_ToSystemTime
End Function

Public Function SystemTime_ToDate(this As SYSTEMTIME) As Date
    With this
        If .wYear = 0 And .wMonth <> 0 Then
            SystemTime_ToDate = TimeZoneInfoSystemTime_ToDate(this)
        End If
        SystemTime_ToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Public Function SystemTime_ToUniversalTimeCoordinated(this As SYSTEMTIME) As SYSTEMTIME
    SystemTime_ToUniversalTimeCoordinated = TzSpecificLocalTime_ToSystemTime(this)
End Function

Public Function SystemTime_ToFileTime(this As SYSTEMTIME) As FILETIME
    SystemTimeToFileTime this, SystemTime_ToFileTime
End Function

Public Function SystemTime_ToUnixTime(this As SYSTEMTIME) As Double
    SystemTime_ToUnixTime = Date_ToUnixTime(SystemTime_ToDate(this))
End Function

Public Function SystemTime_ToDosTime(this As SYSTEMTIME) As DOSTIME
    SystemTime_ToDosTime = Date_ToDosTime(SystemTime_ToDate(this))
End Function

Public Function SystemTime_ToWindowsFoundationDateTime(this As SYSTEMTIME) As WindowsFoundationDateTime
    LSet SystemTime_ToWindowsFoundationDateTime = SystemTime_ToFileTime(this)
End Function

Public Function SystemTime_ToDateTimeStamp(this As SYSTEMTIME) As Long
    SystemTime_ToDateTimeStamp = Date_ToDateTimeStamp(SystemTime_ToDate(this))
End Function

Public Function SystemTime_ToStr(this As SYSTEMTIME) As String
    With this
        SystemTime_ToStr = "y: " & CStr(.wYear) & "; m: " & CStr(.wMonth) & "; d: " & CStr(.wDay) & "; dow: " & CStr(.wDayOfWeek) & _
                         "; h: " & CStr(.wHour) & "; min: " & CStr(.wMinute) & "; s: " & CStr(.wSecond) & "; ms: " & CStr(.wMilliseconds)
    End With
End Function

Public Function SystemTime_Equals(this As SYSTEMTIME, other As SYSTEMTIME) As Boolean
    Dim b As Boolean
    With this
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

Public Function SystemTime_ToHex(this As SYSTEMTIME) As String
    Dim th As THexBytes: LSet th = this
    SystemTime_ToHex = THexBytes_ToStr(th)
End Function

Public Function SystemTime_ToStrISO8601(this As SYSTEMTIME) As String
    With this
        SystemTime_ToStrISO8601 = .wYear & "-" & .wMonth & "-" & .wDay & "T" & .wHour & ":" & .wMinute & ":" & .wSecond
    End With
End Function

Public Function SystemTime_ToHexNStr(this As SYSTEMTIME) As String
    SystemTime_ToHexNStr = SystemTime_ToHex(this) & "; " & SystemTime_ToStr(this)
End Function

' ############################## '      Coordinated Universal Time       ' ############################## '
'SystemTime here is coordinated universal time
Public Function UniversalTimeCoordinated(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer, ByVal Millisecond As Integer) As SYSTEMTIME
    UniversalTimeCoordinated = SYSTEMTIME(Year, Month, Day, Hour, Minute, Second, Millisecond)
End Function

Public Function UniversalTimeCoordinated_Now() As SYSTEMTIME
    GetSystemTime UniversalTimeCoordinated_Now
End Function

Public Function UniversalTimeCoordinated_ToDate(this As SYSTEMTIME) As Date
    Dim syt As SYSTEMTIME: syt = UniversalTimeCoordinated_ToSystemTime(this)
    UniversalTimeCoordinated_ToDate = SystemTime_ToDate(syt)
End Function

Public Function UniversalTimeCoordinated_ToSystemTime(this As SYSTEMTIME) As SYSTEMTIME
    UniversalTimeCoordinated_ToSystemTime = MTime.SystemTime_ToTzSpecificLocalTime(this)
End Function

Public Function UniversalTimeCoordinated_ToFileTime(this As SYSTEMTIME) As FILETIME
    Dim syt As SYSTEMTIME: syt = UniversalTimeCoordinated_ToSystemTime(this)
    UniversalTimeCoordinated_ToFileTime = SystemTime_ToFileTime(syt)
End Function

Public Function UniversalTimeCoordinated_ToUnixTime(this As SYSTEMTIME) As Double
    Dim syt As SYSTEMTIME: syt = UniversalTimeCoordinated_ToSystemTime(this)
    UniversalTimeCoordinated_ToUnixTime = SystemTime_ToUnixTime(syt)
End Function

Public Function UniversalTimeCoordinated_ToDOSTime(this As SYSTEMTIME) As DOSTIME
    Dim syt As SYSTEMTIME: syt = UniversalTimeCoordinated_ToSystemTime(this)
    UniversalTimeCoordinated_ToDOSTime = SystemTime_ToDosTime(syt)
End Function
    
Public Function UniversalTimeCoordinated_ToWindowsFoundationDateTime(this As SYSTEMTIME) As WindowsFoundationDateTime
    Dim syt As SYSTEMTIME: syt = UniversalTimeCoordinated_ToSystemTime(this)
    UniversalTimeCoordinated_ToWindowsFoundationDateTime = MTime.SystemTime_ToWindowsFoundationDateTime(syt)
End Function
    
Public Function UniversalTimeCoordinated_ToDateTimeStamp(this As SYSTEMTIME) As Long
    Dim syt As SYSTEMTIME: syt = UniversalTimeCoordinated_ToSystemTime(this)
    UniversalTimeCoordinated_ToDateTimeStamp = MTime.SystemTime_ToDateTimeStamp(syt)
End Function

Public Function UniversalTimeCoordinated_ToStr(this As SYSTEMTIME) As String
    UniversalTimeCoordinated_ToStr = SystemTime_ToStr(this)
End Function

Public Function UniversalTimeCoordinated_Equals(this As SYSTEMTIME, other As SYSTEMTIME) As Boolean
    UniversalTimeCoordinated_Equals = SystemTime_Equals(this, other)
End Function

Public Function UniversalTimeCoordinated_ToHex(this As SYSTEMTIME) As String
    Dim th As THexBytes: LSet th = this
    UniversalTimeCoordinated_ToHex = THexBytes_ToStr(th)
End Function

Public Function UniversalTimeCoordinated_ToStrISO8601(this As SYSTEMTIME) As String
    With this
        UniversalTimeCoordinated_ToStrISO8601 = .wYear & "-" & .wMonth & "-" & .wDay & "T" & .wHour & ":" & .wMinute & ":" & .wSecond
    End With
End Function

Public Function UniversalTimeCoordinated_ToHexNStr(this As SYSTEMTIME) As String
    UniversalTimeCoordinated_ToHexNStr = SystemTime_ToHexNStr(this)
End Function

' ############################## '      TimeZoneInfoSystemTime       ' ############################## '
Public Function TimeZoneInfoSystemTime_ToDate(this As SYSTEMTIME) As Date
    
    'the structure Time_Zone_Information has 2 Systemtime structures StandardDate and DaylightDate
    'the date is not straight, it is indeed a rule
    With this
        
        If .wYear <> 0 Then
            TimeZoneInfoSystemTime_ToDate = SystemTime_ToDate(this)
            Exit Function
        End If
        
        If .wMonth = 0 Then
            'If the time zone does not support daylight saving time or if the caller needs to disable
            'daylight saving time, the wMonth member in the SYSTEMTIME structure must be zero
            Exit Function
        End If
        
        Dim y As Integer: y = Year(Now)
        Dim m As Integer: m = .wMonth
        Dim d As Integer: d = 1
        'the date of the first day in month m
        Dim dt As Date: dt = DateSerial(y, m, d)
        
        'the weekday of the first day in month m
        Dim dow As Integer: dow = Weekday(dt) - 1 'the vbenum is one based
        d = d + DaysUntilWeekday(dow, .wDayOfWeek)
        
        'the wDay member to indicate the occurrence of the day of the week within the month
        '(1 to 5, where 5 indicates the final occurrence during the month if that day of the week does not occur 5 times)
        
        Dim idm As Integer: idm = MTime.DaysInMonth(y, .wMonth) - 7
        Dim i As Long
        For i = 1 To .wDay
            If d <= idm Then d = d + 7
        Next
        
        TimeZoneInfoSystemTime_ToDate = DateSerial(y, m, d) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Public Function DaysUntilWeekday(ByVal wd0 As Integer, ByVal wd1 As Integer, Optional FirstDayOfWeek As VbDayOfWeek = VbDayOfWeek.vbUseSystemDayOfWeek) As Integer
    'returns the number of days between different weekdays
    'e.g. from thursday to tuesday there are 5 days
    'if FirstdayOfWeek is vbUseSystemDayOfWeek the function assumes minday = 0 and maxday = 6 like in WinAPI or in .net
    'if FirstdayOfWeek is vbSunday the function assumes minday = 1 and maxday = 7 like in VB6
    If wd0 = wd1 Then Exit Function
    
    Dim Min As Integer, Max As Integer: Max = 6
    If FirstDayOfWeek = VbDayOfWeek.vbSunday Then
        Min = 1
        Max = 7
    End If
    If wd0 < Min Or Max < wd0 Then Exit Function
    If wd1 < Min Or Max < wd1 Then Exit Function
        
    DaysUntilWeekday = wd1 - wd0
    
    If wd0 < wd1 Then Exit Function
        
    DaysUntilWeekday = 7 + DaysUntilWeekday
    
End Function

'?DaysUntilWeekday(vbSunday, vbSunday,    FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 0
'?DaysUntilWeekday(vbSunday, vbMonday,    FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 1
'?DaysUntilWeekday(vbSunday, vbFriday,    FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 5
'?DaysUntilWeekday(vbSunday, vbSaturday,  FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 6
'
'?DaysUntilWeekday(vbMonday, vbSunday,    FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 6
'?DaysUntilWeekday(vbMonday, vbMonday,    FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 0
'?DaysUntilWeekday(vbMonday, vbTuesday,   FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 1
'?DaysUntilWeekday(vbMonday, vbFriday,    FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 4
'?DaysUntilWeekday(vbMonday, vbSaturday,  FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 5
'
'?DaysUntilWeekday(vbTuesday, vbSunday,    FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 5
'?DaysUntilWeekday(vbTuesday, vbMonday,    FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 6
'?DaysUntilWeekday(vbTuesday, vbTuesday,   FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 0
'?DaysUntilWeekday(vbTuesday, vbWednesday, FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 1
'?DaysUntilWeekday(vbTuesday, vbSaturday,  FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 4
'
'?DaysUntilWeekday(vbWednesday, vbSunday,    FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 4
'?DaysUntilWeekday(vbWednesday, vbTuesday,   FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 6
'?DaysUntilWeekday(vbWednesday, vbWednesday, FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 0
'?DaysUntilWeekday(vbWednesday, vbThursday,  FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 1
'?DaysUntilWeekday(vbWednesday, vbSaturday,  FirstdayOfWeek:=VbDayOfWeek.vbSunday)    ' 3

'?DaysUntilWeekday(0, 0)    ' 0
'?DaysUntilWeekday(0, 1)    ' 1
'?DaysUntilWeekday(0, 5)    ' 5
'?DaysUntilWeekday(0, 6)    ' 6

'?DaysUntilWeekday(1, 0)    ' 0
'?DaysUntilWeekday(1, 1)    ' 1
'?DaysUntilWeekday(1, 5)    ' 5
'?DaysUntilWeekday(1, 6)    ' 6


'?DaysUntilWeekday(6, 1)    ' 2

' ############################## '      FileTime       ' ############################## '
Public Function FILETIME(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer, ByVal Millisecond As Integer) As FILETIME
    FILETIME = SystemTime_ToFileTime(SYSTEMTIME(Year, Month, Day, Hour, Minute, Second, Millisecond))
End Function

Public Property Get FileTime_Now() As FILETIME
    'FileTime_Now = SystemTime_ToFileTime(SystemTime_Now)
    GetSystemTimePreciseAsFileTime FileTime_Now
    FileTime_Now = FileTime_ToLocalFileTime(FileTime_Now)
End Property

Public Function FileTime_ToLocalFileTime(this As FILETIME) As FILETIME
'    FileTimeToLocalFileTime aFt, FileTime_ToLocalFileTime
    Dim st_in As SYSTEMTIME: st_in = FileTime_ToSystemTime(this)
    Dim stout As SYSTEMTIME
    SystemTimeToTzSpecificLocalTime m_DynTZI.TZI, st_in, stout
    FileTime_ToLocalFileTime = SystemTime_ToFileTime(stout)
End Function

Public Function LocalFileTime_ToFileTime(this As FILETIME) As FILETIME
'    LocalFileTimeToFileTime aFt, LocalFileTime_ToFileTime
    Dim st_in As SYSTEMTIME: st_in = FileTime_ToSystemTime(this)
    Dim stout As SYSTEMTIME
    TzSpecificLocalTimeToSystemTime m_DynTZI.TZI, st_in, stout
    LocalFileTime_ToFileTime = SystemTime_ToFileTime(stout)
End Function

Public Property Get FileTime_ToDosTime(this As FILETIME) As DOSTIME
    Dim dt As DOSTIME
    FileTimeToDosDateTime this, dt.wDate, dt.wTime
    FileTime_ToDosTime = dt
End Property

Public Function FileTime_ToDate(this As FILETIME) As Date
    With FileTime_ToSystemTime(this) 'st
        FileTime_ToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Public Function FileTime_ToSystemTime(this As FILETIME) As SYSTEMTIME
    FileTimeToSystemTime this, FileTime_ToSystemTime
End Function

Public Function FileTime_ToUniversalTimeCoordinated(this As FILETIME) As SYSTEMTIME
    Dim syt As SYSTEMTIME: syt = MTime.FileTime_ToSystemTime(this)
    FileTime_ToUniversalTimeCoordinated = MTime.SystemTime_ToUniversalTimeCoordinated(syt)
End Function

Public Function FileTime_ToUnixTime(this As FILETIME) As Double
    FileTime_ToUnixTime = Date_ToUnixTime(FileTime_ToDate(this))
End Function

Public Function FileTime_ToWindowsFoundationDateTime(this As FILETIME) As WindowsFoundationDateTime
    LSet FileTime_ToWindowsFoundationDateTime = this
End Function

Public Function FileTime_ToDateTimeStamp(this As FILETIME) As Long
    FileTime_ToDateTimeStamp = Date_ToDateTimeStamp(FileTime_ToDate(this))
End Function

Public Function FileTime_ToStr(this As FILETIME) As String
    With this
        FileTime_ToStr = "lo: " & CStr(.dwLowDateTime) & "; hi: " & CStr(.dwHighDateTime)
    End With
End Function

Public Function FileTime_Equals(this As FILETIME, other As FILETIME) As Boolean
    Dim b As Boolean
    With this
        b = .dwHighDateTime = other.dwHighDateTime: If Not b Then Exit Function
        b = .dwLowDateTime = other.dwLowDateTime ':   If Not b Then Exit Function
    End With
    FileTime_Equals = b
End Function

Public Function FileTime_ToHex(this As FILETIME) As String
    Dim th As THexBytes: LSet th = this
    FileTime_ToHex = THexBytes_ToStr(th)
End Function

Public Function FileTime_ToStrISO8601(this As FILETIME) As String
    FileTime_ToStrISO8601 = SystemTime_ToStrISO8601(FileTime_ToSystemTime(this))
End Function

Public Function FileTime_ToHexNStr(this As FILETIME) As String
    FileTime_ToHexNStr = FileTime_ToHex(this) & "; " & FileTime_ToStr(this)
End Function

' ############################## '      UnixTime       ' ############################## '
' In Unix und Linux werden Datumsangaben intern immer als die Anzahl der Sekunden seit
' dem 1. Januar 1970 um 00:00 Greenwhich Mean Time (GMT, heute UTC) dargestellt.
' Dieses Urdatum wird manchmal auch "The Epoch" genannt. In manchen Situationen muss
' man in Shellskripten die Unix-Zeit in ein normales Datum umrechnen und umgekehrt.
Public Function UnixTime(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer) As Double
    UnixTime = Date_ToUnixTime(New_Date(Year, Month, Day, Hour, Minute, Second))
End Function

Public Property Get UnixTime_Now() As Double
    UnixTime_Now = Date_ToUnixTime(Date_Now)
End Property

Public Function UnixTime_ToDate(ByVal uts As Double) As Date
    UnixTime_ToDate = DateAdd("s", uts + GetSummerTimeCorrector, DateSerial(1970, 1, 1))
End Function

Public Function UnixTime_ToSystemTime(ByVal uts As Double) As SYSTEMTIME
    UnixTime_ToSystemTime = Date_ToSystemTime(UnixTime_ToDate(uts))
End Function

Public Function UnixTime_ToUniversalTimeCoordinated(this As Double) As SYSTEMTIME
    Dim syt As SYSTEMTIME: syt = MTime.UnixTime_ToSystemTime(this)
    UnixTime_ToUniversalTimeCoordinated = MTime.SystemTime_ToUniversalTimeCoordinated(syt)
End Function

Public Function UnixTime_ToFileTime(ByVal uts As Double) As FILETIME
    UnixTime_ToFileTime = Date_ToFileTime(UnixTime_ToDate(uts))
End Function

Public Function UnixTime_ToDosTime(ByVal uts As Double) As DOSTIME
    UnixTime_ToDosTime = Date_ToDosTime(UnixTime_ToDate(uts))
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

Public Function UnixTime_Equals(ByVal this As Double, ByVal other As Double) As Boolean
    UnixTime_Equals = this = other
End Function

Public Function UnixTime_ToHex(ByVal this As Double) As String
    Dim th As THexBytes, td As THexDbl: td.Value = this: LSet th = td
    UnixTime_ToHex = THexBytes_ToStr(th)
End Function

Public Function UnixTime_ToStrISO8601(ByVal this As Double) As String
    UnixTime_ToStrISO8601 = SystemTime_ToStrISO8601(UnixTime_ToSystemTime(this))
End Function

Public Function UnixTime_ToHexNStr(ByVal uts As Double) As String
    UnixTime_ToHexNStr = UnixTime_ToHex(uts) & "; " & UnixTime_ToStr(uts)
End Function

' ############################## '       DosTime       ' ############################## '
' oder auch FAT-Time also die Zeit die unter DOS in der FAT der Festplatte gespeichert wird
Public Function DOSTIME(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer) As DOSTIME
    DOSTIME = Date_ToDosTime(New_Date(Year, Month, Day, Hour, Minute, Second))
End Function

Public Function DosTime_Now() As DOSTIME
    DosTime_Now = FileTime_ToDosTime(Date_ToFileTime(Date_Now))
End Function

Public Function DosTime_ToDate(this As DOSTIME) As Date
    DosTime_ToDate = FileTime_ToDate(DosTime_ToFileTime(this))
End Function

Public Function DosTime_ToSystemTime(this As DOSTIME) As SYSTEMTIME
    DosTime_ToSystemTime = FileTime_ToSystemTime(DosTime_ToFileTime(this))
End Function

Public Function DosTime_ToUniversalTimeCoordinated(this As DOSTIME) As SYSTEMTIME
    Dim syt As SYSTEMTIME: syt = MTime.DosTime_ToSystemTime(this)
    DosTime_ToUniversalTimeCoordinated = SystemTime_ToUniversalTimeCoordinated(syt)
End Function

Public Property Get DosTime_ToFileTime(this As DOSTIME) As FILETIME
    DosDateTimeToFileTime this.wDate, this.wTime, DosTime_ToFileTime
End Property

Public Function DosTime_ToUnixTime(this As DOSTIME) As Double
    DosTime_ToUnixTime = Date_ToUnixTime(DosTime_ToDate(this))
End Function

Public Function DosTime_ToWindowsFoundationDateTime(this As DOSTIME) As WindowsFoundationDateTime
    LSet DosTime_ToWindowsFoundationDateTime = DosTime_ToFileTime(this)
End Function

Public Function DosTime_ToDateTimeStamp(this As DOSTIME) As Long
    DosTime_ToDateTimeStamp = Date_ToDateTimeStamp(DosTime_ToDate(this))
End Function

Public Function DosTime_ToStr(this As DOSTIME) As String
    ' Bits    Description
    ' 0 - 4   Day of the month (1–31)
    ' 5 - 8   Month (1 = January, 2 = February, and so on)
    ' 9 -15   Year offset from 1980 (add 1980 to get actual year)
    'wDate As Integer
    ' Bits Description
    ' 0 - 4   Second divided by 2
    ' 5 -10   Minute (0–59)
    '11 -15   Hour (0–23 on a 24-hour clock)
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
    DosTime_ToStr = "wDate: " & CStr(this.wDate) & "; wTime: " & CStr(this.wTime)
End Function

'Private Function Str2(ByVal by As Byte) As String
'    Str2 = CStr(by): If Len(Str2) < 2 Then Str2 = "0" & Str2
'End Function
Public Function DosTime_Equals(this As DOSTIME, other As DOSTIME) As Boolean
    Dim b As Boolean
    With this
        b = .wDate = other.wDate: If Not b Then Exit Function
        b = .wTime = other.wTime ': If Not b Then Exit Function
    End With
    DosTime_Equals = b
End Function

Public Function DosTime_ToHex(this As DOSTIME) As String
    Dim th As THexBytes: LSet th = this
    DosTime_ToHex = THexBytes_ToStr(th)
End Function

Public Function DosTime_ToStrISO8601(this As DOSTIME) As String
    DosTime_ToStrISO8601 = SystemTime_ToStrISO8601(DosTime_ToSystemTime(this))
End Function

Public Function DosTime_ToHexNStr(this As DOSTIME) As String
    DosTime_ToHexNStr = DosTime_ToHex(this) & "; " & DosTime_ToStr(this)
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
'    'hmm müßte eigentlich sein Date_ToCyTime
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
Public Function WindowsFoundationDateTime(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer) As WindowsFoundationDateTime
    WindowsFoundationDateTime = Date_ToWindowsFoundationDateTime(New_Date(Year, Month, Day, Hour, Minute, Second))
End Function

Public Function WindowsFoundationDateTime_Now() As WindowsFoundationDateTime
    LSet WindowsFoundationDateTime_Now = FileTime_Now
End Function

Public Function WindowsFoundationDateTime_ToFileTime(this As WindowsFoundationDateTime) As FILETIME
    LSet WindowsFoundationDateTime_ToFileTime = this
End Function

Public Function WindowsFoundationDateTime_ToSystemTime(this As WindowsFoundationDateTime) As SYSTEMTIME
    WindowsFoundationDateTime_ToSystemTime = FileTime_ToSystemTime(WindowsFoundationDateTime_ToFileTime(this))
End Function

Public Function WindowsFoundationDateTime_ToUniversalTimeCoordinated(this As WindowsFoundationDateTime) As SYSTEMTIME
    WindowsFoundationDateTime_ToUniversalTimeCoordinated = WindowsFoundationDateTime_ToSystemTime(this)
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

Public Function WindowsFoundationDateTime_ToStrISO8601(this As WindowsFoundationDateTime) As String
    WindowsFoundationDateTime_ToStrISO8601 = SystemTime_ToStrISO8601(WindowsFoundationDateTime_ToSystemTime(this))
End Function

' ############################## '       THexBytes       ' ############################## '
Private Function THexBytes_ToStr(this As THexBytes) As String
    Dim S As String: S = "&&H"
    Dim i As Long, u As Long: u = UBound(this.Value)
    For i = u To 0 Step -1
        S = S & Hex2(this.Value(i))
    Next
    THexBytes_ToStr = S
End Function
Private Function Hex2(ByVal b As Byte) As String
    Hex2 = Hex(b): If Len(Hex2) < 2 Then Hex2 = "0" & Hex2
End Function
