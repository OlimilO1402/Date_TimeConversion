VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Datetime-Conversions"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13950
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnEditDate 
      Caption         =   "Edit Date"
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton Btn_UTCTime_Now 
      Caption         =   "Coordin.Univers.T.(UTC)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton Command17 
      Caption         =   "DosDate{22222;22222}"
      Height          =   375
      Left            =   0
      TabIndex        =   28
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Easter sunday?"
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Is it a leap year?"
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton Command11 
      Caption         =   "GetPCStartNewDate"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   25
      ToolTipText     =   "Returns the date and time when your pc got started new"
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton Command10 
      Caption         =   "GetSystemUpTime"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   24
      ToolTipText     =   "Returns the timespan since the last new start of your pc"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Btn_DateTimeStamp_Now 
      Caption         =   "DateTimeStamp_Now"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Btn_WinFndDateTime_Now 
      Caption         =   "WinFndDateTime_Now"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton Btn_IsSummerTime 
      Caption         =   "IsSummerTime"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton Btn_DosTime_Now 
      Caption         =   "DosTime_Now"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   9
      Top             =   3840
      Width           =   11655
   End
   Begin VB.CommandButton BtnSomeUnixTimeTests 
      Caption         =   "Some UnixTime tests"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton Btn_UnixTime_Now 
      Caption         =   "UnixTime_Now"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Btn_FileTime_Now 
      Caption         =   "FileTime_Now"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Btn_SystemTime_Now 
      Caption         =   "SystemTime_Now"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Btn_Date_Now 
      Caption         =   "Date_Now"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label LblDTStampNow 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3720
      TabIndex        =   31
      Top             =   3480
      Width           =   630
   End
   Begin VB.Label LblUTCTime 
      AutoSize        =   -1  'True
      Caption         =   "UTC-Time:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   30
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label LblWinRTTimeNow 
      AutoSize        =   -1  'True
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3720
      TabIndex        =   23
      Top             =   3000
      Width           =   630
   End
   Begin VB.Label LblDTStamp 
      AutoSize        =   -1  'True
      Caption         =   "DateTimeStamp:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   22
      Top             =   3480
      Width           =   1245
   End
   Begin VB.Label LblWinRTTime 
      AutoSize        =   -1  'True
      Caption         =   "WinRt.DateTime:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   20
      Top             =   3000
      Width           =   1290
   End
   Begin VB.Label LblDosTimeNow 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3720
      TabIndex        =   19
      Top             =   2520
      Width           =   630
   End
   Begin VB.Label LblUnixTimeNow 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3720
      TabIndex        =   16
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label LblDosTime 
      AutoSize        =   -1  'True
      Caption         =   "DosTime:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   15
      Top             =   2520
      Width           =   705
   End
   Begin VB.Label LblUnixTime 
      AutoSize        =   -1  'True
      Caption         =   "UnixTime:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   13
      Top             =   2040
      Width           =   750
   End
   Begin VB.Label LblFileTime 
      AutoSize        =   -1  'True
      Caption         =   "FileTime:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   12
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label LblSystemTime 
      AutoSize        =   -1  'True
      Caption         =   "SystemTime:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   11
      Top             =   600
      Width           =   930
   End
   Begin VB.Label LblDate 
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   405
   End
   Begin VB.Label LblDateNow 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   630
   End
   Begin VB.Label LblFileTimeNow 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3720
      TabIndex        =   3
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label LblUTCTimeNow 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3720
      TabIndex        =   2
      Top             =   1080
      Width           =   630
   End
   Begin VB.Label LblSystemTimeNow 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   630
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_DateTime   As Date
Private m_SystemTime As MTime.SYSTEMTIME
Private m_UTCTime    As MTime.SYSTEMTIME
Private m_FileTime   As MTime.FILETIME
Private m_UnixTime   As Double
Private m_DOSTime    As MTime.DOSTIME
Private m_WndFndDTim As MTime.WindowsFoundationDateTime
Private m_DTimeStamp As Long

Private Sub BtnEditDate_Click()
    Dim dat As SYSTEMTIME: dat = m_SystemTime
    If FEditDateTime.ShowDialog(Me, dat) = vbCancel Then Exit Sub
    UpdateFromSYSTEMTIME dat
End Sub

Private Sub Form_Load()
    MTime.Init
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Btn_Date_Now_Click
    'Btn_DosTime_Now_Click
    Btn_IsSummerTime_Click
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = Text1.Left
    Dim t As Single: t = Text1.Top
    Dim W As Single: W = Me.ScaleWidth - L
    Dim h As Single: h = Me.ScaleHeight - t
    If W > 0 And h > 0 Then Text1.Move L, t, W, h
End Sub

Private Sub Btn_Date_Now_Click()
    
    UpdateFromDate MTime.Date_Now
    
End Sub

Private Sub UpdateFromDate(ByVal NewDate As Date)
    
    m_DateTime = NewDate
    
    m_SystemTime = MTime.Date_ToSystemTime(m_DateTime)
    m_UTCTime = MTime.Date_ToUniversalTimeCoordinated(m_DateTime)
    m_FileTime = MTime.Date_ToFileTime(m_DateTime)
    m_UnixTime = MTime.Date_ToUnixTime(m_DateTime)
    m_DOSTime = MTime.Date_ToDosTime(m_DateTime)
    m_WndFndDTim = MTime.Date_ToWindowsFoundationDateTime(m_DateTime)
    m_DTimeStamp = MTime.Date_ToDateTimeStamp(m_DateTime)
    
    UpdateView
End Sub

Private Sub Btn_SystemTime_Now_Click()
    
    UpdateFromSYSTEMTIME MTime.SystemTime_Now
    
End Sub

Private Sub UpdateFromSYSTEMTIME(NewSystemTime As SYSTEMTIME)
    
    m_SystemTime = NewSystemTime
    
    m_DateTime = MTime.SystemTime_ToDate(m_SystemTime)
    m_UTCTime = MTime.SystemTime_ToUniversalTimeCoordinated(m_SystemTime)
    m_FileTime = MTime.SystemTime_ToFileTime(m_SystemTime)
    m_UnixTime = MTime.SystemTime_ToUnixTime(m_SystemTime)
    m_DOSTime = MTime.SystemTime_ToDosTime(m_SystemTime)
    m_WndFndDTim = MTime.SystemTime_ToWindowsFoundationDateTime(m_SystemTime)
    m_DTimeStamp = MTime.SystemTime_ToDateTimeStamp(m_SystemTime)
    
    UpdateView
End Sub

Private Sub Btn_UTCTime_Now_Click()
        
    m_UTCTime = MTime.UniversalTimeCoordinated_Now
    
    m_DateTime = MTime.UniversalTimeCoordinated_ToDate(m_UTCTime)
    m_SystemTime = MTime.UniversalTimeCoordinated_ToSystemTime(m_UTCTime)
    m_FileTime = MTime.UniversalTimeCoordinated_ToFileTime(m_UTCTime)
    m_UnixTime = MTime.UniversalTimeCoordinated_ToUnixTime(m_UTCTime)
    m_DOSTime = MTime.UniversalTimeCoordinated_ToDOSTime(m_UTCTime)
    m_WndFndDTim = MTime.UniversalTimeCoordinated_ToWindowsFoundationDateTime(m_UTCTime)
    m_DTimeStamp = MTime.UniversalTimeCoordinated_ToDateTimeStamp(m_UTCTime)
    
    UpdateView
End Sub

Private Sub Btn_FileTime_Now_Click()
    
    m_FileTime = MTime.FileTime_Now
    
    m_DateTime = MTime.FileTime_ToDate(m_FileTime)
    m_SystemTime = MTime.FileTime_ToSystemTime(m_FileTime)
    m_UTCTime = MTime.FileTime_ToUniversalTimeCoordinated(m_FileTime)
    m_UnixTime = MTime.FileTime_ToUnixTime(m_FileTime)
    m_DOSTime = MTime.FileTime_ToDosTime(m_FileTime)
    m_WndFndDTim = MTime.FileTime_ToWindowsFoundationDateTime(m_FileTime)
    m_DTimeStamp = MTime.FileTime_ToDateTimeStamp(m_FileTime)
    
    UpdateView
End Sub

Private Sub Btn_UnixTime_Now_Click()
    
    m_UnixTime = MTime.UnixTime_Now
    
    m_DateTime = MTime.UnixTime_ToDate(m_UnixTime)
    m_SystemTime = MTime.UnixTime_ToSystemTime(m_UnixTime)
    m_UTCTime = MTime.UnixTime_ToUniversalTimeCoordinated(m_UnixTime)
    m_FileTime = MTime.UnixTime_ToFileTime(m_UnixTime)
    m_DOSTime = MTime.UnixTime_ToDosTime(m_UnixTime)
    m_WndFndDTim = MTime.UnixTime_ToWindowsFoundationDateTime(m_UnixTime)
    m_DTimeStamp = MTime.UnixTime_ToDateTimeStamp(m_UnixTime)
    
    UpdateView
End Sub

Private Sub Btn_DosTime_Now_Click()
    
    m_DOSTime = MTime.DosTime_Now
    
    m_DateTime = MTime.DosTime_ToDate(m_DOSTime)
    m_SystemTime = MTime.DosTime_ToSystemTime(m_DOSTime)
    m_UTCTime = MTime.DosTime_ToUniversalTimeCoordinated(m_DOSTime)
    m_FileTime = MTime.DosTime_ToFileTime(m_DOSTime)
    m_UnixTime = MTime.DosTime_ToUnixTime(m_DOSTime)
    m_WndFndDTim = MTime.DosTime_ToWindowsFoundationDateTime(m_DOSTime)
    m_DTimeStamp = MTime.DosTime_ToDateTimeStamp(m_DOSTime)
    
    UpdateView
End Sub

Private Sub Btn_WinFndDateTime_Now_Click()
    
    m_WndFndDTim = MTime.WindowsFoundationDateTime_Now
        
    m_DateTime = MTime.WindowsFoundationDateTime_ToDate(m_WndFndDTim)
    m_SystemTime = MTime.WindowsFoundationDateTime_ToSystemTime(m_WndFndDTim)
    m_UTCTime = MTime.WindowsFoundationDateTime_ToUniversalTimeCoordinated(m_WndFndDTim)
    m_FileTime = MTime.WindowsFoundationDateTime_ToFileTime(m_WndFndDTim)
    m_UnixTime = MTime.WindowsFoundationDateTime_ToUnixTime(m_WndFndDTim)
    m_DOSTime = MTime.WindowsFoundationDateTime_ToDosTime(m_WndFndDTim)
    m_DTimeStamp = MTime.WindowsFoundationDateTime_ToDateTimeStamp(m_WndFndDTim)
    
    UpdateView
End Sub

Private Sub Btn_DateTimeStamp_Now_Click()
    
    m_DTimeStamp = MTime.DateTimeStamp_Now
    
    m_DateTime = MTime.DateTimeStamp_ToDate(m_DTimeStamp)
    m_SystemTime = MTime.DateTimeStamp_ToSystemTime(m_DTimeStamp)
    m_UTCTime = MTime.DateTimeStamp_ToUniversalTimeCoordinated(m_DTimeStamp)
    m_FileTime = MTime.DateTimeStamp_ToFileTime(m_DTimeStamp)
    m_UnixTime = MTime.DateTimeStamp_ToUnixTime(m_DTimeStamp)
    m_DOSTime = MTime.DateTimeStamp_ToDosTime(m_DTimeStamp)
    m_WndFndDTim = MTime.DateTimeStamp_ToWindowsFoundationDateTime(m_DTimeStamp)
    
    UpdateView
End Sub

Private Sub UpdateView()
    LblDateNow.Caption = MTime.Date_ToHexNStr(m_DateTime)
    LblSystemTimeNow.Caption = MTime.SystemTime_ToHexNStr(m_SystemTime)
    LblUTCTimeNow.Caption = MTime.UniversalTimeCoordinated_ToHexNStr(m_UTCTime)
    LblFileTimeNow.Caption = MTime.FileTime_ToHexNStr(m_FileTime)
    LblUnixTimeNow.Caption = MTime.UnixTime_ToHexNStr(m_UnixTime)
    LblDosTimeNow.Caption = MTime.DosTime_ToHexNStr(m_DOSTime)
    LblWinRTTimeNow.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(m_WndFndDTim)
    LblDTStampNow.Caption = MTime.DateTimeStamp_ToHexNStr(m_DTimeStamp)
End Sub

Private Sub BtnSomeUnixTimeTests_Click()
    
    Dim dat As Date
    Dim uxs As Double
    Dim s   As String
    
    'https://alexander-fischer-online.net/fuer-webmaster/unix-timestamp-in-datumsformat-wandeln.html
    'https://checkmk.com/de/linux-wissen/datum-umrechnen
    
    'user@linux> date -d @1234567890
    'Sa 14. Feb 00:31:30 CET 2009
    Dim d As Integer, h As Integer
    If IsSummerTime Then
        d = 13: h = 23
    Else
        d = 14: h = 0
    End If
    dat = DateSerial(2009, 2, d) + TimeSerial(h, 31, 30)
    uxs = Date_ToUnixTime(dat)
    s = s & "1234567890 = " & uxs & " : " & CBool(1234567890 = uxs) & vbCrLf
    
    'user@linux > Date - d '2008-12-18 12:34:00' +%s
    '1229600040
    If IsSummerTime Then
        h = 11
    Else
        h = 12
    End If
    dat = DateSerial(2008, 12, 18) + TimeSerial(h, 34, 0)
    uxs = Date_ToUnixTime(dat)
    s = s & "1229600040 = " & uxs & " : " & CBool(1229600040 = uxs) & vbCrLf
    
    'user@ linux > Date - d '1970-01-01 00:00:00' +%s
    '-3600
    dat = DateSerial(1970, 1, 1) '+ TimeSerial(0, 0, 0)
    uxs = Date_ToUnixTime(dat)
    s = s & "-3600 = " & uxs & " : " & CBool(-3600 = uxs) & vbCrLf
    
    'convert from unixtime to date
    uxs = 1234567890
    'dat = UnixTime_ToDate(uxs)
    s = s & "The unixtimestamp: " & uxs & " stands for the date: " & Date_ToStr(UnixTime_ToDate(uxs)) & vbCrLf
    
    Text1.Text = s
    
End Sub

Private Sub Btn_IsSummerTime_Click()
    Dim s As String
    s = MTime.DynTimeZoneInfo_ToStr
    Dim dat As Date: dat = DateTime.Now
    Dim utc As Date: utc = MTime.TimeZoneInfo_ConvertTimeToUtc(dat)
    s = s & "dat: " & CStr(dat) & vbCrLf & "utc: " & CStr(utc) & vbCrLf
    Dim BiasMin As Long: BiasMin = MTime.Date_BiasMinutesToUTC(dat)
    s = s & "UtcBias (minutes): " & BiasMin & vbCrLf
    s = s & "SystemUpTime: " & MTime.GetSystemUpTime & vbCrLf
    s = s & "PCStartTime : " & MTime.GetPCStartTime & vbCrLf
    Dim y As Long: y = CLng(Int(Year(dat)))
    
    s = s & "The year " & y & " is " & IIf(Not MTime.IsLeapYear(y), "not ", "") & "a leap year." & vbCrLf
    Dim esd As Date: esd = MTime.OsternShort2(y)
    s = s & "The eastern sunday in the year " & y & " is " & FormatDateTime(esd, VBA.VbDateTimeFormat.vbShortDate) & vbCrLf

    Text1.Text = s
End Sub

Private Sub Command10_Click()
    'Label8.Caption = MTime.GetSystemUpTime ' GetPCStartTime
    Text1.Text = "SystemUpTime: " & MTime.GetSystemUpTime
End Sub

Private Sub Command11_Click()
    'Label9.Caption = MTime.GetPCStartTime
    Text1.Text = "PCStartTime: " & MTime.GetPCStartTime
End Sub

Private Sub Command12_Click()
    Dim s As String, d As Date ' empty Date
    s = s & FormatDateTime(d, VBA.VbDateTimeFormat.vbLongDate) & " " & FormatDateTime(d, VBA.VbDateTimeFormat.vbLongTime)
    d = VBA.DateTime.Now
    s = s & FormatDateTime(d, VBA.VbDateTimeFormat.vbLongDate) & " " & FormatDateTime(d, VBA.VbDateTimeFormat.vbLongTime)
    Text1.Text = s
End Sub

Private Sub Command13_Click()
    MsgBox MTime.TimeZoneInfo_ToStr
    Dim dt As Date: dt = DateTime.Now
    Dim utc As Date: utc = MTime.TimeZoneInfo_ConvertTimeToUtc(dt)
    MsgBox dt & vbCrLf & utc
End Sub

Private Sub Command14_Click()
'    For i = 1970 To 2023
'        If IsLeapYear(i) Then
'            Debug.Print i & " leap year"
'        Else
'            Debug.Print i
'        End If
'    Next
'    Exit Sub
    Dim i As Long
    Dim y As Integer, m As Integer, d As Integer
    For i = 0 To 1000
        y = 1970 + Rnd * (2023 - 1970)
        m = 1 + Rnd * 11
        d = 1 + Rnd * (DaysInMonth(y, m) - 1)
        CheckOneDate y, m, d
    Next
End Sub

Private Sub CheckOneDate(ByVal y As Integer, ByVal m As Integer, ByVal d As Integer)
    Dim dt  As Date:     dt = DateSerial(y, m, d)
    Dim wkd As Integer: wkd = Weekday(dt) - 1
    Dim dow As Integer: dow = GetDayOfWeek(y, m, d)
    If wkd <> dow Then
        Debug.Print "wkd = " & wkd & " <> " & "dow = " & dow & " " & dt
    End If
End Sub

Private Sub Command15_Click()
    Dim s As String: s = InputBox("Year:", "Calculates if the year is a leap-year", Year(Now))
    If StrPtr(s) = 0 Then Exit Sub
    If Not IsNumeric(s) Then Exit Sub
    Dim y As Long: y = CLng(Int(Val(s)))
    Text1.Text = "The year " & y & " is " & IIf(Not MTime.IsLeapYear(y), "not ", "") & "a leap year."
End Sub

Private Sub Command16_Click()
    Dim s As String: s = InputBox("Year:", "Calculates the easter sunday for the year", Year(Now))
    If StrPtr(s) = 0 Then Exit Sub
    If Not IsNumeric(s) Then Exit Sub
    Dim y As Long: y = CLng(Int(Val(s)))
    Dim dat As Date: dat = MTime.OsternShort2(y)
    Text1.Text = "The eastern sunday in the year " & y & " is " & Date_ToStr(dat)
End Sub

Private Sub Command17_Click()
    Dim dd As DOSTIME
    dd.wDate = 22222 'this was today 14.jun.2023
    dd.wTime = 22222
    MsgBox "DosTime{" & dd.wDate & "; " & dd.wTime & "} = " & vbCrLf & DosTime_ToDate(dd)
End Sub

