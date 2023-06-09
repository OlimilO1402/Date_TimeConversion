VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Datetime-Conversions"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13695
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   13695
   StartUpPosition =   3  'Windows-Standard
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
      TabIndex        =   27
      ToolTipText     =   "Returns the date and time when your pc got started new"
      Top             =   5160
      Width           =   1935
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
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
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
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
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
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
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
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
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
      Top             =   2040
      Width           =   1935
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
      Height          =   3615
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   9
      Top             =   3480
      Width           =   11655
   End
   Begin VB.CommandButton Command8 
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
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
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
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
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
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
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
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
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
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "PCStartNewDate..."
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
      Left            =   120
      TabIndex        =   26
      Top             =   5640
      Width           =   1395
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "SystemUpTime..."
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
      Left            =   120
      TabIndex        =   25
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label Label7 
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
      Left            =   3480
      TabIndex        =   23
      Top             =   3000
      Width           =   630
   End
   Begin VB.Label Label17 
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
      Left            =   2040
      TabIndex        =   22
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Label Label16 
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
      Left            =   2040
      TabIndex        =   20
      Top             =   2520
      Width           =   1290
   End
   Begin VB.Label Label6 
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
      Left            =   3480
      TabIndex        =   19
      Top             =   2520
      Width           =   630
   End
   Begin VB.Label Label5 
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
      Left            =   3480
      TabIndex        =   16
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label Labe15 
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
      Left            =   2040
      TabIndex        =   15
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label Label14 
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
      Left            =   2040
      TabIndex        =   13
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label Label13 
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
      Left            =   2040
      TabIndex        =   12
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label Labe12 
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
      Left            =   2040
      TabIndex        =   11
      Top             =   600
      Width           =   930
   End
   Begin VB.Label Label11 
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
      Left            =   2040
      TabIndex        =   10
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Label1 
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
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label4 
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
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label Label3 
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
      Left            =   3480
      TabIndex        =   2
      Top             =   1080
      Width           =   630
   End
   Begin VB.Label Label2 
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
      Left            =   3480
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

Private Sub Form_Load()
    MTime.Init
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Command1_Click
    Command5_Click
    Command9_Click
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = Text1.Left
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth - L
    Dim h As Single: h = Me.ScaleHeight - T
    If W > 0 And h > 0 Then Text1.Move L, T, W, h
End Sub

Private Sub Command1_Click()
    
    Dim dat As Date: dat = MTime.Date_Now
    Label1.Caption = MTime.Date_ToHexNStr(dat)
    Label2.Caption = MTime.SystemTime_ToHexNStr(MTime.Date_ToSystemTime(dat))
    Label3.Caption = MTime.FileTime_ToHexNStr(MTime.Date_ToFileTime(dat))
    Label4.Caption = MTime.UnixTime_ToHexNStr(MTime.Date_ToUnixTime(dat))
    Label5.Caption = MTime.DosTime_ToHexNStr(MTime.Date_ToDosTime(dat))
    Label6.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(MTime.Date_ToWindowsFoundationDateTime(dat))
    Label7.Caption = MTime.DateTimeStamp_ToHexNStr(MTime.Date_ToDateTimeStamp(dat))
    
End Sub

Private Sub Command2_Click()
    
    Dim syt As SYSTEMTIME: syt = MTime.SystemTime_Now
    Label1.Caption = MTime.Date_ToHexNStr(MTime.SystemTime_ToDate(syt))
    Label2.Caption = MTime.SystemTime_ToHexNStr(syt)
    Label3.Caption = MTime.FileTime_ToHexNStr(MTime.SystemTime_ToFileTime(syt))
    Label4.Caption = MTime.UnixTime_ToHexNStr(MTime.SystemTime_ToUnixTime(syt))
    Label5.Caption = MTime.DosTime_ToHexNStr(MTime.SystemTime_ToDosTime(syt))
    Label6.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(MTime.SystemTime_ToWindowsFoundationDateTime(syt))
    Label7.Caption = MTime.DateTimeStamp_ToHexNStr(MTime.SystemTime_ToDateTimeStamp(syt))
    
End Sub

Private Sub Command3_Click()
    
    Dim fit As FILETIME: fit = MTime.FileTime_Now
    Label1.Caption = MTime.Date_ToHexNStr(MTime.FileTime_ToDate(fit))
    Label2.Caption = MTime.SystemTime_ToHexNStr(MTime.FileTime_ToSystemTime(fit))
    Label3.Caption = MTime.FileTime_ToHexNStr(fit)
    Label4.Caption = MTime.UnixTime_ToHexNStr(MTime.FileTime_ToUnixTime(fit))
    Label5.Caption = MTime.DosTime_ToHexNStr(MTime.FileTime_ToDosTime(fit))
    Label6.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(MTime.FileTime_ToWindowsFoundationDateTime(fit))
    Label7.Caption = MTime.DateTimeStamp_ToHexNStr(MTime.FileTime_ToDateTimeStamp(fit))
    
End Sub

Private Sub Command4_Click()
    
    Dim uxt As Currency: uxt = MTime.UnixTime_Now
    Label1.Caption = MTime.Date_ToHexNStr(MTime.UnixTime_ToDate(uxt))
    Label2.Caption = MTime.SystemTime_ToHexNStr(MTime.UnixTime_ToSystemTime(uxt))
    Label3.Caption = MTime.FileTime_ToHexNStr(MTime.UnixTime_ToFileTime(uxt))
    Label4.Caption = MTime.UnixTime_ToHexNStr(uxt)
    Label5.Caption = MTime.DosTime_ToHexNStr(MTime.UnixTime_ToDosTime(uxt))
    Label6.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(MTime.UnixTime_ToWindowsFoundationDateTime(uxt))
    Label7.Caption = MTime.DateTimeStamp_ToHexNStr(MTime.UnixTime_ToDateTimeStamp(uxt))
    
End Sub

Private Sub Command5_Click()
    
    Dim Dst As DOSTIME: Dst = MTime.DosTime_Now
    Label1.Caption = MTime.Date_ToHexNStr(MTime.DosTime_ToDate(Dst))
    Label2.Caption = MTime.SystemTime_ToHexNStr(MTime.DosTime_ToSystemTime(Dst))
    Label3.Caption = MTime.FileTime_ToHexNStr(MTime.DosTime_ToFileTime(Dst))
    Label4.Caption = MTime.UnixTime_ToHexNStr(MTime.DosTime_ToUnixTime(Dst))
    Label5.Caption = MTime.DosTime_ToHexNStr(Dst)
    Label6.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(MTime.DosTime_ToWindowsFoundationDateTime(Dst))
    Label7.Caption = MTime.DateTimeStamp_ToHexNStr(MTime.DosTime_ToDateTimeStamp(Dst))
    
End Sub

Private Sub Command6_Click()
    
    Dim wfdt As WindowsFoundationDateTime: wfdt = MTime.WindowsFoundationDateTime_Now
    Label1.Caption = MTime.Date_ToHexNStr(MTime.WindowsFoundationDateTime_ToDate(wfdt))
    Label2.Caption = MTime.SystemTime_ToHexNStr(MTime.WindowsFoundationDateTime_ToSystemTime(wfdt))
    Label3.Caption = MTime.FileTime_ToHexNStr(MTime.WindowsFoundationDateTime_ToFileTime(wfdt))
    Label4.Caption = MTime.UnixTime_ToHexNStr(MTime.WindowsFoundationDateTime_ToUnixTime(wfdt))
    Label5.Caption = MTime.DosTime_ToHexNStr(MTime.WindowsFoundationDateTime_ToDosTime(wfdt))
    Label6.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(wfdt)
    Label7.Caption = MTime.DateTimeStamp_ToHexNStr(MTime.WindowsFoundationDateTime_ToDateTimeStamp(wfdt))
    
End Sub

Private Sub Command7_Click()
    
    Dim dts As Long: dts = MTime.DateTimeStamp_Now
    Label1.Caption = MTime.Date_ToHexNStr(MTime.DateTimeStamp_ToDate(dts))
    Label2.Caption = MTime.SystemTime_ToHexNStr(MTime.DateTimeStamp_ToSystemTime(dts))
    Label3.Caption = MTime.FileTime_ToHexNStr(MTime.DateTimeStamp_ToFileTime(dts))
    Label4.Caption = MTime.UnixTime_ToHexNStr(MTime.DateTimeStamp_ToUnixTime(dts))
    Label5.Caption = MTime.DosTime_ToHexNStr(MTime.DateTimeStamp_ToDosTime(dts))
    Label6.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(MTime.DateTimeStamp_ToWindowsFoundationDateTime(dts))
    Label7.Caption = MTime.DateTimeStamp_ToHexNStr(dts)
    
End Sub

Private Sub Command8_Click()
    
    Dim dat As Date
    Dim uxs As Double
    Dim s   As String
    
    'https://alexander-fischer-online.net/fuer-webmaster/unix-timestamp-in-datumsformat-wandeln.html
    'https://checkmk.com/de/linux-wissen/datum-umrechnen
    
    'user@linux> date -d @1234567890
    'Sa 14. Feb 00:31:30 CET 2009
    dat = DateSerial(2009, 2, 14) + TimeSerial(0, 31, 30)
    uxs = Date_ToUnixTime(dat)
    s = s & "1234567890 = " & uxs & " : " & CBool(1234567890 = uxs) & vbCrLf
    
    'user@ linux > Date - d '2008-12-18 12:34:00' +%s
    '1229600040
    dat = DateSerial(2008, 12, 18) + TimeSerial(12, 34, 0)
    uxs = Date_ToUnixTime(dat)
    s = s & "1229600040 = " & uxs & " : " & CBool(1229600040 = uxs) & vbCrLf
    
    'user@ linux > Date - d '1970-01-01 00:00:00' +%s
    '-3600
    dat = DateSerial(1970, 1, 1) '+ TimeSerial(0, 0, 0)
    uxs = Date_ToUnixTime(dat)
    s = s & "-3600 = " & uxs & " : " & CBool(-3600 = uxs) & vbCrLf
    
    'convert from unixtime to date
    uxs = 1234567890
    dat = UnixTime_ToDate(uxs)
    s = s & "The unixtimestamp: " & uxs & " stands for the date: " & Date_ToStr(dat) & vbCrLf
    
    Text1.Text = s
    
End Sub

Private Sub Command9_Click()
    Dim s As String
    s = MTime.TimeZoneInfo_ToStr
    Dim dat As Date: dat = DateTime.Now
    Dim utc As Date: utc = MTime.TimeZoneInfo_ConvertTimeToUtc(dat)
    s = s & "dat: " & CStr(dat) & vbCrLf & "utc: " & CStr(utc) & vbCrLf
    Dim BiasMin As Long: BiasMin = MTime.Date_BiasMinutesToUTC(dat)
    s = s & "UtcBias (minutes): " & BiasMin & vbCrLf
    Text1.Text = s
End Sub

Private Sub Command10_Click()
    Label8.Caption = MTime.GetSystemUpTime ' GetPCStartTime
End Sub

Private Sub Command11_Click()
    Label9.Caption = MTime.GetPCStartTime
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

