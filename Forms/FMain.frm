VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Datetime-Conversions"
   ClientHeight    =   7590
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows-Standard
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
      Height          =   3615
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

Private Sub Form_Load()
    MTime.Init
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Btn_Date_Now_Click
    Btn_DosTime_Now_Click
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
    
    Dim dat As Date: dat = MTime.Date_Now
    LblDateNow.Caption = MTime.Date_ToHexNStr(dat)
    LblSystemTimeNow.Caption = MTime.SystemTime_ToHexNStr(MTime.Date_ToSystemTime(dat))
    LblUTCTimeNow.Caption = MTime.SystemTime_ToHexNStr(Date_ToUniversalTimeCoordinated(dat))
    
    LblFileTimeNow.Caption = MTime.FileTime_ToHexNStr(MTime.Date_ToFileTime(dat))
    LblUnixTimeNow.Caption = MTime.UnixTime_ToHexNStr(MTime.Date_ToUnixTime(dat))
    LblDosTimeNow.Caption = MTime.DosTime_ToHexNStr(MTime.Date_ToDosTime(dat))
    LblWinRTTimeNow.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(MTime.Date_ToWindowsFoundationDateTime(dat))
    LblDTStampNow.Caption = MTime.DateTimeStamp_ToHexNStr(MTime.Date_ToDateTimeStamp(dat))
    
End Sub

Private Sub Btn_SystemTime_Now_Click()
    
    Dim syt As SYSTEMTIME: syt = MTime.SystemTime_Now
    LblDateNow.Caption = MTime.Date_ToHexNStr(MTime.SystemTime_ToDate(syt))
    LblSystemTimeNow.Caption = MTime.SystemTime_ToHexNStr(syt)
    LblUTCTimeNow.Caption = MTime.SystemTime_ToHexNStr(MTime.UniversalTimeCoordinated_Now)
    
    LblFileTimeNow.Caption = MTime.FileTime_ToHexNStr(MTime.SystemTime_ToFileTime(syt))
    LblUnixTimeNow.Caption = MTime.UnixTime_ToHexNStr(MTime.SystemTime_ToUnixTime(syt))
    LblDosTimeNow.Caption = MTime.DosTime_ToHexNStr(MTime.SystemTime_ToDosTime(syt))
    LblWinRTTimeNow.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(MTime.SystemTime_ToWindowsFoundationDateTime(syt))
    LblDTStampNow.Caption = MTime.DateTimeStamp_ToHexNStr(MTime.SystemTime_ToDateTimeStamp(syt))
    
End Sub

Private Sub Btn_UTCTime_Now_Click()
    
    Dim utc As SYSTEMTIME: utc = MTime.UniversalTimeCoordinated_Now
    Dim dat As Date: dat = MTime.UniversalTimeCoordinated_ToDate(utc)
    LblDateNow.Caption = MTime.Date_ToHexNStr(dat)
    LblSystemTimeNow.Caption = MTime.SystemTime_ToHexNStr(MTime.Date_ToSystemTime(dat))
    LblUTCTimeNow.Caption = MTime.SystemTime_ToHexNStr(utc)
    
    LblFileTimeNow.Caption = MTime.FileTime_ToHexNStr(MTime.Date_ToFileTime(dat))
    LblUnixTimeNow.Caption = MTime.UnixTime_ToHexNStr(MTime.Date_ToUnixTime(dat))
    LblDosTimeNow.Caption = MTime.DosTime_ToHexNStr(MTime.Date_ToDosTime(dat))
    LblWinRTTimeNow.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(MTime.Date_ToWindowsFoundationDateTime(dat))
    LblDTStampNow.Caption = MTime.DateTimeStamp_ToHexNStr(MTime.Date_ToDateTimeStamp(dat))
    
End Sub

Private Sub Btn_FileTime_Now_Click()
    
    Dim fit As FILETIME: fit = MTime.FileTime_Now
    LblDateNow.Caption = MTime.Date_ToHexNStr(MTime.FileTime_ToDate(fit))
    LblSystemTimeNow.Caption = MTime.SystemTime_ToHexNStr(MTime.FileTime_ToSystemTime(fit))
    LblUTCTimeNow.Caption = MTime.SystemTime_ToHexNStr(MTime.UniversalTimeCoordinated_Now)
    
    LblFileTimeNow.Caption = MTime.FileTime_ToHexNStr(fit)
    LblUnixTimeNow.Caption = MTime.UnixTime_ToHexNStr(MTime.FileTime_ToUnixTime(fit))
    LblDosTimeNow.Caption = MTime.DosTime_ToHexNStr(MTime.FileTime_ToDosTime(fit))
    LblWinRTTimeNow.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(MTime.FileTime_ToWindowsFoundationDateTime(fit))
    LblDTStampNow.Caption = MTime.DateTimeStamp_ToHexNStr(MTime.FileTime_ToDateTimeStamp(fit))
    
End Sub

Private Sub Btn_UnixTime_Now_Click()
    
    Dim uxt As Currency: uxt = MTime.UnixTime_Now
    LblDateNow.Caption = MTime.Date_ToHexNStr(MTime.UnixTime_ToDate(uxt))
    LblSystemTimeNow.Caption = MTime.SystemTime_ToHexNStr(MTime.UnixTime_ToSystemTime(uxt))
    LblUTCTimeNow.Caption = MTime.SystemTime_ToHexNStr(MTime.UniversalTimeCoordinated_Now)
    
    LblFileTimeNow.Caption = MTime.FileTime_ToHexNStr(MTime.UnixTime_ToFileTime(uxt))
    LblUnixTimeNow.Caption = MTime.UnixTime_ToHexNStr(uxt)
    LblDosTimeNow.Caption = MTime.DosTime_ToHexNStr(MTime.UnixTime_ToDosTime(uxt))
    LblWinRTTimeNow.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(MTime.UnixTime_ToWindowsFoundationDateTime(uxt))
    LblDTStampNow.Caption = MTime.DateTimeStamp_ToHexNStr(MTime.UnixTime_ToDateTimeStamp(uxt))
    
End Sub

Private Sub Btn_DosTime_Now_Click()
    
    Dim Dst As DOSTIME: Dst = MTime.DosTime_Now
    LblDateNow.Caption = MTime.Date_ToHexNStr(MTime.DosTime_ToDate(Dst))
    LblSystemTimeNow.Caption = MTime.SystemTime_ToHexNStr(MTime.DosTime_ToSystemTime(Dst))
    LblUTCTimeNow.Caption = MTime.SystemTime_ToHexNStr(MTime.UniversalTimeCoordinated_Now)
    
    LblFileTimeNow.Caption = MTime.FileTime_ToHexNStr(MTime.DosTime_ToFileTime(Dst))
    LblUnixTimeNow.Caption = MTime.UnixTime_ToHexNStr(MTime.DosTime_ToUnixTime(Dst))
    LblDosTimeNow.Caption = MTime.DosTime_ToHexNStr(Dst)
    LblWinRTTimeNow.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(MTime.DosTime_ToWindowsFoundationDateTime(Dst))
    LblDTStampNow.Caption = MTime.DateTimeStamp_ToHexNStr(MTime.DosTime_ToDateTimeStamp(Dst))
    
End Sub

Private Sub Btn_WinFndDateTime_Now_Click()
    
    Dim wfdt As WindowsFoundationDateTime: wfdt = MTime.WindowsFoundationDateTime_Now
    LblDateNow.Caption = MTime.Date_ToHexNStr(MTime.WindowsFoundationDateTime_ToDate(wfdt))
    LblSystemTimeNow.Caption = MTime.SystemTime_ToHexNStr(MTime.WindowsFoundationDateTime_ToSystemTime(wfdt))
    LblUTCTimeNow.Caption = MTime.SystemTime_ToHexNStr(MTime.UniversalTimeCoordinated_Now)
    
    LblFileTimeNow.Caption = MTime.FileTime_ToHexNStr(MTime.WindowsFoundationDateTime_ToFileTime(wfdt))
    LblUnixTimeNow.Caption = MTime.UnixTime_ToHexNStr(MTime.WindowsFoundationDateTime_ToUnixTime(wfdt))
    LblDosTimeNow.Caption = MTime.DosTime_ToHexNStr(MTime.WindowsFoundationDateTime_ToDosTime(wfdt))
    LblWinRTTimeNow.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(wfdt)
    LblDTStampNow.Caption = MTime.DateTimeStamp_ToHexNStr(MTime.WindowsFoundationDateTime_ToDateTimeStamp(wfdt))
    
End Sub

Private Sub Btn_DateTimeStamp_Now_Click()
    
    Dim dts As Long: dts = MTime.DateTimeStamp_Now
    LblDateNow.Caption = MTime.Date_ToHexNStr(MTime.DateTimeStamp_ToDate(dts))
    LblSystemTimeNow.Caption = MTime.SystemTime_ToHexNStr(MTime.DateTimeStamp_ToSystemTime(dts))
    LblUTCTimeNow.Caption = MTime.SystemTime_ToHexNStr(MTime.UniversalTimeCoordinated_Now)
    
    LblFileTimeNow.Caption = MTime.FileTime_ToHexNStr(MTime.DateTimeStamp_ToFileTime(dts))
    LblUnixTimeNow.Caption = MTime.UnixTime_ToHexNStr(MTime.DateTimeStamp_ToUnixTime(dts))
    LblDosTimeNow.Caption = MTime.DosTime_ToHexNStr(MTime.DateTimeStamp_ToDosTime(dts))
    LblWinRTTimeNow.Caption = MTime.WindowsFoundationDateTime_ToHexNStr(MTime.DateTimeStamp_ToWindowsFoundationDateTime(dts))
    LblDTStampNow.Caption = MTime.DateTimeStamp_ToHexNStr(dts)
    
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
    Text1.Text = s
End Sub

Private Sub Command10_Click()
    'Label8.Caption = MTime.GetSystemUpTime ' GetPCStartTime
    Text1.Text = MTime.GetSystemUpTime
End Sub

Private Sub Command11_Click()
    'Label9.Caption = MTime.GetPCStartTime
    Text1.Text = MTime.GetPCStartTime
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

