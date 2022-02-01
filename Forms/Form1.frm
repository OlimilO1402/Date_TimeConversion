VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "DosTime_Now"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2640
      Width           =   6255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Some UnixTime tests"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UnixTime_Now"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "FileTime_Now"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SystemTime_Now"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Date_Now"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   2040
      Width           =   5175
   End
   Begin VB.Label Labe15 
      Caption         =   "DosTime:"
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "UnixTime:"
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "FileTime:"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Labe12 
      Caption         =   "SystemTime:"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Date:"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command7_Click()
    Dim dat As Date: dat = Now
    MsgBox Date_ToStr(dat)
    Dim dtstmp As Long: dtstmp = Date_ToDateTimeStamp(dat)
    MsgBox DateTimeStamp_ToStr(dtstmp)
End Sub

Private Sub Form_Load()
    Command1_Click
    Command5_Click
End Sub

Private Sub Command1_Click()

    Dim dat As Date: dat = MTime.Date_Now
    Label1.Caption = MTime.Date_ToStr(dat)
    Label2.Caption = MTime.SystemTime_ToStr(MTime.Date_ToSystemTime(dat))
    Label3.Caption = MTime.FileTime_ToStr(MTime.Date_ToFileTime(dat))
    Label4.Caption = MTime.UnixTime_ToStr(MTime.Date_ToUnixTime(dat))
    Label5.Caption = MTime.DosTime_ToStr(MTime.Date_ToDosTime(dat))
    
End Sub

Private Sub Command2_Click()
    
    Dim syt As SYSTEMTIME: syt = MTime.SystemTime_Now
    Label1.Caption = MTime.Date_ToStr(MTime.SystemTime_ToDate(syt))
    Label2.Caption = MTime.SystemTime_ToStr(syt)
    Label3.Caption = MTime.FileTime_ToStr(MTime.SystemTime_ToFileTime(syt))
    Label4.Caption = MTime.UnixTime_ToStr(MTime.SystemTime_ToUnixTime(syt))

End Sub

Private Sub Command3_Click()
    
    Dim fit As FILETIME: fit = MTime.FileTime_Now
    Label1.Caption = MTime.Date_ToStr(MTime.FileTime_ToDate(fit))
    Label2.Caption = MTime.SystemTime_ToStr(MTime.FileTime_ToSystemTime(fit))
    Label3.Caption = MTime.FileTime_ToStr(fit)
    Label4.Caption = MTime.UnixTime_ToStr(MTime.FileTime_ToUnixTime(fit))
    
End Sub

Private Sub Command4_Click()
    
    Dim uxt As Currency: uxt = MTime.UnixTime_Now
    Label1.Caption = MTime.Date_ToStr(MTime.UnixTime_ToDate(uxt))
    Label2.Caption = MTime.SystemTime_ToStr(MTime.UnixTime_ToSystemTime(uxt))
    Label3.Caption = MTime.FileTime_ToStr(MTime.UnixTime_ToFileTime(uxt))
    Label4.Caption = MTime.UnixTime_ToStr(uxt)
    
End Sub

Private Sub Command5_Click()
    
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
