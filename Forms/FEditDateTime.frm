VERSION 5.00
Begin VB.Form FEditDateTime 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Edit Date and Time"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5055
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TBSecond 
      Alignment       =   1  'Rechts
      Height          =   330
      Left            =   3600
      TabIndex        =   11
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox TBMinute 
      Alignment       =   1  'Rechts
      Height          =   330
      Left            =   3600
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox TBHour 
      Alignment       =   1  'Rechts
      Height          =   330
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox TBDay 
      Alignment       =   1  'Rechts
      Height          =   330
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox TBMonth 
      Alignment       =   1  'Rechts
      Height          =   330
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox TBYear 
      Alignment       =   1  'Rechts
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Second :"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   840
      Width           =   855
   End
   Begin VB.Label LblMinute 
      Caption         =   "Minute :"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label LblHour 
      Caption         =   "Hour :"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LblDay 
      Caption         =   "Day :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label LblMonth 
      Caption         =   "Month :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label LblYear 
      Caption         =   "Year :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FEditDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_y As Integer
Private m_m As Integer
Private m_d As Integer
Private m_h As Integer
Private m_n As Integer
Private m_s As Integer 'Double
Private m_Result As VbMsgBoxResult
Private m_Date As Date

Public Function ShowDialog(Owner As Form, DateInOut As Date) As VbMsgBoxResult
    m_Date = DateInOut
    m_y = Year(m_Date):    m_m = Month(m_Date):    m_d = Day(m_Date)
    m_h = Hour(m_Date):    m_n = Minute(m_Date):   m_s = Second(m_Date) '+ Millisecond(m_Date) / 1000 'nope!
    UpdateView
    Me.Show vbModal, Owner
    ShowDialog = m_Result
    DateInOut = DateSerial(m_y, m_m, m_d) + TimeSerial(m_h, m_n, m_s)
End Function

Sub UpdateView()
    TBYear.Text = CStr(m_y)
    TBMonth.Text = CStr(m_m)
    TBDay.Text = CStr(m_d)
    TBHour.Text = CStr(m_h)
    TBMinute.Text = CStr(m_n)
    TBSecond.Text = CStr(m_s)
End Sub

Private Sub BtnOK_Click()
    m_Result = VbMsgBoxResult.vbOK
    Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_Result = VbMsgBoxResult.vbCancel
    Unload Me
End Sub

Private Sub TBYear_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KeyCodeConstants.vbKeyDown Then
        'how to get the next atbstopped textbox/control???
    End If
End Sub

Private Sub TBYear_LostFocus():   TBDateTime_LostFocus TBYear, "year", -1000, 5000, m_y: End Sub
Private Sub TBMonth_LostFocus():  TBDateTime_LostFocus TBMonth, "month", 1, 12, m_m:     End Sub
Private Sub TBDay_LostFocus():    TBDateTime_LostFocus TBDay, "day", 1, MTime.DaysInMonth(m_y, m_m), m_d: End Sub
Private Sub TBHour_LostFocus():   TBDateTime_LostFocus TBHour, "hour", 0, 23, m_h:       End Sub
Private Sub TBMinute_LostFocus(): TBDateTime_LostFocus TBMinute, "minute", 0, 59, m_n:   End Sub
Private Sub TBsecond_LostFocus(): TBDateTime_LostFocus TBSecond, "second", 0, 59, m_s:   End Sub

Private Sub TBDateTime_LostFocus(aTB As TextBox, ByVal PropName As String, ByVal RangeMin As Long, ByVal RangeMax As Long, ByRef Value_inout As Integer)
    Dim v As Long: v = Value_inout
    If TryParseNCheckRange(aTB.Text, PropName, RangeMin, RangeMax, v) Then Value_inout = v
    UpdateView
End Sub
'Validate
Private Function TryParseNCheckRange(ByVal StrValue As String, ByVal PropName As String, ByVal RangeMin As Long, ByVal RangeMax As Long, ByRef Value_inout As Long) As Boolean
Try: On Error GoTo Catch
    Dim v As Long: v = CLng(StrValue)
    If v < RangeMin Or RangeMax < v Then GoTo Catch
    Value_inout = v
    TryParseNCheckRange = True
    Exit Function
Catch:
    MsgBox "Please give a valid value for """ & PropName & """ in the range between (" & RangeMin & " - " & RangeMax & ")!" & vbCrLf & """" & StrValue & """ is not a valid value."
End Function
