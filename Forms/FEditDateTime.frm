VERSION 5.00
Begin VB.Form FEditDateTime 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Edit Date and Time"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4335
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
   ScaleHeight     =   1455
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TBSecond 
      Alignment       =   2  'Zentriert
      Height          =   330
      Left            =   3240
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox TBMinute 
      Alignment       =   2  'Zentriert
      Height          =   330
      Left            =   2760
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox TBHour 
      Alignment       =   2  'Zentriert
      Height          =   330
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox TBDay 
      Alignment       =   2  'Zentriert
      Height          =   330
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox TBMonth 
      Alignment       =   2  'Zentriert
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox TBYear 
      Alignment       =   2  'Zentriert
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label LblHour 
      Alignment       =   2  'Zentriert
      Caption         =   "hh:mm:ss.xxx"
      Height          =   225
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1890
   End
   Begin VB.Label LblDate 
      Alignment       =   2  'Zentriert
      Caption         =   "YYYY-MM-DD"
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1905
   End
End
Attribute VB_Name = "FEditDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private m_y As Integer
'Private m_m As Integer
'Private m_d As Integer
'Private m_h As Integer
'Private m_n As Integer
'Private m_s As Integer 'Double
'Private m_ms As Integer
Private m_Result As VbMsgBoxResult
Private m_Date As SYSTEMTIME

Friend Function ShowDialog(Owner As Form, ByRef SystemTime_InOut As SYSTEMTIME) As VbMsgBoxResult
    m_Date = SystemTime_InOut
    'With m_Date
    'm_y = .wYear:     m_m = Month(m_Date):    m_d = Day(m_Date)
    'm_h = Hour(m_Date):    m_n = Minute(m_Date):   m_s = Second(m_Date) '
    'm_ms Millisecond(m_Date) / 1000  'nope!
    UpdateView
    Me.Show vbModal, Owner
    ShowDialog = m_Result
    'DateInOut = DateSerial(m_y, m_m, m_d) + TimeSerial(m_h, m_n, m_s)
    SystemTime_InOut = m_Date ' MTime.SYSTEMTIME(m_y, m_m, m_d, m_h, m_n, m_s, m_ms)
End Function

Sub UpdateView()
    'With m_Date
    TBYear.Text = CStr(m_Date.wYear)
    TBMonth.Text = CStr(m_Date.wMonth)
    TBDay.Text = CStr(m_Date.wDay)
    TBHour.Text = CStr(m_Date.wHour)
    TBMinute.Text = CStr(m_Date.wMinute)
    TBSecond.Text = CStr(CSng(m_Date.wSecond) + CSng(CSng(m_Date.wMilliseconds) / CSng(1000)))
    'End With
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

Private Sub TBYear_LostFocus():   TB_LostFocusInt TBYear, "year", -1000, 5000, m_Date.wYear: End Sub
Private Sub TBMonth_LostFocus():  TB_LostFocusInt TBMonth, "month", 1, 12, m_Date.wMonth:     End Sub
Private Sub TBDay_LostFocus():    TB_LostFocusInt TBDay, "day", 1, MTime.DaysInMonth(m_Date.wYear, m_Date.wMonth), m_Date.wDay: End Sub
Private Sub TBHour_LostFocus():   TB_LostFocusInt TBHour, "hour", 0, 23, m_Date.wHour:       End Sub
Private Sub TBMinute_LostFocus(): TB_LostFocusInt TBMinute, "minute", 0, 59, m_Date.wMinute:   End Sub

Private Sub TBsecond_LostFocus()
    Dim seconds As Single: seconds = SystemTime_SecondsAsSingle(m_Date)
    TB_LostFocusSng TBSecond, "second", 0!, 59.999!, seconds
    SystemTime_SecondsAsSingle(m_Date) = seconds
    'm_s = Int(seconds)
    'm_ms = Int((seconds - Int(seconds)) * 1000)
End Sub

Private Sub TB_LostFocusInt(aTB As TextBox, ByVal PropName As String, ByVal RangeMin As Long, ByVal RangeMax As Long, ByRef Value_inout As Integer)
    Dim v As Integer: v = Value_inout
    If TryParseNCheckRangeInt(aTB.Text, PropName, RangeMin, RangeMax, v) Then Value_inout = v
    UpdateView
End Sub

Private Sub TB_LostFocusSng(aTB As TextBox, ByVal PropName As String, ByVal RangeMin As Single, ByVal RangeMax As Single, ByRef Value_inout As Single)
    Dim v As Single: v = Value_inout
    If TryParseNCheckRangeSng(aTB.Text, PropName, RangeMin, RangeMax, v) Then Value_inout = v
    UpdateView
End Sub

'Validate
Private Function TryParseNCheckRangeInt(ByVal StrValue As String, ByVal PropName As String, ByVal RangeMin As Long, ByVal RangeMax As Long, ByRef Value_inout As Integer) As Boolean
Try: On Error GoTo Catch
    Dim v As Integer: v = CInt(StrValue)
    If v < RangeMin Or RangeMax < v Then GoTo Catch
    Value_inout = v
    TryParseNCheckRangeInt = True
    Exit Function
Catch:
    MsgBox "Please give a valid value for """ & PropName & """ in the range between (" & RangeMin & " - " & RangeMax & ")!" & vbCrLf & """" & StrValue & """ is not a valid value."
End Function

Private Function TryParseNCheckRangeSng(ByVal StrValue As String, ByVal PropName As String, ByVal RangeMin As Double, ByVal RangeMax As Double, ByRef Value_inout As Single) As Boolean
Try: On Error GoTo Catch
    Dim v As Single: v = Val(Replace(StrValue, ",", "."))
    If v < RangeMin Or RangeMax < v Then GoTo Catch
    Value_inout = v
    TryParseNCheckRangeSng = True
    Exit Function
Catch:
    MsgBox "Please give a valid value for """ & PropName & """ in the range between (" & RangeMin & " - " & RangeMax & ")!" & vbCrLf & """" & StrValue & """ is not a valid value."
End Function

