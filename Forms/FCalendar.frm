VERSION 5.00
Begin VB.Form FCalendar 
   Caption         =   "Form1"
   ClientHeight    =   9690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16350
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   16350
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CmbYear 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   60
      Width           =   1215
   End
   Begin VB.PictureBox PBCalendar 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      FillColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   9015
      Left            =   0
      ScaleHeight     =   599
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1079
      TabIndex        =   0
      Top             =   480
      Width           =   16215
   End
   Begin VB.Label LblYear 
      AutoSize        =   -1  'True
      Caption         =   "Year:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "FCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Calendar As MDECalendar.CalendarYear
Private m_CalView  As MDECalendar.CalendarView

Private Sub Form_Load()
    FillCmbYears
End Sub

Private Sub FillCmbYears()
    CmbYear.Clear
    Dim y As Integer
    For y = 1900 To 2100
        CmbYear.AddItem y
    Next
    CmbYear.ListIndex = Year(Now) - 1900 + 1
End Sub

Private Sub CmbYear_Click()
    Dim y As Integer: y = CInt(CmbYear.Text)
    m_Calendar = MDECalendar.New_CalendarYear(y)
    m_CalView = MDECalendar.New_CalendarView(m_Calendar, Me.PBCalendar)
    PBCalendar.Refresh
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single: T = PBCalendar.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then PBCalendar.Move L, T, W, H
End Sub

Private Sub PBCalendar_Paint()
    PBCalendar.Cls
    MDECalendar.CalendarView_DrawYear m_CalView, m_Calendar
End Sub

Private Sub PBCalendar_Resize()
    PBCalendar.Cls
    MDECalendar.CalendarView_DrawYear m_CalView, m_Calendar
End Sub
