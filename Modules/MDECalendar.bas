Attribute VB_Name = "MDECalendar"
Option Explicit

'for computing the legal or religious festivals / holidays of one year in every country of germany
'in this enum-series the exponent of the enum const matches the land-key (see AGS: https://de.wikipedia.org/wiki/Amtlicher_Gemeindeschl%C3%BCssel)
Public Enum EGermanLand
    SchleswigHolstein = &H2&        ' 2 ^ 01 ' Land
    Hamburg = &H4&                  ' 2 ^ 02 ' Freie und Hansestadt
    Niedersachsen = &H8&            ' 2 ^ 03 ' Land
    Bremen = &H10&                  ' 2 ^ 04 ' Freie und Hansestadt
    NordrheinWestfalen = &H20&      ' 2 ^ 05 ' Land
    Hessen = &H40&                  ' 2 ^ 06 ' Land
    Rheinlandpfalz = &H80&          ' 2 ^ 07 ' Land
    BadenWuerttemberg = &H100&      ' 2 ^ 08 ' Land
    Bayern = &H200&                 ' 2 ^ 09 ' Freistaat
    Bayern_Augsburg = &H201&        '
    Saarland = &H400&               ' 2 ^ 10 ' Land
    Berlin = &H800&                 ' 2 ^ 11 ' Stadtstaat
    Brandenburg = &H1000&            ' 2 ^ 12 ' Land
    MecklenburgVorpommern = &H2000& ' 2 ^ 13 ' Land
    Sachsen = &H4000&               ' 2 ^ 14 ' Freistaat
    SachsenAnhalt = &H8000&         ' 2 ^ 15 ' Land
    Thueringen = &H10000            ' 2 ^ 16 ' Freistaat
    AllLands = &H1FFFE
End Enum

'Public Enum ReligiousFestival
'
'End Enum
Public Enum ELegalFestivals
    Neujahr = 1               ' 1  01.01.
    HeiligeDreiKönige         ' 2  06.01.
    InternationalerFrauentag  ' 3  08.03.
    Karfreitag                ' 4  2 days before Ostersonntag
    Ostersonntag              ' 5  calculate according to Gauss
    Ostermontag               ' 6  1 day after Ostersonntag
    TagDerArbeit              ' 7  01.05.
    ChristiHimmelfahrt        ' 8  10 days before Pfingstsonntag
    Pfingstsonntag            ' 9  7 weeks = 49 days after Ostersonntag
    Pfingstmontag             '10  1 day after Pfingstsonntag
    Fronleichnam              '11  10 days after Pfingstmontag
    AugsburgerFriedensfest    '12  08.08.
    MariaeHimmelfahrt         '13  15.08.
    Weltkindertag             '14  20.09.
    TagDerDeutschenEinheit    '15  03.10.
    Reformationstag           '16  31.10.
    Allerheiligen             '17  01.11.
    BussUndBettag             '18  20.11
    '                         '19  24.12.
    Weihnachtsfeiertag1 = 20  '20  25.12.
    Weihnachtsfeiertag2       '21  26.12.
    '                         '22  31.12.
End Enum

Public Enum EContractFestivals
    Heiligabend = 19          '19  24.12. (according to agreement half holiday)
    Silvester = 22            '22  31.12. (according to agreement half holiday)
End Enum

Public Type LegalFestival
    Date     As Date
    Festival As ELegalFestivals
    Land     As EGermanLand
End Type

'Public Festivals() As LegalFestival
'Private m_Festivals_Initialized As Boolean

Public Type CalendarDay
    Day  As Integer
    Date As Date
    FestivalIndex As Integer '0 = no festivalday
End Type
Public Type CalendarMonth
    Year   As Integer
    Month  As Integer
    Days() As CalendarDay
End Type
Public Type CalendarYear
    Year     As Integer
    Months() As CalendarMonth
    Fests()  As LegalFestival
End Type

Public Type CalendarView
    Canvas          As PictureBox
    HasDecLastYear  As Boolean
    HasJanNextYear  As Boolean
    HasMonthNames   As Boolean
    HasWeekDayNames As Boolean
    HasWeekNumbers  As Boolean
    MarginCalLeft   As Single
    MarginCalTop    As Single
    MarginCalRight  As Single
    MarginCalBottom As Single
    MarginMonLeft   As Single
    MarginMonTop    As Single
    MarginMonRight  As Single
    MarginMonBottom As Single
    MarginDayLeft   As Single
    MarginDayTop    As Single
    MarginDayRight  As Single
    MarginDayBottom As Single
    ColorWeekday    As Long
    ColorSaturday   As Long
    ColorSunday     As Long
    ColorLNWeekday  As Long
    ColorLNSaturday As Long
    ColorLNSunday   As Long
    ColTmpWeekday   As Long
    ColTmpSaturday  As Long
    ColTmpSunday    As Long
    FontMonthName   As StdFont
    FontDayNrName   As StdFont
    FontWeekNrName  As StdFont
    TmpDayWidth     As Single
    TmpDayHeight    As Single
End Type

' v ############################## v '       the legal and religious holidays / festivals       ' v ############################## v '
Private Function New_LegalFestival(ByVal aDate As Date, ByVal aFest As ELegalFestivals, ByVal aLand As EGermanLand) As LegalFestival
    With New_LegalFestival:  .Date = aDate: .Festival = aFest: .Land = aLand: End With
End Function

Public Function ELegalFestivals_ToStr(ByVal e As ELegalFestivals) As String
    Dim s As String
    Select Case e
    Case ELegalFestivals.Neujahr:                   s = "Neujahr"                    ' 1  01.01.
    Case ELegalFestivals.HeiligeDreiKönige:         s = "Heilige Drei Könige"        ' 2  06.01.
    Case ELegalFestivals.InternationalerFrauentag:  s = "Internat. Frauentag"  ' 3  08.03.
    Case ELegalFestivals.Karfreitag:                s = "Karfreitag"                 ' 4  2 days before Ostersonntag"
    Case ELegalFestivals.Ostersonntag:              s = "Ostersonntag"               ' 5  calculate accodring to Gauss"
    Case ELegalFestivals.Ostermontag:               s = "Ostermontag"                ' 6  1 day after Ostersonntag"
    Case ELegalFestivals.TagDerArbeit:              s = "Tag Der Arbeit"             ' 7  01.05."
    Case ELegalFestivals.ChristiHimmelfahrt:        s = "Christi Himmelf."        ' 8  10 days before Pfingstsonntag"
    Case ELegalFestivals.Pfingstsonntag:            s = "Pfingstsonntag"             ' 9  7 weeks = 49 days after Ostersonntag"
    Case ELegalFestivals.Pfingstmontag:             s = "Pfingstmontag"              '10  1 day after Pfingstsonntag"
    Case ELegalFestivals.Fronleichnam:              s = "Fronleichnam"               '11  10 days after Pfingstmontag"
    Case ELegalFestivals.AugsburgerFriedensfest:    s = "Augsbg. Friedensf."    '12  08.08."
    Case ELegalFestivals.MariaeHimmelfahrt:         s = "Mariae Himmelf."         '13  15.08."
    Case ELegalFestivals.Weltkindertag:             s = "Weltkindertag"              '14  20.09."
    Case ELegalFestivals.TagDerDeutschenEinheit:    s = "Tag d.Dt.Einheit"   '15  03.10."
    Case ELegalFestivals.Reformationstag:           s = "Reformationstag"            '16  31.10."
    Case ELegalFestivals.Allerheiligen:             s = "Allerheiligen"              '17
    Case ELegalFestivals.BussUndBettag:             s = "Buß- & Bettag"            '18  20.11
    Case EContractFestivals.Heiligabend:            s = "Heiligabend"                '19  24.12.
    Case ELegalFestivals.Weihnachtsfeiertag1:       s = "1. Weihnachtsf."  '20  25.12.
    Case ELegalFestivals.Weihnachtsfeiertag2:       s = "2. Weihnachtsf." '21  26.12.
    Case EContractFestivals.Silvester:              s = "Silvester"                  '22  31.12.
    End Select
    ELegalFestivals_ToStr = s
End Function

'Public Sub InitFestivals(ByVal Year As Integer)
'    Festivals = GetFestivals(Year)
'    m_Festivals_Initialized = True
'End Sub

Public Function GetFestivals(ByVal Year As Integer) As LegalFestival()
    Dim EasterSunday As Date: EasterSunday = MTime.OsternShort2(Year)
    ReDim Fests(0 To 22) As LegalFestival
    Dim i As Long
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 1, 1), ELegalFestivals.Neujahr, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 1, 6), ELegalFestivals.HeiligeDreiKönige, EGermanLand.BadenWuerttemberg Or EGermanLand.Bayern Or EGermanLand.SachsenAnhalt)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 3, 8), ELegalFestivals.InternationalerFrauentag, EGermanLand.Berlin Or EGermanLand.MecklenburgVorpommern)
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday - 2, ELegalFestivals.Karfreitag, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday, ELegalFestivals.Ostersonntag, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 1, ELegalFestivals.Ostermontag, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 5, 1), ELegalFestivals.TagDerArbeit, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 39, ELegalFestivals.ChristiHimmelfahrt, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 49, ELegalFestivals.Pfingstsonntag, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 50, ELegalFestivals.Pfingstmontag, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 60, ELegalFestivals.Fronleichnam, EGermanLand.BadenWuerttemberg Or EGermanLand.Bayern Or EGermanLand.Hessen Or EGermanLand.NordrheinWestfalen Or EGermanLand.Rheinlandpfalz Or EGermanLand.Saarland Or EGermanLand.Sachsen)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 8, 8), ELegalFestivals.AugsburgerFriedensfest, EGermanLand.Bayern_Augsburg)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 8, 15), ELegalFestivals.MariaeHimmelfahrt, EGermanLand.Saarland Or EGermanLand.Bayern)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 9, 20), ELegalFestivals.Weltkindertag, EGermanLand.Thueringen)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 10, 3), ELegalFestivals.TagDerDeutschenEinheit, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 10, 31), ELegalFestivals.Reformationstag, EGermanLand.Brandenburg Or EGermanLand.Bremen Or EGermanLand.Hamburg Or EGermanLand.MecklenburgVorpommern Or EGermanLand.Niedersachsen Or EGermanLand.Sachsen Or EGermanLand.SachsenAnhalt Or EGermanLand.SchleswigHolstein Or EGermanLand.Thueringen)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 11, 1), ELegalFestivals.Allerheiligen, EGermanLand.BadenWuerttemberg Or EGermanLand.Bayern Or EGermanLand.NordrheinWestfalen Or EGermanLand.Rheinlandpfalz Or EGermanLand.Saarland)
    
    Dim AdvSund1 As Date: AdvSund1 = AdventSunday1(Year)
    'Der Buß- und Bettag findet jedes Jahr am Mittwoch vor Totensonntag und damit genau elf Tage vor dem ersten Adventssonntag statt
    i = i + 1:    Fests(i) = New_LegalFestival(AdvSund1 - 11, ELegalFestivals.BussUndBettag, EGermanLand.Sachsen)
    
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 12, 24), EContractFestivals.Heiligabend, EGermanLand.AllLands)
    
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 12, 25), ELegalFestivals.Weihnachtsfeiertag1, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 12, 26), ELegalFestivals.Weihnachtsfeiertag2, EGermanLand.AllLands)
    
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 12, 31), EContractFestivals.Silvester, EGermanLand.AllLands)
    
    GetFestivals = Fests
End Function

Public Property Get Festivals_Index(this() As LegalFestival, ByVal aDate As Date) As Integer
    'returns the index in the array if aDate is a legal, religious or festival holiday otherwise 0
    Dim i As Integer
    For i = LBound(this) To UBound(this)
        If this(i).Date = aDate Then
            Festivals_Index = i
            Exit Property
        End If
    Next
End Property

'Public Function GetLegalFestivals(ByVal Year As Integer, Optional ByVal GermanLand As EGermanLand = EGermanLand.AllLands) As LegalFestival()
'    If Not m_LegalFestivals_Initialized Then InitAllLegalFestivals
'    ReDim lf(0 To 50) As LegalFestival
'    Dim i As Long
'    Dim gl As EGermanLand
'    'do while until i For i = 1 To 50
'    lf(i) = New_LegalFestival(MTime.New_Date(Year, 1, 1, 0, 0, 0), ELegalFestivals.Neujahr, EGermanLand.AllLands): i = i + 1
'    If GermanLand And EGermanLands.Bayern Then
'        lf(i) = New_LegalFestival(MTime.New_Date(Year, 1, 1, 0, 0, 0), ELegalFestivals.Neujahr, EGermanLand.Bayern): i = i + 1
'        lf(i) = New_LegalFestival(MTime.New_Date(Year, 1, 1, 0, 0, 0), ELegalFestivals.Neujahr, EGermanLand.BadenWuerttemberg): i = i + 1
'        lf(i) = New_LegalFestival(MTime.New_Date(Year, 1, 1, 0, 0, 0), ELegalFestivals.Neujahr, EGermanLand.Bayern_Augsburg): i = i + 1
'    End If
'    lf(i) = New_LegalFestival(MTime.New_Date(Year, 1, 1, 0, 0, 0), ELegalFestivals.HeiligeDreiKönige, EGermanLand.AllLands): i = i + 1
'    lf(i) = New_LegalFestival(MTime.New_Date(Year, 8, 3, 0, 0, 0), ELegalFestivals.Neujahr, EGermanLand.AllLands): i = i + 1
'
'    'Next
'End Function

Public Function New_CalendarYear(ByVal Year As Integer, Optional ByVal StartMonth As Integer = 1, Optional ByVal EndMonth As Integer = 12) As CalendarYear
    Dim y As CalendarYear
    y.Year = Year
    y.Fests = GetFestivals(Year)
    StartMonth = IIf(0 < StartMonth And StartMonth <= 12, StartMonth, 1)
    EndMonth = IIf(StartMonth < EndMonth And EndMonth <= 12, EndMonth, 12)
    ReDim y.Months(StartMonth To EndMonth)
    Dim m As Integer
    For m = StartMonth To EndMonth
        y.Months(m) = New_CalendarMonth(y, m)
    Next
    New_CalendarYear = y
End Function

Public Function New_CalendarMonth(CalYear As CalendarYear, ByVal Month As Integer) As CalendarMonth
    With New_CalendarMonth
        .Year = CalYear.Year
        .Month = Month
        Dim mds As Integer: mds = DaysInMonth(.Year, Month)
        ReDim .Days(1 To mds)
        Dim d As Integer
        For d = 1 To mds
            .Days(d) = New_CalendarDay(CalYear, Month, d)
        Next
    End With
End Function

Public Function New_CalendarDay(CalYear As CalendarYear, ByVal Month As Integer, ByVal Day As Integer) As CalendarDay
    'Dim d As CalendarDay
    With New_CalendarDay
        .Day = Day
        .Date = DateSerial(CalYear.Year, Month, Day)
        .FestivalIndex = Festivals_Index(CalYear.Fests, .Date)
    End With
    'New_CalendarDay = d
End Function

Public Function New_StdFont(ByVal FontName As String) As StdFont
    Set New_StdFont = New StdFont: New_StdFont.Name = FontName
    New_StdFont.Size = 10
End Function

Public Function New_CalendarView(CalYear As CalendarYear, Canvas As PictureBox) As CalendarView
    With New_CalendarView
        Set .Canvas = Canvas
        .ColorWeekday = RGB(255, 255, 255)
        .ColorSaturday = RGB(230, 244, 253)
        .ColorSunday = RGB(137, 189, 226)
        .ColorLNWeekday = RGB(255, 255, 255)
        .ColorLNSaturday = RGB(200, 202, 201)
        .ColorLNSunday = RGB(157, 157, 157)
        Set .FontDayNrName = New_StdFont("Consolas")
        '.FontDayNrName.Size = "Consolas"
        Set .FontMonthName = New_StdFont("Consolas")
        Set .FontWeekNrName = New_StdFont("Consolas")
        .HasDecLastYear = True
        .HasJanNextYear = True
        .HasMonthNames = True
        .HasWeekDayNames = True
        .HasWeekNumbers = True
        .MarginCalLeft = 10 'px
        .MarginCalTop = 10 'px
        .MarginCalRight = 10 'px
        .MarginCalBottom = 10 'px
    End With
End Function

Public Property Get CalendarView_DayWidth(this As CalendarView, CalYear As CalendarYear) As Single
    With this
        Dim n As Long: n = UBound(CalYear.Months) - LBound(CalYear.Months) + 1 + IIf(.HasDecLastYear, 1, 0) + IIf(.HasJanNextYear, 1, 0)
        CalendarView_DayWidth = (.Canvas.ScaleWidth - .MarginCalLeft - .MarginCalRight) / n
    End With
End Property

Public Property Get CalendarView_DayHeight(this As CalendarView) As Single
    With this
        Dim n As Single: n = 32
        CalendarView_DayHeight = (.Canvas.ScaleHeight - .MarginCalTop - .MarginCalBottom - IIf(.HasMonthNames, .FontMonthName.Size, 0)) / n
    End With
End Property

Public Sub CalendarView_DrawYear(this As CalendarView, CalYear As CalendarYear)
    
    With this
        Dim nx As Integer
        .Canvas.CurrentX = .MarginCalLeft
        .Canvas.CurrentY = .MarginCalTop
        
        .TmpDayWidth = CalendarView_DayWidth(this, CalYear)
        .TmpDayHeight = CalendarView_DayHeight(this)
        
        '.Canvas.CurrentX = View.MarginCalLeft
        '.Canvas.CurrentY = View.MarginCalTop
        'View.Canvas.CurrentY
        If .HasDecLastYear Then
            Dim CalLastYear As CalendarYear:  CalLastYear = New_CalendarYear(CalYear.Year - 1, 12, 12)
            Dim DecLastYear As CalendarMonth: DecLastYear = New_CalendarMonth(CalLastYear, 12)
            .ColTmpWeekday = .ColorLNWeekday
            .ColTmpSaturday = .ColorLNSaturday
            .ColTmpSunday = .ColorLNSunday
            CalendarView_DrawMonth this, DecLastYear
            nx = nx + 1
            .Canvas.CurrentX = .MarginCalLeft + nx * .TmpDayWidth
        End If
        
        .ColTmpWeekday = .ColorWeekday
        .ColTmpSaturday = .ColorSaturday
        .ColTmpSunday = .ColorSunday

        Dim m As Integer
        For m = LBound(CalYear.Months) To UBound(CalYear.Months)
            CalendarView_DrawMonth this, CalYear.Months(m)
            nx = nx + 1
            .Canvas.CurrentX = .MarginCalLeft + nx * .TmpDayWidth
        Next
        
        If .HasJanNextYear Then
            Dim CalNextYear As CalendarYear:  CalNextYear = New_CalendarYear(CalYear.Year + 1, 1, 1)
            Dim JanNextYear As CalendarMonth: JanNextYear = New_CalendarMonth(CalNextYear, 1)
            .ColTmpWeekday = .ColorLNWeekday
            .ColTmpSaturday = .ColorLNSaturday
            .ColTmpSunday = .ColorLNSunday
            CalendarView_DrawMonth this, JanNextYear
            nx = nx + 1
            .Canvas.CurrentX = .MarginCalLeft + nx * .TmpDayWidth
        End If
    End With
End Sub

Public Sub CalendarView_DrawMonth(this As CalendarView, CalMonth As CalendarMonth)
    With this
        Dim x As Single: x = .Canvas.CurrentX
        Dim y As Single: y = .MarginCalTop
        Dim ny As Integer
        If .HasMonthNames Then
            .Canvas.FontName = .FontMonthName.Name
            Dim s As String: s = MonthName(CalMonth.Month) & " '" & Right(CStr(CalMonth.Year), 2)
            .Canvas.Print s
            ny = ny + 1
            .Canvas.CurrentY = .MarginCalTop + ny * .TmpDayHeight
        End If
        .Canvas.CurrentX = x
        Dim d As Integer
        Dim L As Integer: L = LBound(CalMonth.Days)
        Dim u As Integer: u = UBound(CalMonth.Days)
        For d = L To u
            CalendarView_DrawDay this, CalMonth.Days(d)
            ny = ny + 1
            .Canvas.CurrentY = .MarginCalTop + ny * .TmpDayHeight
        Next
        .Canvas.CurrentY = x
        .Canvas.CurrentY = y
    End With
End Sub

Public Sub CalendarView_DrawDay(this As CalendarView, CalDay As CalendarDay)
    With this
        Dim x As Single: x = .Canvas.CurrentX
        Dim y As Single: y = .Canvas.CurrentY
        Dim wd As VbDayOfWeek: wd = Weekday(CalDay.Date)
        Dim c As Long: c = IIf(wd = vbSaturday, .ColTmpSaturday, IIf(wd = VbDayOfWeek.vbSunday, .ColTmpSunday, .ColTmpWeekday))
        
        .Canvas.Line (x, y)-(x + .TmpDayWidth - 1, y + .TmpDayHeight - 1), c, BF
        If CalDay.FestivalIndex Then
            c = RGB(222, 141, 245)
            .Canvas.Line (x, y)-(x + .TmpDayWidth - 1, y + .TmpDayHeight - 1), c, B
        End If
        
        .Canvas.CurrentX = x
        .Canvas.CurrentY = y
        
        Dim s As String
        s = CStr(CalDay.Day) & " " & VbWeekDay_ToStr(wd, vbSunday, True)
        If CalDay.FestivalIndex Then
            s = s & " " & MDECalendar.ELegalFestivals_ToStr(CalDay.FestivalIndex)
        End If
        
        If Weekday(CalDay.Date) = vbSunday Then
            c = RGB(255, 255, 255)
        Else
            c = RGB(0, 0, 0)
        End If
        .Canvas.ForeColor = c
        .Canvas.Print s
        'View.Canvas.Print VbWeekDay_ToStr(wd, vbSunday, True)
        
        .Canvas.CurrentX = x
        .Canvas.CurrentY = y
        
    End With
End Sub
