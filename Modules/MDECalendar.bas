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

Public AllGermanLegalFestivals() As LegalFestival
Private m_LegalFestivals_Initialized As Boolean


' v ############################## v '       the legal and religious holidays / festivals       ' v ############################## v '
Private Function New_LegalFestival(ByVal aDate As Date, ByVal aFest As ELegalFestivals, ByVal aLand As EGermanLand) As LegalFestival
    With New_LegalFestival:  .Date = aDate: .Festival = aFest: .Land = aLand: End With
End Function

Public Function Advent1Sunday(ByVal Year As Integer) As Date
    Dim Nov26 As Date: Nov26 = DateSerial(Year, 11, 26)
    Dim wd As VbDayOfWeek: wd = Weekday(Nov26, VbDayOfWeek.vbMonday)
    Advent1Sunday = Nov26 + 7 - wd
End Function

'Public Enum ELegalFestivals
'    Neujahr = 1               ' 1  01.01.
'    HeiligeDreiKönige         ' 2  06.01.
'    InternationalerFrauentag  ' 3  08.03.
'    Karfreitag                ' 4  2 days before Ostersonntag
'    Ostersonntag              ' 5  calculate due to Gauss
'    Ostermontag               ' 6  1 day after Ostersonntag
'    TagDerArbeit              ' 7  01.05.
'    ChristiHimmelfahrt        ' 8  10 days before Pfingstsonntag
'    Pfingstsonntag            ' 9  7 weeks = 49 days after Ostersonntag
'    Pfingstmontag             '10  1 day after Pfingstsonntag
'    Fronleichnam              '11  10 days after Pfingstmontag
'    AugsburgerFriedensfest    '12  08.08.
'    MariaeHimmelfahrt         '13  15.08.
'    Weltkindertag             '14  20.09.
'    TagDerDeutschenEinheit    '15  03.10.
'    Reformationstag           '16  31.10.
'    Allerheiligen             '17  01.11.
'    BussUndBettag             '18  20.11
'    '                         '19  24.12.
'    Weihnachtsfeiertag1 = 20  '20  25.12.
'    Weihnachtsfeiertag2       '21  26.12.
'    '                         '22  31.12.
'End Enum
'
'Public Enum EContractFestivals
'    Heiligabend = 19          '21  24.12. (according to agreement half holiday)
'    Silvester = 22            '22  31.12. (according to agreement half holiday)
'End Enum

Public Function ELegalFestivals_ToStr(ByVal e As ELegalFestivals) As String
    Dim s As String
    Select Case e
    Case ELegalFestivals.Neujahr:                   s = "Neujahr"                    ' 1  01.01.
    Case ELegalFestivals.HeiligeDreiKönige:         s = "Heilige Drei Könige"        ' 2  06.01.
    Case ELegalFestivals.InternationalerFrauentag:  s = "Internationaler Frauentag"  ' 3  08.03.
    Case ELegalFestivals.Karfreitag:                s = "Karfreitag"                 ' 4  2 days before Ostersonntag"
    Case ELegalFestivals.Ostersonntag:              s = "Ostersonntag"               ' 5  calculate accodring to Gauss"
    Case ELegalFestivals.Ostermontag:               s = "Ostermontag"                ' 6  1 day after Ostersonntag"
    Case ELegalFestivals.TagDerArbeit:              s = "Tag Der Arbeit"             ' 7  01.05."
    Case ELegalFestivals.ChristiHimmelfahrt:        s = "Christi Himmelfahrt"        ' 8  10 days before Pfingstsonntag"
    Case ELegalFestivals.Pfingstsonntag:            s = "Pfingstsonntag"             ' 9  7 weeks = 49 days after Ostersonntag"
    Case ELegalFestivals.Pfingstmontag:             s = "Pfingstmontag"              '10  1 day after Pfingstsonntag"
    Case ELegalFestivals.Fronleichnam:              s = "Fronleichnam"               '11  10 days after Pfingstmontag"
    Case ELegalFestivals.AugsburgerFriedensfest:    s = "Augsburger Friedensfest"    '12  08.08."
    Case ELegalFestivals.MariaeHimmelfahrt:         s = "Mariae Himmelfahrt"         '13  15.08."
    Case ELegalFestivals.Weltkindertag:             s = "Weltkindertag"              '14  20.09."
    Case ELegalFestivals.TagDerDeutschenEinheit:    s = "Tag De rDeutschenEinheit"   '15  03.10."
    Case ELegalFestivals.Reformationstag:           s = "Reformationstag"            '16  31.10."
    Case ELegalFestivals.Allerheiligen:             s = "Allerheiligen"              '17
    Case ELegalFestivals.BussUndBettag:             s = "Buß- und Bettag"            '18  20.11
    Case EContractFestivals.Heiligabend:            s = "Heiligabend"                '19  24.12.
    Case ELegalFestivals.Weihnachtsfeiertag1:       s = "Erster Weihnachtsfeiertag"  '20  25.12.
    Case ELegalFestivals.Weihnachtsfeiertag2:       s = "Zweiter Weihnachtsfeiertag" '21  26.12.
    Case EContractFestivals.Silvester:              s = "Silvester"                  '22  31.12.
    End Select
    ELegalFestivals_ToStr = s
End Function

'Public Type LegalFestival
'    Date     As Date
'    Festival As ELegalFestivals
'    Land     As EGermanLand
'End Type
'
'Public AllGermanLegalFestivals() As LegalFestival
'Private m_LegalFestivals_Initialized As Boolean

Public Sub InitAllLegalFestivals(ByVal Year As Integer)
    Dim EasterSunday As Date: EasterSunday = MTime.CalcEasterdateGaussCorrected1900(Year)
    ReDim AllGermanLegalFestivals(0 To 22)
    Dim i As Long
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 1, 1), ELegalFestivals.Neujahr, EGermanLand.AllLands)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 1, 6), ELegalFestivals.HeiligeDreiKönige, EGermanLand.BadenWuerttemberg Or EGermanLand.Bayern Or EGermanLand.SachsenAnhalt)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 3, 8), ELegalFestivals.InternationalerFrauentag, EGermanLand.Berlin Or EGermanLand.MecklenburgVorpommern)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(EasterSunday - 2, ELegalFestivals.Karfreitag, EGermanLand.AllLands)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(EasterSunday, ELegalFestivals.Ostersonntag, EGermanLand.AllLands)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(EasterSunday + 1, ELegalFestivals.Ostermontag, EGermanLand.AllLands)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 5, 1), ELegalFestivals.TagDerArbeit, EGermanLand.AllLands)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(EasterSunday + 39, ELegalFestivals.ChristiHimmelfahrt, EGermanLand.AllLands)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(EasterSunday + 49, ELegalFestivals.Pfingstsonntag, EGermanLand.AllLands)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(EasterSunday + 50, ELegalFestivals.Pfingstmontag, EGermanLand.AllLands)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(EasterSunday + 61, ELegalFestivals.Fronleichnam, EGermanLand.BadenWuerttemberg Or EGermanLand.Bayern Or EGermanLand.Hessen Or EGermanLand.NordrheinWestfalen Or EGermanLand.Rheinlandpfalz Or EGermanLand.Saarland Or EGermanLand.Sachsen)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 8, 8), ELegalFestivals.AugsburgerFriedensfest, EGermanLand.Bayern_Augsburg)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 8, 15), ELegalFestivals.MariaeHimmelfahrt, EGermanLand.Saarland Or EGermanLand.Bayern)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 9, 20), ELegalFestivals.Weltkindertag, EGermanLand.Thueringen)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 10, 3), ELegalFestivals.TagDerDeutschenEinheit, EGermanLand.AllLands)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 10, 31), ELegalFestivals.Reformationstag, EGermanLand.Brandenburg Or EGermanLand.Bremen Or EGermanLand.Hamburg Or EGermanLand.MecklenburgVorpommern Or EGermanLand.Niedersachsen Or EGermanLand.Sachsen Or EGermanLand.SachsenAnhalt Or EGermanLand.SchleswigHolstein Or EGermanLand.Thueringen)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 11, 1), ELegalFestivals.Allerheiligen, EGermanLand.BadenWuerttemberg Or EGermanLand.Bayern Or EGermanLand.NordrheinWestfalen Or EGermanLand.Rheinlandpfalz Or EGermanLand.Saarland)
    
    Dim Adv1Sund As Date: Adv1Sund = Advent1Sunday(Year)
    'Der Buß- und Bettag findet jedes Jahr am Mittwoch vor Totensonntag und damit genau elf Tage vor dem ersten Adventssonntag statt
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(Adv1Sund - 11, ELegalFestivals.BussUndBettag, EGermanLand.Sachsen)
    
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 12, 24), EContractFestivals.Heiligabend, EGermanLand.AllLands)
    
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 12, 25), ELegalFestivals.Weihnachtsfeiertag1, EGermanLand.AllLands)
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 12, 26), ELegalFestivals.Weihnachtsfeiertag2, EGermanLand.AllLands)
    
    i = i + 1:    AllGermanLegalFestivals(i) = New_LegalFestival(DateSerial(Year, 12, 31), EContractFestivals.Silvester, EGermanLand.AllLands)
        
    m_LegalFestivals_Initialized = True
End Sub

Public Function GetLegalFestivals(ByVal Year As Integer, Optional ByVal GermanLand As EGermanLand = EGermanLand.AllLands) As LegalFestival()
    If Not m_LegalFestivals_Initialized Then InitAllLegalFestivals
    ReDim lf(0 To 50) As LegalFestival
    Dim i As Long
    Dim gl As EGermanLand
    'do while until i For i = 1 To 50
    lf(i) = New_LegalFestival(MTime.New_Date(Year, 1, 1, 0, 0, 0), ELegalFestivals.Neujahr, EGermanLand.AllLands): i = i + 1
    If GermanLand And EGermanLands.Bayern Then
        lf(i) = New_LegalFestival(MTime.New_Date(Year, 1, 1, 0, 0, 0), ELegalFestivals.Neujahr, EGermanLand.Bayern): i = i + 1
        lf(i) = New_LegalFestival(MTime.New_Date(Year, 1, 1, 0, 0, 0), ELegalFestivals.Neujahr, EGermanLand.BadenWuerttemberg): i = i + 1
        lf(i) = New_LegalFestival(MTime.New_Date(Year, 1, 1, 0, 0, 0), ELegalFestivals.Neujahr, EGermanLand.Bayern_Augsburg): i = i + 1
    End If
    lf(i) = New_LegalFestival(MTime.New_Date(Year, 1, 1, 0, 0, 0), ELegalFestivals.HeiligeDreiKönige, EGermanLand.AllLands): i = i + 1
    lf(i) = New_LegalFestival(MTime.New_Date(Year, 8, 3, 0, 0, 0), ELegalFestivals.Neujahr, EGermanLand.AllLands): i = i + 1
        
    'Next
End Function

