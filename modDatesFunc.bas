Attribute VB_Name = "modDatesFunc"
Option Explicit

'********************************************************************************
' Created By S.Y. Kim
' Create date/time:11/07/2002, 02:00:41
' This module includes the following public functions:
' Major functions
'     Public Function DateFunc(...) As Variant
'     Public Function DateMatch(...) As Date
'     Public Function DateMatchDesc(...) As String
'Complemetary Functions for decription of enumerations
'     Public Function eDateFuncDesc(...) As String
'     Public Function eLanguageDesc(...) As String
'     Public Function eVbDayOccurrenceDesc(...) As String
'     Public Function eVbMonthDesc(...) As String
'IMPORTANT NOTE:
'     You should need to localize or add language features for the following  functions,
'     because they requries some hard-codings. (No language-aware API exists for them.)
'     Private Function ReplaceTimeMarker(...) As String - needs to add time marker e.g. AM/PM
'     Public Function DateMatchDesc(...) As String - needs to adjust sequences
'     Public Function eVbDayOccurrenceDesc(...) As String - needs to add ordinal numer strings e.g. First
'********************************************************************************


'NOTE:

Public Enum eLanguage
   eLanguage_Default = &H400&
   Neutral = &H0& '0
   Arabic = &H1& '1
   Bulgarian = &H2& '2
   Catalan = &H3& '3
   Chinese = &H4& '4
   Czech = &H5& '5
   Danish = &H6& '6
   German = &H7& '7
   Greek = &H8& '8
   English = &H9& '9
   Spanish = &HA& '10
   Finnish = &HB& '11
   French = &HC& '12
   Hebrew = &HD& '13
   Hungarian = &HE& '14
   Icelandic = &HF& '15
   Italian = &H10& '16
   Japanese = &H11& '17
   Korean = &H12& '18
   Dutch = &H13& '19
   Norwegian = &H14& '20
   Polish = &H15& '21
   Portuguese = &H16& '22
   Romanian = &H18& '24
   Russian = &H19& '25
   Croatian = &H1A& '26
   Serbian = &H1A& '26
   Slovak = &H1B& '27
   Albanian = &H1C& '28
   Swedish = &H1D& '29
   Thai = &H1E& '30
   Turkish = &H1F& '31
   Urdu = &H20& '32
   Indonesian = &H21& '33
   Ukrainian = &H22& '34
   Belarusian = &H23& '35
   Slovenian = &H24& '36
   Estonian = &H25& '37
   Latvian = &H26& '38
   Lithuanian = &H27& '39
   Farsi = &H29& '41
   Armenian = &H2B& '43
   Azeri = &H2C& '44
   Basque = &H2D& '45
   Macedonian = &H2F& '47
   Afrikaans = &H36& '54
   Georgian = &H37& '55
   Faeroese = &H38& '56
   Hindi = &H39& '57
   Malay = &H3E& '62
   Kazak = &H3F& '63
   Swahili = &H41& '65
   Uzbek = &H43& '67
   Tatar = &H44& '68
   Bengali = &H45& '69
   Punjabi = &H46& '70
   Gujarati = &H47& '71
   Oriya = &H48& '72
   Tamil = &H49& '73
   Telugu = &H4A& '74
   Kannada = &H4B& '75
   Malayalam = &H4C& '76
   Assamese = &H4D& '77
   Marathi = &H4E& '78
   Sanskrit = &H4F& '79
   Konkani = &H57& '87
   Manipuri = &H58& '88
   Sindhi = &H59& '89
   Kashmiri = &H60& '96
   Nepali = &H61& '97
   
   eLanguage_MIN = 0
   eLanguage_MAX = Nepali
End Enum 'eLanguage


Public Enum eDateFunc
   [dtYY/MM/DD] = 0&
   [dtYYYY/MM/DD]
   dtYYMMDD
   dtYYMM
   dtMMDD
   [dtYYMMDD-HHMMSS]
   [dtYY/MM/DD-HH:MM:SS]
   [dtYYYY/MM/DD-HH:MM:SS]
   [dtYYYY/MM/DD AM/PM HH:MM:SS]
   [dtYYYY/MM/DD AM/PM H:M:S]
   dtHHMMSS
   [dtHH:MM]
   [dtMM:SS]
   [dtHH:MM:SS]
   dtDateSerial
   dtTimeSerial
   dtLongDate
   dtShortDate
   dtLongTime
   dtShortTime
   dtLongDateTime
   dtShortDateTime
   dtLongDateShortTime
   dtShortDateLongTime
   dtFileDateTime
   dtYear
   dtQuarter
   dtMonth
   dtMonthName
   dtShortMonthName
   dtDay
   dtDayOfTheYear
   dtWeekday
   dtWeekdayName
   dtShortWeekdayName
   dtWeekOfTheYear
   dtHour
   dtMinute
   dtSecond
   dtNextMonthLastDate
   dtMonthLastDate
   dtDaysInMonth
   dtWeeksInMonth
   dtWeekInMonth
   dtIsReapYear
   dtFirstDateOfYear
   dtFirstWeekDayOfYear
   dtFirstWeekdayNameOfYear
   dtFirstDateOfWeek
   dtLastDateOfYear
   dtLastWeekDayOfYear
   dtLastWeekdayNameOfYear
   dtLastDateOfWeek
   dtNextDate
   dtNextWeekDate
   dtNextMonthDate
   dtNextYearDate
   dtPreviousDate
   dtPreviousWeekDate
   dtPreviousMonthDate
   dtPreviousYearDate
   dtDateMatchDesc
   dtDateMatch
   dtDateFunc_MIN = [dtYY/MM/DD]
   dtDateFunc_MAX = dtDateMatch
End Enum

Public Enum eDateFunc2
   dtIntervalYear
   dtIntervalQuarter
   dtIntervalMonth
   dtIntervalDayOfYear
   dtIntervalDay
   dtIntervalWeekday
   dtIntervalWeek
   dtIntervalHour
   dtIntervalMinute
   dtIntervalSecond
   eDateFunc2_MIN = dtIntervalYear
   eDateFunc2_MAX = dtIntervalSecond
End Enum

Public Enum eVbMonth
   vbJanuary = 1
   vbFebruary = 2
   vbMarch = 3
   vbApril = 4
   vbMay = 5
   vbJune = 6
   vbJuly = 7
   vbAugust = 8
   vbSeptember = 9
   vbOctober = 10
   vbNovember = 11
   vbDecember = 12
End Enum

Public Enum eVbDayOccurrence
   vbFirst = 1
   vbSecond = 2
   vbThird = 3
   vbFourth = 4
   vbFifth = 5
   vbLast = 6
End Enum

Public Enum eMonth
   JanuaryLong = &H38         '  long name for January
   FebruaryLong = &H39         '  long name for February
   MarchLong = &H3A         '  long name for March
   AprilLong = &H3B         '  long name for April
   MayLong = &H3C         '  long name for May
   JuneLong = &H3D         '  long name for June
   JulyLong = &H3E         '  long name for July
   AugustLong = &H3F         '  long name for August
   SeptemberLong = &H40         '  long name for September
   OctoberLong = &H41        '  long name for October
   NovemberLong = &H42        '  long name for November
   DecemberLong = &H43        '  long name for December

   January = &H44   '  abbreviated name for January
   February = &H45   '  abbreviated name for February
   March = &H46   '  abbreviated name for March
   April = &H47   '  abbreviated name for April
   May = &H48   '  abbreviated name for May
   June = &H49   '  abbreviated name for June
   July = &H4A   '  abbreviated name for July
   August = &H4B   '  abbreviated name for August
   September = &H4C   '  abbreviated name for September
   October = &H4D  '  abbreviated name for October
   November = &H4E  '  abbreviated name for November
   December = &H4F  '  abbreviated name for December
   
   eMonth_MIN = JanuaryLong
   eMonth_MAX = December
End Enum

Public Enum eWeekDay
   Monday = &H31     '  abbreviated name for Monday
   Tuesday = &H32     '  abbreviated name for Tuesday
   Wednesday = &H33     '  abbreviated name for Wednesday
   Thursday = &H34     '  abbreviated name for Thursday
   Friday = &H35     '  abbreviated name for Friday
   Saturday = &H36     '  abbreviated name for Saturday
   Sunday = &H37     '  abbreviated name for Sunday

   MondayLong = &H2A           '  long name for Monday
   TuesdayLong = &H2B           '  long name for Tuesday
   WednesdayLong = &H2C           '  long name for Wednesday
   ThursdayLong = &H2D           '  long name for Thursday
   FridayLong = &H2E           '  long name for Friday
   SaturdayLong = &H2F           '  long name for Saturday
   SundayLong = &H30           '  long name for Sunday
   eWeekDay_MIN = MondayLong
   eWeekDay_MAX = Sunday
End Enum

Public Enum eLCInfo
   DecimalSeparator = &HE             '  decimal separator
   ThousandSeparator = &HF            '  thousand separator
   DigitGrouping = &H10           '  digit grouping
   NumberOfFractionalDigits = &H11             '  number of fractional digits
   LeadingZerosForDecimal = &H12              '  leading zeros for decimal
   DateSeparator = &H1D               '  date separator
   TimeSeparator = &H1E               '  time separator
   ShortDateFormat = &H1F          '  short date format string
   CurrencySymbol = &H14           '  local monetary symbol
   LongDateFormat = &H20           '  long date format string
   
   PositiveSign = &H50       '  positive sign
   NegativeSign = &H51       '  negative sign
   PositiveSignPosition = &H52        '  positive sign position
   NegativeSignPosition = &H53        '  negative sign position
   LanguageEnglishName = &H1001      '  English name of language
   CountryEnglishName = &H1002       '  English name of country
   TimeFormat = &H1003       '  time format string
   
   eLCInfo_MIN = DecimalSeparator
   eLCInfo_MAX = TimeFormat
End Enum 'eLCInfo

Private Declare Function apiGetLocaleInfo Lib "Kernel32" Alias "GetLocaleInfoA" (ByVal cnLanguage As Long, ByVal LCType As eLCInfo, ByVal lpLCData As String, ByVal cchData As Long) As Long


Public Sub TestDateFunc()
   Dim i As Long
   Dim j As eDateFunc

   For j = dtDateFunc_MIN To dtDateFunc_MAX
      Debug.Print eDateFuncDesc(j) & "=" & DateFunc(j, Now(), , , , English)
   Next j
   Debug.Print ""
   
   Debug.Print DateMatchDesc(2002, vbNovember, vbTuesday, vbFourth, , English)
   Debug.Print DateMatch(2002, vbNovember, vbTuesday, vbFourth)
   Debug.Print ""
   
   For j = dtDateFunc_MIN To dtDateFunc_MAX
      Debug.Print eDateFuncDesc(j) & "=" & DateFunc(j, Now() + 4, , , , English)
   Next j

End Sub


'Enhanced VB DateFunc function
Public Function DateFunc( _
   ByVal Interval As eDateFunc, _
   Optional ByVal newDate As Date, _
   Optional FirstDayOfWeek As VbDayOfWeek = vbSunday, _
   Optional FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1, _
   Optional ByVal iOccurrence As eVbDayOccurrence = eVbDayOccurrence.vbFirst, _
   Optional ByVal cnLanguage As eLanguage = eLanguage.eLanguage_Default) As Variant
   
    If newDate = 0 Then
      newDate = VBA.Now()
    End If
    
    'Select the right and faster date/time function
    Select Case Interval
      Case eDateFunc.[dtYY/MM/DD]
         DateFunc = VBA.Replace(VBA.Format$(newDate, "yy-mm-dd"), "-", "/")
         
      Case eDateFunc.[dtYYYY/MM/DD]
         DateFunc = VBA.Replace(VBA.Format$(newDate, "yyyy-mm-dd"), "-", "/")
      
      Case eDateFunc.dtYYMMDD
         DateFunc = VBA.Format$(newDate, "yymmdd")
      
      Case eDateFunc.dtYYMM
         DateFunc = VBA.Format$(newDate, "yymm")
      
      Case eDateFunc.dtMMDD
         DateFunc = VBA.Format$(newDate, "mmdd")
      
      Case eDateFunc.[dtYYMMDD-HHMMSS]
         DateFunc = VBA.Format$(newDate, "yymmdd-hhmmss")
      
      Case eDateFunc.[dtYY/MM/DD-HH:MM:SS]
         DateFunc = VBA.Replace(VBA.Format$(newDate, "yy mm dd-hh:mm:ss"), " ", "/")
      
      Case eDateFunc.[dtYYYY/MM/DD-HH:MM:SS]
         DateFunc = VBA.Replace(VBA.Format$(newDate, "yyyy mm dd-hh:mm:ss"), " ", "/")
      
      Case eDateFunc.[dtYYYY/MM/DD AM/PM HH:MM:SS]
         DateFunc = VBA.Replace(VBA.Format$(newDate, "yyyy-mm-dd AM/PM hh:mm:ss"), "-", "/")
      
      Case eDateFunc.[dtYYYY/MM/DD AM/PM H:M:S]
         DateFunc = VBA.Replace(VBA.Format$(newDate, "yyyy-mm-dd AM/PM h:m:s"), "-", "/")
      
      Case eDateFunc.dtHHMMSS
         DateFunc = VBA.Format$(newDate, "hhmmss")
      
      Case eDateFunc.[dtHH:MM]
         DateFunc = VBA.Format$(newDate, "hh:mm")
      
      Case eDateFunc.[dtMM:SS]
         DateFunc = VBA.Mid$(VBA.Format$(newDate, "hh:mm:ss"), 4)
      
      Case eDateFunc.[dtHH:MM:SS]
         DateFunc = VBA.Format$(newDate, "hh:mm:ss")
      
      Case eDateFunc.dtLongDate
         'DateFunc = VBA.Format$(newDate, "Long Date")
         'Debug.Print GetLocaleInfo(LongDateFormat, Korean)
         'Debug.Print GetLocaleInfo(LongDateFormat, English)
         'DateFunc = VBA.Format$(newDate, GetLocaleInfo(LongDateFormat, cnLanguage))
         DateFunc = GetLongDate(newDate, cnLanguage)
      
      Case eDateFunc.dtShortDate
         'DateFunc = VBA.Format$(newDate, "Short Date")
         'Debug.Print GetLocaleInfo(ShortDateFormat, Korean)
         'Debug.Print GetLocaleInfo(ShortDateFormat, English)
         'Debug.Print GetLocaleInfo(DateSeparator, Korean)
         'Debug.Print GetLocaleInfo(DateSeparator, English)
         'DateFunc = VBA.Replace(VBA.Format$(newDate, GetLocaleInfo(ShortDateFormat, cnLanguage)), _
                              GetLocaleInfo(DateSeparator, eLanguage_Default), GetLocaleInfo(DateSeparator, cnLanguage))
         DateFunc = GetShortDate(newDate, cnLanguage)
      
      Case eDateFunc.dtLongTime
         'DateFunc = VBA.Format$(newDate, "Long Time")
         'DateFunc = ReplaceTimeMarker(VBA.Format$(newDate, GetLongTimeFormatTemp(cnLanguage)), cnLanguage)
         DateFunc = GetLongTime(newDate, cnLanguage)
         
         'Debug.Print GetLocaleInfo(TimeFormat, Korean)
         'Debug.Print GetLocaleInfo(TimeFormat, English)
         'Debug.Print GetLocaleInfo(TimeSeparator, Korean)
         'Debug.Print GetLocaleInfo(TimeSeparator, English)
         'Debug.Print GetLongTimeFormatTemp(Korean)
         'Debug.Print GetLongTimeFormatTemp(English)
         'Debug.Print VBA.Format$(newDate, GetLongTimeFormatTemp(Korean))
         'Debug.Print VBA.Format$(newDate, GetLongTimeFormatTemp(English))
         'DateFunc = VBA.Format$(newDate, GetLongTimeFormatTemp(cnLanguage))
         'Debug.Print GetLocaleInfo(TimeFormat)
         'Debug.Print GetLocaleInfo(TimeFormat, cnLanguage)
         'DateFunc = VBA.Format$(newDate, GetLocaleInfo(TimeFormat, cnLanguage))
      
         
      Case eDateFunc.dtShortTime
         DateFunc = VBA.Format$(newDate, "Short Time")
      
      Case eDateFunc.dtLongDateTime
         'DateFunc = VBA.Format$(newDate, "Long Date") & " " & VBA.Format$(newDate, "Long Time")
         'DateFunc = VBA.Format$(newDate, "Long Date") & " " & VBA.Format$(newDate, "Long Time")
         DateFunc = GetLongDate(newDate, cnLanguage) & " " & GetLongTime(newDate, cnLanguage)
      
      Case eDateFunc.dtShortDateTime
         'DateFunc = VBA.Format$(newDate, "Short Date") & " " & VBA.Format$(newDate, "Short Time")
         DateFunc = GetShortDate(newDate, cnLanguage) & " " & GetShortTime(newDate, cnLanguage)
      
      Case eDateFunc.dtFileDateTime
         'DateFunc = VBA.Format$(newDate, "Long Time")
         'DateFunc = VBA.Format$(newDate, "Short Date") & " " & VBA.Left$(DateFunc, VBA.InStrRev(DateFunc, ":") - 1)
         DateFunc = GetFileDateTime(newDate, cnLanguage)
         
         
      
      Case eDateFunc.dtLongDateShortTime
         'DateFunc = VBA.Format$(newDate, "Long Date") & " " & VBA.Format$(newDate, "Short Time")
         DateFunc = GetLongDate(newDate, cnLanguage) & " " & GetShortTime(newDate, cnLanguage)
      
      Case eDateFunc.dtShortDateLongTime
         'DateFunc = VBA.Format$(newDate, "Short Date") & " " & VBA.Format$(newDate, "Long Time")
         DateFunc = GetShortDate(newDate, cnLanguage) & " " & GetLongTime(newDate, cnLanguage)
      
      Case eDateFunc.dtDateSerial
         DateFunc = VBA.DateSerial(VBA.Year(newDate), VBA.Month(newDate), VBA.Day(newDate))
      
      Case eDateFunc.dtTimeSerial
         DateFunc = VBA.TimeSerial(VBA.Hour(newDate), VBA.Minute(newDate), VBA.Second(newDate))
         DateFunc = GetLongTime(DateFunc, cnLanguage)
         
      Case eDateFunc.dtYear
          DateFunc = VBA.Year(newDate)
      
      Case eDateFunc.dtQuarter
          DateFunc = VBA.DatePart("q", newDate, FirstDayOfWeek, FirstWeekOfYear)
      
      Case eDateFunc.dtMonth
          DateFunc = VBA.Month(newDate)
          
      Case eDateFunc.dtMonthName
          DateFunc = MonthName(JanuaryLong + VBA.Month(newDate) - 1, cnLanguage)
      
      Case eDateFunc.dtShortMonthName
          DateFunc = MonthName(January + VBA.Month(newDate) - 1, cnLanguage)
          
      Case eDateFunc.dtDay
          DateFunc = VBA.Day(newDate)
      
      Case eDateFunc.dtDayOfTheYear
          DateFunc = VBA.DatePart("y", newDate, FirstDayOfWeek, FirstWeekOfYear)
      
      Case eDateFunc.dtWeekday
          DateFunc = VBA.Weekday(newDate, FirstDayOfWeek)
      
      Case eDateFunc.dtWeekdayName
          DateFunc = WeekdayName(eWeekDay.MondayLong + VBA.Weekday(newDate, VbDayOfWeek.vbMonday) - 1, cnLanguage)
          'DateFunc = VBA.Format(newDate, "dddd")
      
      Case eDateFunc.dtShortWeekdayName
          'DateFunc = VBA.Format(newDate, "ddd")
          DateFunc = WeekdayName(eWeekDay.Monday + VBA.Weekday(newDate, VbDayOfWeek.vbMonday) - 1, cnLanguage)
      
      Case eDateFunc.dtWeekOfTheYear
          DateFunc = VBA.DatePart("ww", newDate, FirstDayOfWeek, FirstWeekOfYear)
      
      Case eDateFunc.dtHour
          DateFunc = VBA.Hour(newDate)
      
      Case eDateFunc.dtMinute
          DateFunc = VBA.Minute(newDate)
      
      Case eDateFunc.dtSecond
          DateFunc = VBA.Second(newDate)
      
      Case eDateFunc.dtNextMonthLastDate
         'Add 2 to the current month's first date (VBA.Month(newDate - 1) & "/01/" & VBA.Year(newDate - 1)).
         'Then, subtract one day.
          DateFunc = VBA.DateAdd("d", -1, VBA.DateAdd("m", 2, VBA.Month(newDate - 1) & "/1/" & VBA.Year(newDate - 1)))
      
      Case eDateFunc.dtDaysInMonth
         DateFunc = VBA.DateDiff("D", VBA.Format$(newDate, "MM/01/YYYY"), VBA.DateAdd("M", 1, VBA.Format$(newDate, "MM/01/YYYY")))
         
      Case eDateFunc.dtWeeksInMonth
         DateFunc = VBA.DateDiff("D", VBA.Format$(newDate, "MM/01/YYYY"), VBA.DateAdd("M", 1, VBA.Format$(newDate, "MM/01/YYYY")))
         If DateFunc Mod 7 <> 0 Then
            DateFunc = DateFunc \ 7 + 1
         Else
            DateFunc = DateFunc \ 7
         End If
         
      Case eDateFunc.dtWeekInMonth
         DateFunc = VBA.Day(newDate) + VBA.Weekday(VBA.Format$(newDate, "MM/01/YYYY"), FirstDayOfWeek) - 1
         If DateFunc Mod 7 <> 0 Then
            DateFunc = DateFunc \ 7 + 1
         Else
            DateFunc = DateFunc \ 7
         End If
         
      Case eDateFunc.dtMonthLastDate
         DateFunc = VBA.DateSerial(VBA.Year(newDate), VBA.Month(newDate) + 1, 0)
      
      Case eDateFunc.dtIsReapYear
         DateFunc = (VBA.Year(newDate) Mod 4 = 0 And (VBA.Year(newDate) Mod 100 <> 0 Or VBA.Year(newDate) Mod 400 = 0))
      
      Case eDateFunc.dtFirstDateOfYear
         DateFunc = CDate("01/01/" & VBA.Year(newDate))
         
      Case eDateFunc.dtFirstWeekDayOfYear
         DateFunc = VBA.Weekday(CDate("01/01/" & VBA.Year(newDate)), FirstDayOfWeek)
      
      Case eDateFunc.dtFirstWeekdayNameOfYear
         DateFunc = WeekdayName(MondayLong + VBA.Weekday(CDate("01/01/" & VBA.Year(newDate)), vbMonday) - 1, cnLanguage)
      
      Case eDateFunc.dtFirstDateOfWeek
         newDate = newDate - VBA.TimeSerial(VBA.Hour(newDate), VBA.Minute(newDate), VBA.Second(newDate))
         DateFunc = newDate - VBA.Weekday(newDate, FirstDayOfWeek) + 1
         
      
      Case eDateFunc.dtLastDateOfYear
         DateFunc = CDate("12/31/" & VBA.Year(newDate))
         
      Case eDateFunc.dtLastWeekDayOfYear
         DateFunc = VBA.Weekday(CDate("12/31/" & VBA.Year(newDate)), FirstDayOfWeek)
      
      Case eDateFunc.dtLastWeekdayNameOfYear
         DateFunc = WeekdayName(MondayLong + VBA.Weekday(CDate("12/31/" & VBA.Year(newDate)), vbMonday) - 1, cnLanguage)
      
      Case eDateFunc.dtLastDateOfWeek
         newDate = newDate - VBA.TimeSerial(VBA.Hour(newDate), VBA.Minute(newDate), VBA.Second(newDate))
         DateFunc = newDate - VBA.Weekday(newDate, FirstDayOfWeek) + 7
         
      Case eDateFunc.dtNextDate
         newDate = newDate - VBA.TimeSerial(VBA.Hour(newDate), VBA.Minute(newDate), VBA.Second(newDate))
         DateFunc = newDate + 1
      
      Case eDateFunc.dtNextWeekDate
         newDate = newDate - VBA.TimeSerial(VBA.Hour(newDate), VBA.Minute(newDate), VBA.Second(newDate))
         DateFunc = newDate + 7
      
      Case eDateFunc.dtNextMonthDate
         newDate = newDate - VBA.TimeSerial(VBA.Hour(newDate), VBA.Minute(newDate), VBA.Second(newDate))
         DateFunc = VBA.DateAdd("M", 1, newDate)
      
      Case eDateFunc.dtNextYearDate
         newDate = newDate - VBA.TimeSerial(VBA.Hour(newDate), VBA.Minute(newDate), VBA.Second(newDate))
         DateFunc = VBA.DateSerial(VBA.Year(newDate) + 1, VBA.Month(newDate), VBA.Day(newDate))
         
      Case eDateFunc.dtPreviousDate
         newDate = newDate - VBA.TimeSerial(VBA.Hour(newDate), VBA.Minute(newDate), VBA.Second(newDate))
         DateFunc = newDate - 1
      
      Case eDateFunc.dtPreviousWeekDate
         newDate = newDate - VBA.TimeSerial(VBA.Hour(newDate), VBA.Minute(newDate), VBA.Second(newDate))
         DateFunc = newDate - 7
      
      Case eDateFunc.dtPreviousMonthDate
         newDate = newDate - VBA.TimeSerial(VBA.Hour(newDate), VBA.Minute(newDate), VBA.Second(newDate))
         DateFunc = VBA.DateAdd("M", -1, newDate)
      
      Case eDateFunc.dtPreviousYearDate
         newDate = newDate - VBA.TimeSerial(VBA.Hour(newDate), VBA.Minute(newDate), VBA.Second(newDate))
         DateFunc = VBA.DateSerial(VBA.Year(newDate) - 1, VBA.Month(newDate), VBA.Day(newDate))
         
      Case eDateFunc.dtDateMatchDesc
         DateFunc = DateMatchDesc(VBA.Year(newDate), VBA.Month(newDate), FirstDayOfWeek, iOccurrence, vbSunday, cnLanguage)
      
      Case eDateFunc.dtDateMatch
         DateFunc = DateMatch(VBA.Year(newDate), VBA.Month(newDate), FirstDayOfWeek, iOccurrence)
         
    End Select
End Function


Public Function DateMatch( _
   ByVal iYear As Integer, _
   ByVal iMonth As eVbMonth, _
   ByVal iWeekday As VbDayOfWeek, _
   ByVal iOccurrence As eVbDayOccurrence) As Date
   
   'Return the date matched with the specified conditions
   
   Dim intWeek As Integer
   Dim intWeekday As Integer
   Dim intDay As Integer
   Dim dtLastDay As Date
   Dim intLastDay As Integer
   Dim intDayTemp As Integer
      
   If iOccurrence = vbLast Then
      'GET LAST OCCURRENCE IN MONTH
      dtLastDay = VBA.DateSerial(iYear, iMonth + 1, 1 - 1)
      intLastDay = VBA.Day(dtLastDay)
      intWeekday = VBA.Weekday(dtLastDay) - 1
      intDay = intLastDay - (intWeekday - (iWeekday - 1))
   Else
      'GET SPECIFIED OCCURRENCE IN MONTH
      intWeek = 1 + ((iOccurrence - 1) * 7)
      intWeekday = VBA.Weekday(VBA.DateSerial(iYear, iMonth, intWeek))
      intDayTemp = iWeekday - intWeekday
      If intDayTemp < 0 Then
         intDayTemp = intDayTemp + 7
      End If
      intDay = (intWeek + intDayTemp)
   End If
   
   'CHECK TO SEE IF THERE IS NO Nth DAY OF THE MONTH
   If VBA.IsDate(iMonth & "/" & intDay & "/" & iYear) Then
      DateMatch = VBA.DateSerial(iYear, iMonth, intDay)
   End If
   
End Function

Public Function DateMatchDesc( _
   ByVal iYear As Integer, _
   ByVal iMonth As eVbMonth, _
   ByVal iWeekday As VbDayOfWeek, _
   ByVal iOccurrence As eVbDayOccurrence, _
   Optional FirstDayOfWeek As VbDayOfWeek = vbSunday, _
   Optional ByVal cnLanguage As eLanguage = eLanguage.eLanguage_Default) As String
   
   If cnLanguage = eLanguage.Korean Then
      If iWeekday = vbSunday Then
         DateMatchDesc = iYear & "³â " & MonthName(JanuaryLong + iMonth - 1, eLanguage.Korean) & " " & _
                                 eVbDayOccurrenceDesc(iOccurrence, eLanguage.Korean) & " " & _
                                 WeekdayName(eWeekDay.SundayLong, eLanguage.Korean)
      Else
         DateMatchDesc = iYear & "³â " & MonthName(JanuaryLong + iMonth - 1, eLanguage.Korean) & " " & _
                                 eVbDayOccurrenceDesc(iOccurrence, eLanguage.Korean) & " " & _
                                 WeekdayName(eWeekDay.MondayLong + iWeekday - 2, eLanguage.Korean)
      End If
                                 
   Else
      If iWeekday = vbSunday Then
         DateMatchDesc = eVbDayOccurrenceDesc(iOccurrence, cnLanguage) & " " & _
                                 WeekdayName(eWeekDay.SundayLong, cnLanguage) & " of " & _
                                 MonthName(JanuaryLong + iMonth - 1, cnLanguage) & " in " & iYear
      Else
         DateMatchDesc = eVbDayOccurrenceDesc(iOccurrence, cnLanguage) & " " & _
                                 WeekdayName(eWeekDay.MondayLong + iWeekday - 2, cnLanguage) & " of " & _
                                 MonthName(JanuaryLong + iMonth - 1, cnLanguage) & " in " & iYear
      End If
   End If

End Function


Public Function eDateFuncDesc( _
   Index As eDateFunc) As String
   Dim retVal As String
   Select Case Index
   Case eDateFunc.[dtYY/MM/DD]  '= 0&
      retVal = "YY/MM/DD"
   Case eDateFunc.[dtYYYY/MM/DD]
      retVal = "YYYY/MM/DD"
   Case eDateFunc.dtYYMMDD
      retVal = "YYMMDD"
   Case eDateFunc.dtYYMM
      retVal = "YYMM"
   Case eDateFunc.dtMMDD
      retVal = "MMDD"
   Case eDateFunc.[dtYYMMDD-HHMMSS]
      retVal = "YYMMDD-HHMMSS"
   Case eDateFunc.[dtYY/MM/DD-HH:MM:SS]
      retVal = "YY/MM/DD-HH:MM:SS"
   Case eDateFunc.[dtYYYY/MM/DD-HH:MM:SS]
      retVal = "YYYY/MM/DD-HH:MM:SS"
   Case eDateFunc.[dtYYYY/MM/DD AM/PM HH:MM:SS]
      retVal = "YYYY/MM/DD AM/PM HH:MM:SS"
   Case eDateFunc.[dtYYYY/MM/DD AM/PM H:M:S]
      retVal = "YYYY/MM/DD AM/PM H:M:S"
   Case eDateFunc.dtHHMMSS
      retVal = "HHMMSS"
   Case eDateFunc.[dtHH:MM]
      retVal = "HH:MM"
   Case eDateFunc.[dtMM:SS]
      retVal = "MM:SS"
   Case eDateFunc.[dtHH:MM:SS]
      retVal = "HH:MM:SS"
   Case eDateFunc.dtDateSerial
      retVal = "Date Serial"
   Case eDateFunc.dtTimeSerial
      retVal = "Time Serial"
   Case eDateFunc.dtLongDate
      retVal = "Long Date"
   Case eDateFunc.dtShortDate
      retVal = "Short Date"
   Case eDateFunc.dtLongTime
      retVal = "Long Time"
   Case eDateFunc.dtShortTime
      retVal = "Short Time"
   Case eDateFunc.dtLongDateTime
      retVal = "Long Date + Long Time"
   Case eDateFunc.dtShortDateTime
      retVal = "Short Date + Short Time"
   Case eDateFunc.dtLongDateShortTime
      retVal = "Long Date + Short Time"
   Case eDateFunc.dtShortDateLongTime
      retVal = "Short Date + Long Time"
   Case eDateFunc.dtFileDateTime
      retVal = "File Date Time"
   Case eDateFunc.dtYear
      retVal = "Year"
   Case eDateFunc.dtQuarter
      retVal = "Quarter"
   Case eDateFunc.dtMonth
      retVal = "Month"
   Case eDateFunc.dtMonthName
      retVal = "Month Name"
   Case eDateFunc.dtShortMonthName
      retVal = "Short Month Name"
   Case eDateFunc.dtDay
      retVal = "Day"
   Case eDateFunc.dtDayOfTheYear
      retVal = "Day Of The Year"
   Case eDateFunc.dtWeekday
      retVal = "Weekday"
   Case eDateFunc.dtWeekdayName
      retVal = "Weekday Name"
   Case eDateFunc.dtShortWeekdayName
      retVal = "Short Weekday Name"
   Case eDateFunc.dtWeekOfTheYear
      retVal = "Week Of The Year"
   Case eDateFunc.dtHour
      retVal = "Hour"
   Case eDateFunc.dtMinute
      retVal = "Minute"
   Case eDateFunc.dtSecond
      retVal = "Second"
   Case eDateFunc.dtNextMonthLastDate
      retVal = "Next Month's Last Date"
   Case eDateFunc.dtMonthLastDate
      retVal = "Month's Last Date"
   Case eDateFunc.dtDaysInMonth
      retVal = "Days In Month"
   Case eDateFunc.dtWeeksInMonth
      retVal = "Weeks In Month"
   Case eDateFunc.dtWeekInMonth
      retVal = "Week In Month"
   Case eDateFunc.dtIsReapYear
      retVal = "Is Reap Year"
   Case eDateFunc.dtFirstDateOfYear
      retVal = "First Date Of Year"
   Case eDateFunc.dtFirstWeekDayOfYear
      retVal = "First Weekday Of Year"
   Case eDateFunc.dtFirstWeekdayNameOfYear
      retVal = "First Weekday's Name Of Year"
   Case eDateFunc.dtFirstDateOfWeek
      retVal = "First Date Of Week"
   Case eDateFunc.dtLastDateOfYear
      retVal = "Last Date Of Year"
   Case eDateFunc.dtLastWeekDayOfYear
      retVal = "Last Weekday Of Year"
   Case eDateFunc.dtLastWeekdayNameOfYear
      retVal = "Last Weekday's Name Of Year"
   Case eDateFunc.dtLastDateOfWeek
      retVal = "Last Date Of Week"
   Case eDateFunc.dtNextDate
      retVal = "Next Date"
   Case eDateFunc.dtNextWeekDate
      retVal = "Next Week Date"
   Case eDateFunc.dtNextMonthDate
      retVal = "Next Month Date"
   Case eDateFunc.dtNextYearDate
      retVal = "Next Year Date"
   Case eDateFunc.dtPreviousDate
      retVal = "Previous Date"
   Case eDateFunc.dtPreviousWeekDate
      retVal = "Previous Week Date"
   Case eDateFunc.dtPreviousMonthDate
      retVal = "Previous Month Date"
   Case eDateFunc.dtPreviousYearDate
      retVal = "Previous Year Date"
   Case eDateFunc.dtDateMatchDesc
      retVal = "Date Match Description"
   Case eDateFunc.dtDateMatch
      retVal = "Date Match"
   Case eDateFunc.dtDateFunc_MIN  '= [dtYY/MM/DD]
      retVal = "DateFunc_MIN"
   Case eDateFunc.dtDateFunc_MAX  '= dtDateMatch
      retVal = "DateFunc_MAX"
   Case Else
   End Select 'eDateFunc
   eDateFuncDesc = retVal
End Function 'eDateFuncDesc


Public Function eVbMonthDesc( _
   Index As eVbMonth) As String
   Dim retVal As String
   Select Case Index
   Case eVbMonth.vbJanuary  '= 1
      retVal = "January"
   Case eVbMonth.vbFebruary  '= 2
      retVal = "February"
   Case eVbMonth.vbMarch  '= 3
      retVal = "March"
   Case eVbMonth.vbApril  '= 4
      retVal = "April"
   Case eVbMonth.vbMay  '= 5
      retVal = "May"
   Case eVbMonth.vbJune  '= 6
      retVal = "June"
   Case eVbMonth.vbJuly  '= 7
      retVal = "July"
   Case eVbMonth.vbAugust  '= 8
      retVal = "August"
   Case eVbMonth.vbSeptember  '= 9
      retVal = "September"
   Case eVbMonth.vbOctober  '= 10
      retVal = "October"
   Case eVbMonth.vbNovember  '= 11
      retVal = "November"
   Case eVbMonth.vbDecember  '= 12
      retVal = "December"
   Case Else
   End Select 'eVbMonth
   eVbMonthDesc = retVal
End Function 'eVbMonthDesc



Public Function eVbDayOccurrenceDesc( _
   Index As eVbDayOccurrence, _
   Optional ByVal Language As eLanguage = eLanguage.Korean) As String
   Dim retVal As String
   Select Case Index
   Case eVbDayOccurrence.vbFirst  '= 1
      Select Case Language
      Case eLanguage.Korean
         retVal = "Ã¹Â°"
      Case eLanguage.English
         retVal = "First"
      Case Else
         retVal = "First"
      End Select
   Case eVbDayOccurrence.vbSecond  '= 2
      Select Case Language
      Case eLanguage.Korean
         retVal = "µÑÂ°"
      Case eLanguage.English
         retVal = "Second"
      Case Else
         retVal = "Second"
      End Select
   Case eVbDayOccurrence.vbThird  '= 3
      Select Case Language
      Case eLanguage.Korean
         retVal = "¼¼Â°"
      Case eLanguage.English
         retVal = "Third"
      Case Else
         retVal = "Third"
      End Select
   Case eVbDayOccurrence.vbFourth  '= 4
      Select Case Language
      Case eLanguage.Korean
         retVal = "³×Â°"
      Case eLanguage.English
         retVal = "Fourth"
      Case Else
         retVal = "Fourth"
      End Select
   Case eVbDayOccurrence.vbFifth  '= 5
      Select Case Language
      Case eLanguage.Korean
         retVal = "´Ù¼¸Â°"
      Case eLanguage.English
         retVal = "Fifth"
      Case Else
         retVal = "Fifth"
      End Select
   Case eVbDayOccurrence.vbLast  '= 6
      Select Case Language
      Case eLanguage.Korean
         retVal = "¸¶Áö¸·"
      Case eLanguage.English
         retVal = "Last"
      Case Else
         retVal = "Last"
      End Select
   Case Else
   End Select 'eVbDayOccurrence
   eVbDayOccurrenceDesc = retVal
End Function 'eVbDayOccurrenceDesc


Private Function LocaleIDFromLangID( _
                           ByVal cnLanguage As eLanguage, _
                           Optional ByVal usSubLanguage As Integer = 1) As Long

   LocaleIDFromLangID = (usSubLanguage * 1024) Or CInt(cnLanguage)

End Function

Private Function MonthName( _
      Month As eMonth, _
      Optional ByVal cnLanguage As eLanguage = eLanguage.eLanguage_Default) As String

   Dim Buffer As String * 100
   Dim ret As Long
   Dim nullpos As Long

   If cnLanguage <> eLanguage.eLanguage_Default Then
      cnLanguage = LocaleIDFromLangID(cnLanguage)
   End If
   ret = apiGetLocaleInfo(cnLanguage, Month, Buffer, 99)
   nullpos& = VBA.InStr(1, Buffer, Chr$(0))
   If nullpos > 0 Then
      MonthName = VBA.Left$(Buffer, nullpos - 1)
   End If

End Function

Private Function WeekdayName( _
      Day As eWeekDay, _
      Optional ByVal cnLanguage As eLanguage = eLanguage.eLanguage_Default) As String

   Dim Buffer As String * 100
   Dim ret As Long
   Dim nullpos As Long

   If cnLanguage <> eLanguage.eLanguage_Default Then
      cnLanguage = LocaleIDFromLangID(cnLanguage)
   End If
   ret = apiGetLocaleInfo(cnLanguage, Day, Buffer, 99)
   nullpos& = VBA.InStr(1, Buffer, Chr$(0))
   If nullpos > 0 Then
      WeekdayName = VBA.Left$(Buffer, nullpos - 1)
   End If

End Function

Private Function TrimNull(StrIn As String) As String

'Trim Nulls
   Dim nul As Long

   'Truncate input string at first null.
   'If no nulls, perform ordinary Trim.
   nul = VBA.InStr(StrIn, vbNullChar)
   Select Case nul
   Case Is > 1
      TrimNull = VBA.Left$(StrIn, nul - 1)
   Case 1
      TrimNull = ""
   Case 0
      TrimNull = VBA.Trim$(StrIn)
   End Select

End Function


Public Function eLanguageDesc( _
   Index As eLanguage) As String
   Dim retVal As String
   Select Case Index
   Case eLanguage.eLanguage_Default  '= &H400&
      retVal = "eLanguage_Default"
   Case eLanguage.Neutral  '= &H0& '0
      retVal = "Neutral"
   Case eLanguage.Arabic  '= &H1& '1
      retVal = "Arabic"
   Case eLanguage.Bulgarian  '= &H2& '2
      retVal = "Bulgarian"
   Case eLanguage.Catalan  '= &H3& '3
      retVal = "Catalan"
   Case eLanguage.Chinese  '= &H4& '4
      retVal = "Chinese"
   Case eLanguage.Czech  '= &H5& '5
      retVal = "Czech"
   Case eLanguage.Danish  '= &H6& '6
      retVal = "Danish"
   Case eLanguage.German  '= &H7& '7
      retVal = "German"
   Case eLanguage.Greek  '= &H8& '8
      retVal = "Greek"
   Case eLanguage.English  '= &H9& '9
      retVal = "English"
   Case eLanguage.Spanish  '= &HA& '10
      retVal = "Spanish"
   Case eLanguage.Finnish  '= &HB& '11
      retVal = "Finnish"
   Case eLanguage.French  '= &HC& '12
      retVal = "French"
   Case eLanguage.Hebrew  '= &HD& '13
      retVal = "Hebrew"
   Case eLanguage.Hungarian  '= &HE& '14
      retVal = "Hungarian"
   Case eLanguage.Icelandic  '= &HF& '15
      retVal = "Icelandic"
   Case eLanguage.Italian  '= &H10& '16
      retVal = "Italian"
   Case eLanguage.Japanese  '= &H11& '17
      retVal = "Japanese"
   Case eLanguage.Korean  '= &H12& '18
      retVal = "Korean"
   Case eLanguage.Dutch  '= &H13& '19
      retVal = "Dutch"
   Case eLanguage.Norwegian  '= &H14& '20
      retVal = "Norwegian"
   Case eLanguage.Polish  '= &H15& '21
      retVal = "Polish"
   Case eLanguage.Portuguese  '= &H16& '22
      retVal = "Portuguese"
   Case eLanguage.Romanian  '= &H18& '24
      retVal = "Romanian"
   Case eLanguage.Russian  '= &H19& '25
      retVal = "Russian"
   Case eLanguage.Croatian  '= &H1A& '26
      retVal = "Croatian"
   Case eLanguage.Serbian  '= &H1A& '26
      retVal = "Serbian"
   Case eLanguage.Slovak  '= &H1B& '27
      retVal = "Slovak"
   Case eLanguage.Albanian  '= &H1C& '28
      retVal = "Albanian"
   Case eLanguage.Swedish  '= &H1D& '29
      retVal = "Swedish"
   Case eLanguage.Thai  '= &H1E& '30
      retVal = "Thai"
   Case eLanguage.Turkish  '= &H1F& '31
      retVal = "Turkish"
   Case eLanguage.Urdu  '= &H20& '32
      retVal = "Urdu"
   Case eLanguage.Indonesian  '= &H21& '33
      retVal = "Indonesian"
   Case eLanguage.Ukrainian  '= &H22& '34
      retVal = "Ukrainian"
   Case eLanguage.Belarusian  '= &H23& '35
      retVal = "Belarusian"
   Case eLanguage.Slovenian  '= &H24& '36
      retVal = "Slovenian"
   Case eLanguage.Estonian  '= &H25& '37
      retVal = "Estonian"
   Case eLanguage.Latvian  '= &H26& '38
      retVal = "Latvian"
   Case eLanguage.Lithuanian  '= &H27& '39
      retVal = "Lithuanian"
   Case eLanguage.Farsi  '= &H29& '41
      retVal = "Farsi"
   Case eLanguage.Armenian  '= &H2B& '43
      retVal = "Armenian"
   Case eLanguage.Azeri  '= &H2C& '44
      retVal = "Azeri"
   Case eLanguage.Basque  '= &H2D& '45
      retVal = "Basque"
   Case eLanguage.Macedonian  '= &H2F& '47
      retVal = "Macedonian"
   Case eLanguage.Afrikaans  '= &H36& '54
      retVal = "Afrikaans"
   Case eLanguage.Georgian  '= &H37& '55
      retVal = "Georgian"
   Case eLanguage.Faeroese  '= &H38& '56
      retVal = "Faeroese"
   Case eLanguage.Hindi  '= &H39& '57
      retVal = "Hindi"
   Case eLanguage.Malay  '= &H3E& '62
      retVal = "Malay"
   Case eLanguage.Kazak  '= &H3F& '63
      retVal = "Kazak"
   Case eLanguage.Swahili  '= &H41& '65
      retVal = "Swahili"
   Case eLanguage.Uzbek  '= &H43& '67
      retVal = "Uzbek"
   Case eLanguage.Tatar  '= &H44& '68
      retVal = "Tatar"
   Case eLanguage.Bengali  '= &H45& '69
      retVal = "Bengali"
   Case eLanguage.Punjabi  '= &H46& '70
      retVal = "Punjabi"
   Case eLanguage.Gujarati  '= &H47& '71
      retVal = "Gujarati"
   Case eLanguage.Oriya  '= &H48& '72
      retVal = "Oriya"
   Case eLanguage.Tamil  '= &H49& '73
      retVal = "Tamil"
   Case eLanguage.Telugu  '= &H4A& '74
      retVal = "Telugu"
   Case eLanguage.Kannada  '= &H4B& '75
      retVal = "Kannada"
   Case eLanguage.Malayalam  '= &H4C& '76
      retVal = "Malayalam"
   Case eLanguage.Assamese  '= &H4D& '77
      retVal = "Assamese"
   Case eLanguage.Marathi  '= &H4E& '78
      retVal = "Marathi"
   Case eLanguage.Sanskrit  '= &H4F& '79
      retVal = "Sanskrit"
   Case eLanguage.Konkani  '= &H57& '87
      retVal = "Konkani"
   Case eLanguage.Manipuri  '= &H58& '88
      retVal = "Manipuri"
   Case eLanguage.Sindhi  '= &H59& '89
      retVal = "Sindhi"
   Case eLanguage.Kashmiri  '= &H60& '96
      retVal = "Kashmiri"
   Case eLanguage.Nepali  '= &H61& '97
      retVal = "Nepali"

   Case eLanguage.eLanguage_MIN  '= 0
      retVal = "eLanguage_MIN"
   Case eLanguage.eLanguage_MAX  '= Nepali
      retVal = "eLanguage_MAX"
   Case Else
   End Select 'eLanguage
   eLanguageDesc = retVal
End Function 'eLanguageDesc

Private Function GetLocaleInfo( _
      InfoType As eLCInfo, _
      Optional ByVal cnLanauge As eLanguage = eLanguage.eLanguage_Default) As String
   
   Dim Buffer As String * 100
   Dim ret As Long
   Dim nullpos As Long
   
   If cnLanauge <> eLanguage.eLanguage_Default Then
      cnLanauge = LocaleIDFromLangID(cnLanauge)
   End If
   ret = apiGetLocaleInfo(cnLanauge, InfoType, Buffer, 99)
   nullpos& = VBA.InStr(1, Buffer, Chr$(0))
   If nullpos > 0 Then
      GetLocaleInfo = VBA.Left$(Buffer, nullpos - 1)
   End If
   Select Case InfoType
      Case eLCInfo.LongDateFormat
         GetLocaleInfo = VBA.Replace(GetLocaleInfo, "'", "")
   End Select
   

End Function


Private Function GetLongDate( _
   ByVal newDate As Date, _
   Optional ByVal cnLanguage As eLanguage = eLanguage.eLanguage_Default) As String
   
   'Debug.Print GetLocaleInfo(LongDateFormat, cnLanguage)
   GetLongDate = VBA.Format$(newDate, GetLocaleInfo(LongDateFormat, cnLanguage))
   GetLongDate = VBA.Replace(GetLongDate, MonthName(JanuaryLong + VBA.Month(newDate) - 1, English), _
                              MonthName(JanuaryLong + VBA.Month(newDate) - 1, cnLanguage))
   GetLongDate = VBA.Replace(GetLongDate, WeekdayName(eWeekDay.MondayLong + VBA.Weekday(newDate, VbDayOfWeek.vbMonday) - 1, English), _
                              WeekdayName(eWeekDay.MondayLong + VBA.Weekday(newDate, VbDayOfWeek.vbMonday) - 1, cnLanguage))
End Function

Private Function GetShortDate( _
   ByVal newDate As Date, _
   Optional ByVal cnLanguage As eLanguage = eLanguage.eLanguage_Default) As String
   
   GetShortDate = VBA.Replace(VBA.Format$(newDate, GetLocaleInfo(ShortDateFormat, cnLanguage)), _
                              GetLocaleInfo(DateSeparator, eLanguage_Default), GetLocaleInfo(DateSeparator, cnLanguage))
                              
   
End Function

Private Function GetLongTime( _
   ByVal newDate As Date, _
   Optional ByVal cnLanguage As eLanguage = eLanguage.eLanguage_Default) As String
   
   GetLongTime = ReplaceTimeMarker(VBA.Format$(newDate, GetLongTimeFormatTemp(cnLanguage)), cnLanguage)
   
End Function

Private Function GetShortTime( _
   ByVal newDate As Date, _
   Optional ByVal cnLanguage As eLanguage = eLanguage.eLanguage_Default) As String
   
   GetShortTime = VBA.Format$(newDate, "Short Time")
   
End Function

Private Function GetLongTimeFormatTemp( _
   Optional ByVal cnLanguage As eLanguage = eLanguage.eLanguage_Default) As String
   
   GetLongTimeFormatTemp = VBA.Replace(GetLocaleInfo(TimeFormat, cnLanguage), "tt", "AM/PM", 1, 1)
   'Debug.Print VBA.Replace(GetLocaleInfo(TimeFormat, cnLanguage), "tt", "AM/PM", 1, 1)
End Function

Private Function GetFileDateTime( _
      ByVal newDate As Date, _
      Optional ByVal cnLanguage As eLanguage = eLanguage.eLanguage_Default) As String
   
   GetFileDateTime = GetShortDate(newDate, cnLanguage) & " " & _
                                 ReplaceTimeMarker(VBA.Format$(newDate, GetFileDateTimeFormatTemp(cnLanguage)), cnLanguage)
   
End Function

Private Function GetFileDateTimeFormatTemp( _
   Optional ByVal cnLanguage As eLanguage = eLanguage.eLanguage_Default) As String
   
   GetFileDateTimeFormatTemp = VBA.Replace(GetLongTimeFormatTemp(cnLanguage), ":ss", "")
   
End Function

Private Function ReplaceTimeMarker( _
   ByVal strLongTime As String, _
   Optional ByVal cnLanguage As eLanguage = eLanguage.eLanguage_Default) As String
   
   Select Case cnLanguage
   Case eLanguage.Korean
      strLongTime = VBA.Replace(VBA.Replace(strLongTime, "AM", "¿ÀÀü", 1, 1), "PM", "¿ÀÈÄ", 1, 1)
   End Select
   ReplaceTimeMarker = strLongTime
End Function


Public Function IsIneLanguage(Index As eLanguage) As Boolean
   IsIneLanguage = (Len(eLanguageDesc(Index)) > 0)
End Function

