Attribute VB_Name = "Utils"
Attribute VB_HelpID = 1635
'Утилиты даты и печати числовых строк

Option Explicit

Public Enum enumGender
 MALE = 0
 FEMALE = 1
End Enum


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Attribute ShellExecute.VB_HelpID = 1670

' число дней в месяце
'Parameters:
'[IN]   M , тип параметра: Integer - месяц,
'[IN]   Y , тип параметра: Integer  - год
'Returns:
'  значение типа Integer - число дней
'See Also:
'  DayOfWeek
'  Kop2str
'  MonthName
'  MonthNameI
'  parseNumber
'  ShellExecute
'Example:
' dim variable as Integer
' variable = me.DaysInMonth(...параметры...)
Public Function DaysInMonth(ByVal M As Integer, ByVal Y As Integer) As Integer
Attribute DaysInMonth.VB_HelpID = 1645
On Error GoTo DaysInMonthErr
    If M = 4 Or M = 6 Or M = 7 Or M = 9 Or M = 11 Then
        DaysInMonth = 30
    ElseIf M = 2 Then
        If Y Mod 4 Then
            DaysInMonth = 28
        ElseIf Y Mod 100 Then
            DaysInMonth = 29
        ElseIf Y Mod 400 Then
            DaysInMonth = 28
        Else
            DaysInMonth = 29
        End If
    Else
        DaysInMonth = 31
    End If
    Exit Function
DaysInMonthErr:
End Function


'название месяца именительный падеж
'Parameters:
'[IN]   id , тип параметра: Integer  - номер месяца
'Returns:
'  значение типа String - название месяца
'See Also:
'  DayOfWeek
'  DaysInMonth
'  Kop2str
'  MonthName
'  parseNumber
'  ShellExecute
'Example:
' dim variable as String
' variable = me.MonthNameI(...параметры...)
Public Function MonthNameI(ByVal id As Integer) As String
Attribute MonthNameI.VB_HelpID = 1660
On Error GoTo MonthNameIErr
    Select Case id
    Case 1:
            MonthNameI = "январь"
    Case 2:
            MonthNameI = "февраль"
    Case 3:
            MonthNameI = "март"
    Case 4:
            MonthNameI = "апрель"
    Case 5:
            MonthNameI = "май"
    Case 6:
            MonthNameI = "июнь"
    Case 7:
            MonthNameI = "июль"
    Case 8:
            MonthNameI = "август"
    Case 9:
            MonthNameI = "сентябрь"
    Case 10:
            MonthNameI = "октябрь"
    Case 11:
            MonthNameI = "ноябрь"
    Case 12:
            MonthNameI = "декабрь"
    Case Else
            MonthNameI = "???"
    End Select
    Exit Function
MonthNameIErr:

End Function

'Название месяца для даты
'Parameters:
'[IN]   id , тип параметра: Integer  - номер месяца
'Returns:
'  значение типа String - название
'See Also:
'  DayOfWeek
'  DaysInMonth
'  Kop2str
'  MonthNameI
'  parseNumber
'  ShellExecute
'Example:
' dim variable as String
' variable = me.MonthName(...параметры...)
Public Function MonthName(ByVal id As Integer) As String
Attribute MonthName.VB_HelpID = 1655
On Error GoTo MonthNameErr
    Select Case id
    Case 1:
            MonthName = "января"
    Case 2:
            MonthName = "февраля"
    Case 3:
            MonthName = "марта"
    Case 4:
            MonthName = "апреля"
    Case 5:
            MonthName = "мая"
    Case 6:
            MonthName = "июня"
    Case 7:
            MonthName = "июля"
    Case 8:
            MonthName = "августа"
    Case 9:
            MonthName = "сентября"
    Case 10:
            MonthName = "октября"
    Case 11:
            MonthName = "ноября"
    Case 12:
            MonthName = "декабря"
    Case Else
            MonthName = "???"
    End Select
    Exit Function
MonthNameErr:

End Function


'количество сотен в текст
Private Function hund2str(ByVal h As Integer) As String

    Select Case h
        Case 0: hund2str = ""
        Case 1: hund2str = "сто"
        Case 2: hund2str = "двести"
        Case 3: hund2str = "триста"
        Case 4: hund2str = "четыреста"
        Case 5: hund2str = "пятьсот"
        Case 6: hund2str = "шестьсот"
        Case 7: hund2str = "семьсот"
        Case 8: hund2str = "восемьсот"
        Case 9: hund2str = "девятьсот"
        Case Else: hund2str = "???"
    End Select

End Function

' количество десятков  в текст
Private Function dec2str(d As Integer) As String
    Select Case d
        Case 0: dec2str = ""
        Case 1: dec2str = "десять"
        Case 2: dec2str = "двадцать"
        Case 3: dec2str = "тридцать"
        Case 4: dec2str = "сорок"
        Case 5: dec2str = "пятьдесят"
        Case 6: dec2str = "шестьдесят"
        Case 7: dec2str = "семьдесят"
        Case 8: dec2str = "восемьдесят"
        Case 9: dec2str = "девяносто"
        Case Else: dec2str = "???"
    End Select
End Function


' 10 - 19 в текст
Private Function decdig2str(ByVal n As Integer) As String
    Select Case n
        Case 10: decdig2str = "десять"
        Case 11: decdig2str = "одиннадцать"
        Case 12: decdig2str = "двенадцать"
        Case 13: decdig2str = "тринадцать"
        Case 14: decdig2str = "четырнадцать"
        Case 15: decdig2str = "пятнадцать"
        Case 16: decdig2str = "шестнадцать"
        Case 17: decdig2str = "семнадцать"
        Case 18: decdig2str = "восемнадцать"
        Case 19: decdig2str = "девятнадцать"
        Case Else: decdig2str = "???"
    End Select
End Function


'1-9 в текст с учетом рода
Private Function dig2str(ByVal d As Integer, ByVal sex As Integer)

If sex = MALE Then
    Select Case d
        Case 0: dig2str = ""
        Case 1: dig2str = "один"
        Case 2: dig2str = "два"
        Case 3: dig2str = "три"
        Case 4: dig2str = "четыре"
        Case 5: dig2str = "пять"
        Case 6: dig2str = "шесть"
        Case 7: dig2str = "семь"
        Case 8: dig2str = "восемь"
        Case 9: dig2str = "девять"
    End Select
Else
    Select Case d
        Case 0: dig2str = ""
        Case 1: dig2str = "одна"
        Case 2: dig2str = "две"
        Case 3: dig2str = "три"
        Case 4: dig2str = "четыре"
        Case 5: dig2str = "пять"
        Case 6: dig2str = "шесть"
        Case 7: dig2str = "семь"
        Case 8: dig2str = "восемь"
        Case 9: dig2str = "девять"
    End Select
End If
End Function

' приведение  названия валюты к правильному склонению
Private Function male2str(ByVal d As Currency, ByVal root As String) As String
 Dim tmp As String, buf As String, s As String
 Dim mode As Integer, n As Integer
 Dim s2 As String, d1 As Long
 's2 = Format(d, "0000000000000000.00")
 s2 = Format(d, "0000000000000000.00")
 'n = CLng(Right(s2, 2))
 n = CLng(Mid(s2, 15, 2))
 'n = d Mod 100
 buf = UCase(root)
 If (Left$(buf, 4) = "РУБЛ") Then
    If (n >= 20) Then n = n Mod 10
    If (n = 1) Then
          tmp = "ь"
    ElseIf (n > 1 And n < 5) Then
          tmp = "я"
    Else
          tmp = "ей"
    End If
    s = LCase(Left$(root, 4))
 Else
     If (n >= 20) Then n = n Mod 10
     If (n = 1) Then
          tmp = ""
      ElseIf (n < 5 And n > 1) Then
          tmp = "а"
      Else
          tmp = "ов"
      End If
      s = root
  End If
 male2str = s + tmp
End Function

' Раскладывает то что меньше тысячи
Private Function num2str(ByVal numb As Integer, ByVal gender As Integer) As String
     Dim out As String, tmp As String, dest As String
     Dim hund As Integer, dec As Integer, dig As Integer
     Dim num As String
     num = Format(numb, "000")
     dest = " "
     hund = MyRound(Left(num, 1))
     If (hund >= 1) Then
      tmp = hund2str(hund)
      dest = dest + tmp
     End If
     dec = MyRound(Mid(num, 2, 1))
     If (dec >= 1) Then
         If (dec = 1) Then
              tmp = decdig2str(MyRound(Right(num, 2)))
              dest = dest + " " + tmp
              num2str = dest
              Exit Function
         Else
             tmp = dec2str(dec)
             dest = dest + " " + tmp
         End If
      End If
      tmp = dig2str(MyRound(Right(num, 1)), gender)
      dest = dest + " " + tmp
      num2str = dest
End Function

'число в текст с учетом рода
'Parameters:
'[IN][OUT]  numberof , тип параметра: Currency - число,
'[IN]   gender , тип параметра: enumGender  - род
'Returns:
'  значение типа String - текстовое представление числа
'See Also:
'  DayOfWeek
'  DaysInMonth
'  Kop2str
'  MonthName
'  MonthNameI
'  ShellExecute
'Example:
' dim variable as String
' variable = me.parseNumber(...параметры...)
Public Function parseNumber(numberof As Currency, ByVal gender As enumGender) As String
Attribute parseNumber.VB_HelpID = 1665
     Dim trl As Currency
     Dim numb As String
     Dim rems As Integer, tail As Currency
     Dim Name As String, tmp As String, dest As String
     Dim i As Integer
     
     If numberof = 0 Then
        parseNumber = "Ноль"
        Exit Function
     End If
     numb = CStr(IIf(numberof < 0, -numberof, numberof))
     For i = 1 To Len(numb)
        If Mid(numb, i, 1) = "." Then Exit For
     Next
     numb = Format(CCur(Left(numb, i)), "000000000000000")
     dest = ""
     rems = MyRound(Mid(numb, 1, 3))
     If (rems >= 1) Then
         tmp = num2str(rems, MALE)
         Name = male2str(rems, " триллион")
         dest = dest + tmp + Name
     End If
     rems = MyRound(Mid(numb, 4, 3))
     If (rems >= 1) Then
         tmp = num2str(rems, MALE)
         Name = male2str(rems, " миллиард")
         dest = dest + tmp + Name
     End If
     rems = MyRound(Mid(numb, 7, 3))
     If (rems >= 1) Then
         tmp = num2str(rems, MALE)
         Name = male2str(rems, " миллион")
         dest = dest + tmp + Name
     End If
     rems = MyRound(Mid(numb, 10, 3))
     If (rems >= 1) Then
        tmp = num2str(rems, FEMALE)
        Name = Thou2str(rems Mod 100)
        dest = dest + tmp + Name
     End If
      rems = MyRound(Right(numb, 3))
      If (rems >= 1) Then
        tmp = num2str(rems, gender)
        dest = dest + tmp
      End If
      dest = Trim(dest)
      parseNumber = UCase(Left(dest, 1)) + LCase(Mid(dest, 2, Len(dest) - 1))
End Function

' склонение тысяч
Private Function Thou2str(ByVal n As Integer) As String
    Dim n1 As Integer
    n1 = n
    If (n1 >= 10 And n1 < 20) Then
        n1 = 0
    ElseIf (n1 >= 20) Then
        n1 = n1 Mod 10
    End If
    
    Select Case n1
      Case 1: Thou2str = " тысяча"
      Case 2 To 4: Thou2str = " тысячи"
      Case Else: Thou2str = " тысяч"
    End Select
End Function

' количесвто копеек в текст
'Parameters:
'[IN]   n , тип параметра: Integer  - числ копеек
'Returns:
'  значение типа String - текстовое представление
'See Also:
'  DayOfWeek
'  DaysInMonth
'  MonthName
'  MonthNameI
'  parseNumber
'  ShellExecute
'Example:
' dim variable as String
' variable = me.Kop2str(...параметры...)
Public Function Kop2str(ByVal n As Integer) As String
Attribute Kop2str.VB_HelpID = 1650
    Dim n1 As Integer
    n1 = n
    If (n1 >= 10 And n1 < 20) Then
        n1 = 0
    ElseIf (n1 >= 20) Then
        n1 = n1 Mod 10
    End If
    
    Select Case n1
      Case 1: Kop2str = " копейка"
      Case 2 To 4: Kop2str = " копейки"
      Case Else: Kop2str = " копеек"
    End Select
End Function

'вычисление дня недели из даты
'Parameters:
'[IN]   d , тип параметра: Date  - дата
'Returns:
'  значение типа Integer - гномер дня в неделе
'See Also:
'  DaysInMonth
'  Kop2str
'  MonthName
'  MonthNameI
'  parseNumber
'  ShellExecute
'Example:
' dim variable as Integer
' variable = me.DayOfWeek(...параметры...)
Public Function DayOfWeek(ByVal d As Date) As Integer
Attribute DayOfWeek.VB_HelpID = 1640

    Dim c4 As Long, century As Long, yr As Long, dw As Long, y2 As Long, m2 As Long, d2 As Long
    y2 = Year(d)
    m2 = Month(d)
    d2 = Day(d)

    If m2 < 3 Then
        m2 = m2 + 10
        y2 = y2 - 1
    Else
        m2 = m2 - 2
    End If

    century = y2 \ 100
    
    
    yr = y2 Mod 100
    
    dw = ((26 * m2 - 2) \ 10 + d2 + yr + (yr \ 4) + (century \ 4) - (2 * century)) Mod 7

    If dw < 0 Then dw = dw + 7

    If dw = 0 Then dw = 7

    DayOfWeek = dw
End Function
