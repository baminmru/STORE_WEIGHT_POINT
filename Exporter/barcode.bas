Attribute VB_Name = "barcode"
Option Explicit



Public Function code128(chaine As String) As String
  
  'Parameters : a string
  'Return : * a string which give the bar code when it is dispayed with CODE128.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i As Long, checksum As Long, mini As Long, dummy As Long, tableB As Boolean
  code128 = ""
  If Len(chaine) > 0 Then
 
  'Check for valid characters
    For i = 1 To Len(chaine)
      Select Case Asc(Mid(chaine, i, 1))
      Case 32 To 126, 198
      Case Else
        i = 0
        Exit For
      End Select
    Next
    'Calculation of the code string with optimized use of tables B and C
    code128 = ""
    tableB = True
    If i > 0 Then
      i = 1 ' i become the string index
      Do While i <= Len(chaine)
        If tableB Then
          ' See if interesting to switch to table C
          ' yes for 4 digits at start or end, else if 6 digits
          mini = IIf(i = 1 Or i + 3 = Len(chaine), 4, 6)
          GoSub testnum
          If mini < 0 Then ' Choice of table C
            If i = 1 Then ' Starting with table C
              code128 = Chr(205)
            Else 'Switch to table C
              code128 = code128 & Chr(199)
            End If
            tableB = False
          Else
            If i = 1 Then code128 = Chr(204) ' Starting with table B
          End If
        End If
        If Not tableB Then
          ' We are on table C, try to process 2 digits
          mini = 2
          GoSub testnum
          If mini < 0 Then ' OK for 2 digits, process it
            dummy = MyRound(Mid(chaine, i, 2))
            dummy = IIf(dummy < 95, dummy + 32, dummy + 100)
            code128 = code128 & Chr(dummy)
            i = i + 2
          Else ' We haven't 2 digits, switch to table B
            code128 = code128 & Chr(200)
            tableB = True
          End If
        End If
        If tableB Then
          ' Process 1 digit with table B
          code128 = code128 & Mid(chaine, i, 1)
          i = i + 1
        End If
      Loop
      ' Calculation of the checksum
      For i = 1 To Len(code128)
        dummy = Asc(Mid(code128, i, 1))
        dummy = IIf(dummy < 127, dummy - 32, dummy - 100)
        If i = 1 Then checksum = dummy
        checksum = (checksum + (i - 1) * dummy) Mod 103
      Next
      ' Calculation of the checksum ASCII code
      checksum = IIf(checksum < 95, checksum + 32, checksum + 100)
      ' Add the checksum and the STOP
      code128 = code128 & Chr(checksum) & Chr(206)
    End If
  End If
  Exit Function
testnum:
  
  'if the mini characters from i are numeric, then mini=0
  mini = mini - 1
  If i + mini <= Len(chaine) Then
    Do While mini >= 0
      If Asc(Mid(chaine, i + mini, 1)) < 48 Or Asc(Mid(chaine, i + mini, 1)) > 57 Then Exit Do
      mini = mini - 1
    Loop
  End If
Return
End Function

