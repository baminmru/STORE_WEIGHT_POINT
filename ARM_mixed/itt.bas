Attribute VB_Name = "itt"
Attribute VB_HelpID = 795
Option Explicit
'Функции для обмена с CORE


Public log As New MTZUtil.FILELOGGER


'Выполнить запрос
'Parameters:
'[IN]   sqlstring , тип параметра: String - строка запроса,
'[IN][OUT]   conn , тип параметра: ADODB.Connection  - коннект к базе CORE
'Returns:
'  объект класса ADODB.Recordset
'  ,или Nothing
'See Also:
'  CleanRCVAtCore
'  FindPoddon
'  GetBRIEFFromXMLField
'  MyRound
'  MyRound2
'  PrintSticker
'  SaveRCVRowToCore
'  SaveShipRowToCore
'  UpdateMyPalet
'Example:
' dim variable as ADODB.Recordset
' Set variable = me.ConnExec(...параметры...)
Public Function ConnExec(ByVal sqlstring As String, ByRef conn As ADODB.Connection) As ADODB.Recordset
Attribute ConnExec.VB_HelpID = 805
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.CursorType = adOpenStatic 'adOpenForwardOnly
  rs.LockType = adLockReadOnly
  rs.CursorLocation = adUseClient
  Call rs.open(sqlstring, conn)
  rs.ActiveConnection = Nothing
  Set ConnExec = rs
  Set rs = Nothing
End Function

'Округление
'Parameters:
'[IN]   s , тип параметра: String  -  число
'Returns:
'  значение типа Double
'See Also:
'  CleanRCVAtCore
'  ConnExec
'  FindPoddon
'  GetBRIEFFromXMLField
'  MyRound2
'  PrintSticker
'  SaveRCVRowToCore
'  SaveShipRowToCore
'  UpdateMyPalet
'Example:
' dim variable as Double
' variable = me.MyRound(...параметры...)
Public Function MyRound(ByVal s As String) As Double
Attribute MyRound.VB_HelpID = 820
  Dim e As String, out As Double
  e = Replace(s, ",", ".")
  out = Val("0" & e)
  out = Round(out + 0.001, 2)
  MyRound = out

End Function

'округление с заменой запятой на точку
'Parameters:
'[IN]   s , тип параметра: String  - число
'Returns:
'  значение типа String - строка для подстановки в запрос
'See Also:
'  CleanRCVAtCore
'  ConnExec
'  FindPoddon
'  GetBRIEFFromXMLField
'  MyRound
'  PrintSticker
'  SaveRCVRowToCore
'  SaveShipRowToCore
'  UpdateMyPalet
'Example:
' dim variable as String
' variable = me.MyRound2(...параметры...)
Public Function MyRound2(ByVal s As String) As String
Attribute MyRound2.VB_HelpID = 825
  Dim e As String, out As Double
  e = Replace(s, ",", ".")
  out = Val("0" & e)
  out = Round(out + 0.001, 2)
  e = out
  MyRound2 = Replace(e, ",", ".")

End Function

'поиск поддона  в базе данных весового комплекса по номеру
'Parameters:
'[IN]   TheNumber , тип параметра: String  - номер поддона
'Returns:
'  объект любого класса Visual Basic
'  ,или Nothing
'See Also:
'  CleanRCVAtCore
'  ConnExec
'  GetBRIEFFromXMLField
'  MyRound
'  MyRound2
'  PrintSticker
'  SaveRCVRowToCore
'  SaveShipRowToCore
'  UpdateMyPalet
'Example:
' dim variable as Object
' Set variable = me.FindPoddon(...параметры...)
Public Function FindPoddon(ByVal TheNumber As String) As Object
Attribute FindPoddon.VB_HelpID = 810
  Dim rs As ADODB.Recordset
  Dim Obj As Object
  Dim id As String
  Set rs = Session.GetData("select * from ittpl_def where code ='" & TheNumber & "'")
  If rs Is Nothing Then Exit Function
  If rs.EOF Then Exit Function
  id = rs!InstanceID
  Set Obj = Manager.GetInstanceObject(id)
  If Obj Is Nothing Then Exit Function
  Set FindPoddon = Obj.ITTPL_DEF.Item(1)
End Function


'получить номер заказа из видимого отображения
Private Function ORDERID(ByVal s As String) As String
Dim out() As String
On Error Resume Next
out = Split(s, " от")

ORDERID = Replace(out(0), " ", "")

End Function

'печать стикера
'Parameters:
'[IN]   poddon , тип параметра: ITTPL_DEF  - описание поддона,
'[IN][OUT]   Optional CaliberWeight , тип параметра: Double = -1  - калиброванный вес
'See Also:
'  CleanRCVAtCore
'  ConnExec
'  FindPoddon
'  GetBRIEFFromXMLField
'  MyRound
'  MyRound2
'  SaveRCVRowToCore
'  SaveShipRowToCore
'  UpdateMyPalet
'Example:
'  call me.PrintSticker(...параметры...)
Public Sub PrintSticker(ByVal poddon As ITTPL_DEF, Optional CaliberWeight As Double = -1)
Attribute PrintSticker.VB_HelpID = 830



If GetSetting("RBH", "ITTSETTINGS", "PSTICKER", 1) Then
  
  If MsgBox("Напечатать стикер на поддон?", vbYesNo) = vbYes Then
  
    Dim strs As ADODB.Recordset
    Dim itemrs As ADODB.Recordset
    Dim locrs As ADODB.Recordset
    Dim ordrs As ADODB.Recordset
    
    Dim conn As ADODB.Connection
    Set conn = GetCoreConn
    If conn.State <> adStateOpen Then
      conn.open
    End If
  
    Set strs = ConnExec("select * from STOCK where PALLET_STATUS is null and  PALLET_ID=" & poddon.CorePalette_ID, conn)
    If strs.EOF Then
      MsgBox "Поддон не принят на склад"
      Exit Sub
    End If
    Set locrs = ConnExec("select * from location where id=" & strs!location_id, conn)
    Set itemrs = ConnExec("select * from item where id=" & strs!item_id, conn)
    Set ordrs = ConnExec("select Name from partner join receiving_order on partner.id = receiving_order.partner_id where receiving_order.number='" & ORDERID(strs!ord_num) & "'", conn)
      
    Dim X As Printer
    For Each X In Printers
    If X.DeviceName = GetSetting("RBH", "ITTSETTINGS", "DOCPRN") Then
    
    Set Printer = X
    
    
    
    Printer.Font = "Arial CYR"
    Printer.FontSize = 48
    
    Printer.FontBold = False
    Printer.Print "Паллет №";
    Printer.FontBold = True
    Printer.Print poddon.code & "  ";
    Printer.Font = "Code 128"
    Printer.FontBold = False
    Printer.FontSize = 48
    Printer.Print code128(poddon.code)
    Printer.Font = "Arial CYR"
    Printer.FontSize = 32
    Printer.FontBold = False
    Printer.Print "Тип паллеты:";
    Printer.FontSize = 32
    Printer.FontBold = True
    Printer.Print poddon.Pltype.brief
        
    Printer.Font = "Arial CYR"
    Printer.FontSize = 48
    Printer.Print "Код: ";
    Printer.FontBold = True
    Printer.Print itemrs!code & "";
    
    Printer.Font = "Code 128"
    
    Printer.FontBold = False
    Printer.FontSize = 48
    Printer.Print code128(itemrs!code & "")
    
    Printer.Font = "Arial CYR"
    Printer.FontSize = 32
    
    Printer.FontBold = False
    Printer.Print "Поклажедатель: ";
    Printer.FontBold = True
    Printer.Print ordrs!Name
    
    Printer.FontBold = False
    Printer.Print "Заказ: ";
    Printer.FontBold = True
    Printer.Print strs!ord_num
    
    Printer.Font = "Arial CYR"
    Printer.FontSize = 28
    
    Printer.FontBold = False
    Printer.Print "Товар: ";
    Printer.FontBold = True
    
    Printer.Print Left(itemrs!Description & "", 30)
    If Len(itemrs!Description & "") > 30 Then
      Printer.Print Mid(itemrs!Description & "", 31, 36)
    End If
    If Len(itemrs!Description & "") > 30 + 36 Then
      Printer.Print Mid(itemrs!Description & "", 31 + 36)
    End If
    
    
       
    Printer.FontBold = False
    Printer.Print "Страна производитель: ";
    Printer.FontBold = True
    Printer.Print strs!custom_field6 & ""
      
    Printer.FontBold = False
    Printer.Print "Производитель: ";
    Printer.FontBold = True
    Printer.Print strs!custom_field4 & ""
    
    Printer.FontBold = False
    Printer.Print "Бойня: ";
    Printer.FontBold = True
    Printer.Print strs!custom_field11 & ""
  
    Printer.FontBold = False
    Printer.Print "Партия: ";
    Printer.FontBold = True
    Printer.Print strs!LOT_SN & ""
    
    Printer.FontBold = False
    Printer.Print "Вес груза НЕТТО (КГ.) : ";
    Printer.FontBold = True
    Printer.Print MyRound(strs!QTY_ON_HAND)
      
    Printer.FontBold = False
    Printer.Print "Вес груза Брутто (КГ.) : ";
    Printer.FontBold = True
    Printer.Print MyRound(strs!QTY_ON_HAND) + MyRound(strs!custom_field3) * MyRound(strs!custom_field1)
      
    Printer.FontBold = False
    Printer.Print "Вес поддона с грузом (КГ.) : ";
    Printer.FontBold = True
    Printer.Print MyRound(strs!QTY_ON_HAND) + MyRound(strs!custom_field3) * MyRound(strs!custom_field1) + poddon.Weight
    
    Printer.FontBold = False
    Printer.Print "Вес упаковки (КГ.) : ";
    Printer.FontBold = True
    Printer.Print MyRound(strs!custom_field3) * MyRound(strs!custom_field1)
    
    
    Printer.FontBold = False
    Printer.Print "Дата выпуска: ";
    Printer.FontBold = True
    Printer.Print strs!custom_field5
    
    Printer.FontBold = False
    Printer.Print "Cрок годности: ";
    Printer.FontBold = True
    Printer.Print strs!exp_date
    
    
    If strs!custom_field2 & "" = "1" Then
      Printer.FontBold = False
      Printer.Print "Калиброванный товар"
    End If
    
    If CaliberWeight > 0 Then
      Printer.Print "Вес одного короба НЕТТО (КГ.): ";
      Printer.FontBold = True
      Printer.Print Round(CaliberWeight + 0.001, 2)
    End If
    
    
    
    Printer.FontBold = False
    Printer.Print "Количество коробов: ";
    Printer.FontBold = True
    Printer.Print strs!custom_field1
    'Printer.EndDoc
    
    If GetSetting("RBH", "ITTSETTINGS", "PCELL", 0) = 1 Then

      Printer.Font = "Arial CYR"
          Printer.Font = "Arial CYR"
          Printer.FontSize = 32
          
          If strs!status = 101 Then
            Printer.Print "БРАК                ";
          End If
      
          Printer.FontBold = False
          Printer.Print "Ячейка .№"
          Printer.Print " ";
          Printer.FontSize = 80
          Printer.FontBold = True
          Printer.Print locrs!code
    Else
          Printer.Font = "Arial CYR"
          Printer.FontSize = 32
          
          If strs!status = 101 Then
            Printer.Print "БРАК "
          End If
    
    End If
    
    
    
    Printer.EndDoc
   Exit For
  End If
 Next
  End If
  End If
bye2:
  
  Exit Sub
  
bye:
  If err.Number <> 0 Then
     MsgBox err.Description, , "Печать документов на поддон"
  End If
End Sub

' обновление паллеты в CORE. Паллета для только что принятого заказа но не провведенного в CORE
'Parameters:
'[IN]   Pallet , тип параметра: ITTIN_PALET  - паллета
'Returns:
' Boolean, семантика результата:
'   true  - удачно
'   false - нет
'See Also:
'  CleanRCVAtCore
'  ConnExec
'  FindPoddon
'  GetBRIEFFromXMLField
'  MyRound
'  MyRound2
'  PrintSticker
'  SaveRCVRowToCore
'  SaveShipRowToCore
'Example:
' dim variable as Boolean
' variable = me.UpdateMyPalet(...параметры...)
Public Function UpdateMyPalet(ByVal Pallet As ITTIN_PALET) As Boolean
Attribute UpdateMyPalet.VB_HelpID = 845

  Dim conn As ADODB.Connection
  Set conn = GetCoreConn
  If conn.State <> adStateOpen Then
      conn.open
  End If
  
  Dim poddon As ITTPL_DEF
  Dim good_id As String, ordid As String
  Dim qry As ITTIN.Application
  Dim qline As ITTIN_QLINE
  Dim cntline As ITTIN_PALET
  Set qry = Pallet.Application
  Set qline = Pallet.Parent.Parent
  
  good_id = Manager.GetIDFromXMLField(qline.good_id)
  ordid = Manager.GetIDFromXMLField(qry.ITTIN_DEF.Item(1).QryCode)
  
  
  Dim i As Long, s As String
  Dim curval As Double, prev As Double, cnt As Long, fpw As Double
  curval = 0
  cnt = 0
  fpw = 0
  For i = 1 To qline.ITTIN_PALET.Count
    With qline.ITTIN_PALET.Item(i)
      .FullPackageWeight = .PackageWeight * .CaliberQuantity
      fpw = fpw + .FullPackageWeight
      curval = curval + .GoodWithPaletWeight - .FullPackageWeight - .PalWeight
      cnt = cnt + .CaliberQuantity
      .save
    End With
  Next
  
  qline.CurValue = curval
  qline.FullPackageWeight = fpw
  qline.save
  
  Set poddon = Pallet.TheNumber
  
  
  poddon.CaliberQuantity = Pallet.CaliberQuantity
  poddon.PackageWeight = Pallet.PackageWeight * Pallet.CaliberQuantity
  poddon.save
  
  Dim RollbackMe As Boolean
  
  RollbackMe = False
  conn.BeginTrans
  err.Clear
  
  s = ""
  s = s & " Update receiving_history"
  s = s & " set QTY_REC=" & MyRound2(Pallet.GoodWithPaletWeight - Pallet.FullPackageWeight - Pallet.PalWeight) & ", custom_field1='" & Pallet.CaliberQuantity & "' "
  s = s & " where pallet ='" & poddon.TheNumber & "' and order_id=" & ordid
  
  conn.Execute s
    If err.Number <> 0 Then
    ''' MsgBox err.Description
    RollbackMe = True
  End If
  
  s = ""
  s = s & " Update receiving_line"
  s = s & " Set QTY_PREV_REC = " & MyRound2(qline.CurValue) & " "
  s = s & " , QTY_ALT_PREV_REC = " & cnt & " "
  s = s & " Where order_id = " & ordid & " and item_id=" & good_id
  
  conn.Execute s
    If err.Number <> 0 Then
    ''' MsgBox err.Description
    RollbackMe = True
  End If
  
  s = ""
  s = s & " Update stock"
  s = s & " set QTY_ON_HAND=" & MyRound2(Pallet.GoodWithPaletWeight - Pallet.FullPackageWeight - poddon.Weight) & ", custom_field1='" & Pallet.CaliberQuantity & "'"
  s = s & " Where pallet_id = " & poddon.CorePalette_ID & " And pallet_status Is Null"
  
  
  conn.Execute s
    If err.Number <> 0 Then
    ''' MsgBox err.Description
    RollbackMe = True
  End If
  
  If err.Number <> 0 Then
    ''' MsgBox err.Description
    RollbackMe = True
  End If

 If RollbackMe Then
    conn.RollbackTrans
  Else
    conn.CommitTrans
    UpdateMyPalet = True
  End If

  
End Function

' Получить отображение из блока SQL Reference
'Parameters:
'[IN]   XML , тип параметра: String  - SQL Reference блок
'Returns:
'  значение типа String
'See Also:
'  CleanRCVAtCore
'  ConnExec
'  FindPoddon
'  MyRound
'  MyRound2
'  PrintSticker
'  SaveRCVRowToCore
'  SaveShipRowToCore
'  UpdateMyPalet
'Example:
' dim variable as String
' variable = me.GetBRIEFFromXMLField(...параметры...)
Public Function GetBRIEFFromXMLField(ByVal XML As String) As String
Attribute GetBRIEFFromXMLField.VB_HelpID = 815
  Dim iFrom As Long
  Dim iTo As Long
  Dim m_ID As String
  iFrom = InStr(1, XML, "<Brief>")
  If (iFrom > 0) Then
      iTo = InStr(iFrom, XML, "</Brief>")
      m_ID = Mid(XML, iFrom + 7, iTo - iFrom - 7)
  End If
  GetBRIEFFromXMLField = m_ID
End Function

'Удалить данные о приемке из CORE
'Parameters:
'[IN]   Item , тип параметра: ITTIN.Application  - заказ а приемку
'Returns:
' Boolean, семантика результата:
'   true  - удалено
'   false - ошибка
'See Also:
'  ConnExec
'  FindPoddon
'  GetBRIEFFromXMLField
'  MyRound
'  MyRound2
'  PrintSticker
'  SaveRCVRowToCore
'  SaveShipRowToCore
'  UpdateMyPalet
'Example:
' dim variable as Boolean
' variable = me.CleanRCVAtCore(...параметры...)
Public Function CleanRCVAtCore(ByVal Item As ITTIN.Application) As Boolean
Attribute CleanRCVAtCore.VB_HelpID = 800
  Dim conn As ADODB.Connection
  Set conn = GetCoreConn
  Dim cmd As ADODB.Command
  Dim rs As ADODB.Recordset
  Dim rsitem As ADODB.Recordset
  Dim poddon As ITTPL_DEF
  
  Dim code As String
  Dim palID As String
  Dim oid As String
  
  code = GetBRIEFFromXMLField(Item.ITTIN_DEF.Item(1).QryCode)
  
  Set conn = GetCoreConn
  If conn.State <> adStateOpen Then
    conn.open
  End If
  
  Dim RollbackMe As Boolean
  
  RollbackMe = False
  conn.BeginTrans
  err.Clear
  
  Set cmd = New ADODB.Command
  cmd.CommandType = adCmdText
  cmd.CommandText = "delete from stock where ORD_NUM ='" & code & "'"
  Set cmd.ActiveConnection = conn
  On Error Resume Next
  cmd.Execute
   If err.Number <> 0 Then
    ''' MsgBox err.Description
    RollbackMe = True
  End If
  
   Set cmd = New ADODB.Command
  cmd.CommandType = adCmdText
  cmd.CommandText = "delete from receiving_history where REF_NUMBER ='" & code & "'"
  Set cmd.ActiveConnection = conn
  On Error Resume Next
  cmd.Execute
   If err.Number <> 0 Then
    ''' MsgBox err.Description
    RollbackMe = True
  End If
     Set cmd = New ADODB.Command
  cmd.CommandType = adCmdText
  cmd.CommandText = "delete from history where code =6 and ORD_NUM ='" & code & "'"
  Set cmd.ActiveConnection = conn
  On Error Resume Next
  cmd.Execute
   If err.Number <> 0 Then
    ''' MsgBox err.Description
    RollbackMe = True
  End If


  If RollbackMe Then
    conn.RollbackTrans
  Else
    conn.CommitTrans
    CleanRCVAtCore = True
  End If
  
End Function

'сохраниеть данные о приемке паллеты в CORE
'Parameters:
'[IN]   Item , тип параметра: ITTIN.Application - заказ в ВК,
'[IN]   CurRow , тип параметра: ITTIN_QLINE -строка заказа в ВК,
'[IN][OUT]   LinePal , тип параметра: ITTIN_PALET - паллета кзаказу в ВК,
'[IN][OUT]   NewPlace , тип параметра: String - новая ячейка,
'[IN][OUT]   QueryCode , тип параметра: String  - код заказа
'Returns:
' Boolean, семантика результата:
'   true  -записано
'   false -нет
'See Also:
'  CleanRCVAtCore
'  ConnExec
'  FindPoddon
'  GetBRIEFFromXMLField
'  MyRound
'  MyRound2
'  PrintSticker
'  SaveShipRowToCore
'  UpdateMyPalet
'Example:
' dim variable as Boolean
' variable = me.SaveRCVRowToCore(...параметры...)
Public Function SaveRCVRowToCore(ByVal Item As ITTIN.Application, ByVal CurRow As ITTIN_QLINE, LinePal As ITTIN_PALET, NewPlace As String, QueryCode As String) As Boolean
Attribute SaveRCVRowToCore.VB_HelpID = 835

 On Error Resume Next
  
  
  Dim conn As ADODB.Connection
  Set conn = GetCoreConn
  Dim cmd As ADODB.Command
  Dim rs As ADODB.Recordset
  Dim rsitem As ADODB.Recordset
  Dim poddon As ITTPL_DEF
  
  Dim rlID As String
  Dim palID As String
  Dim oid As String
  oid = Manager.GetIDFromXMLField(Item.ITTIN_DEF.Item(1).QryCode)
  rlID = Manager.GetIDFromXMLField(CurRow.good_id)
  palID = LinePal.TheNumber.CorePalette_ID
  
  
  
  
  ' запрашиваем свободное место в буферной зоне
  Dim bzrs As ADODB.Recordset
  Dim loccode As ADODB.Recordset
  Dim bzid As String
  Set conn = GetCoreConn
  If conn.State <> adStateOpen Then
    conn.open
  End If
  
  
  Set rs = ConnExec("select * from RECEIVING_LINE where order_id=" & oid & " and item_id=" & rlID, conn)
  If rs.EOF Then
    Exit Function
  End If
  
'  Set bzrs = connExec( _
'    "select  distinct location_id id from stock join location on location.id = stock.location_id " & _
'    " Where stock.item_iD = " & Manager.GetIDFromXMLField(curRow.good_id) & _
'    " group by location_id,location.description " & _
'    " having count(*) < convert(integer, substring(location.description, 0,charindex(';',location.description,0)))" _
'  )
'
  
  Dim s As String
  Dim netto As Double
  Dim partia As String
  Dim kilplace As String
  Dim factoryname As String
  Dim countryname As String
  
  
  If LinePal.made_country Is Nothing Then
    countryname = ""
  Else
    countryname = LinePal.made_country.Name
  End If
    
  If LinePal.factory Is Nothing Then
    factoryname = ""
  Else
    factoryname = LinePal.factory.Name
  End If
  
  If LinePal.KILL_NUMBER Is Nothing Then
    kilplace = ""
  Else
    kilplace = LinePal.KILL_NUMBER.Name
  End If
  
  If LinePal.PartRef Is Nothing Then
    partia = ""
  Else
    partia = CurRow.PartRef.Name
  End If
  
  
  If LinePal.isCalibrated = Boolean_Da Then
    netto = LinePal.KorobNetto * LinePal.CaliberQuantity
  Else
    netto = (LinePal.GoodWithPaletWeight - LinePal.PalWeight - LinePal.FullPackageWeight)
  End If
  
  ' сохраняем чего можем в паддоне
  Set poddon = LinePal.TheNumber
  poddon.CaliberQuantity = LinePal.CaliberQuantity
  poddon.PackageWeight = LinePal.FullPackageWeight
  poddon.CurrentWeightBrutto = LinePal.GoodWithPaletWeight
  
  poddon.save
  bzid = NewPlace
  
  Dim sBrack As String
  Dim sStatus As Integer
  Dim sCaliber As String
  
  If LinePal.isCalibrated = Boolean_Da Then
    sCaliber = "'1'"
  Else
    sCaliber = "''"
  End If
  
  
  If LinePal.IsBrak = Boolean_Da Then
  sBrack = "БРАК"
  sStatus = 101
  Else
  sBrack = ""
  sStatus = 0
  End If

  If Not IsNumeric(bzid) Then
    Dim locrs As ADODB.Recordset
    Set locrs = ConnExec("select * from location where code='" & bzid & "'", conn)
    If Not locrs.EOF Then
      bzid = locrs!id
    End If
    locrs.Close
    Set locrs = Nothing
  End If
  
'  Dim fs As frmSaving
'  Set fs = New frmSaving
'  fs.Show
'  DoEvents
 
  ' формируем запись в сток
  s = "insert into stock(SITE_ID,ITEM_ID,LOCATION_ID,ORDER_ID,QTY_ON_HAND," & _
  "status,UNIT_COST,UOM,LOT_SN,REF_NUM," & _
  "ORD_NUM,PALLET_ID,custom_field1,custom_field6,custom_field11,custom_field5,exp_date,custom_field3,custom_field4,custom_field12,custom_field2,custom_field9,custom_field7)" & _
  "values(" & _
  "1," & Manager.GetIDFromXMLField(CurRow.good_id) & ",'" & bzid & "',null," & MyRound2(netto) & _
   "," & sStatus & ",0,'" & CurRow.edizm & "','" & partia & "','" & QueryCode & "'," & _
  "'" & QueryCode & "'," & palID & "," & MyRound2(LinePal.CaliberQuantity) & ",'" & countryname & "','" & kilplace & "','" & LinePal.made_date & "'," & MakeMSSQLDate(LinePal.exp_date) & ",'" & MyRound2(CurRow.PackageWeight) & "','" & factoryname & "','" & sBrack & "'," & sCaliber & ",'" & LinePal.made_date_to & "','" & Left(LinePal.vetsved, 50) & "') "

  Dim RollbackMe As Boolean
  
  log.message "Сохранение паллеты в заказ " & GetBRIEFFromXMLField(Item.ITTIN_DEF.Item(1).QryCode) & " товар " & GetBRIEFFromXMLField(CurRow.good_id) & " ПОДДОН:" & LinePal.TheNumber.TheNumber
  
  ' начало транзакции
  RollbackMe = False
  conn.BeginTrans
  err.Clear
  
  Set cmd = New ADODB.Command
  cmd.CommandType = adCmdText
  cmd.CommandText = s
  Set cmd.ActiveConnection = conn
  On Error Resume Next
  cmd.Execute
   If err.Number <> 0 Then
    ''' MsgBox err.Description
    log.Error err.Description
    RollbackMe = True
    GoTo finalize
  End If
      
  Set loccode = ConnExec("select code from location where id=" & bzid, conn)
  If Not loccode.EOF Then
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdText
    cmd.CommandText = "update pallet set location_id=" & bzid & " where id=" & palID
    Set cmd.ActiveConnection = conn
    On Error Resume Next
    cmd.Execute
     If err.Number <> 0 Then
      ''' MsgBox err.Description
      RollbackMe = True
      log.Error err.Description
      GoTo finalize
    End If
  End If
  
  
 
  
  
  
'  запись в историю в заказе в CORE
  cmd.CommandText = "INSERT INTO RECEIVING_HISTORY(custom_field12, [REF_NUMBER], [QTY_REC], [UOM], [LOT_SN], [EXP_DATE], [UNIT_PRICE], [COMMENTS], [REC_DATE], [TRACK_NUMBER2], [TRACK_NUMBER3], [LOCATION], [PALLET], [CONTAINER], [STATUS], [ORDER_ID], [ITEM_ID], [USER_ID], custom_field1,custom_field3,custom_field4,custom_field11,custom_field5,custom_field6,custom_field2,custom_field9,custom_field7)" & _
  "VALUES( '" & sBrack & "','" & QueryCode & "', " & MyRound2(netto) & ",'" & CurRow.edizm & "', '" & partia & "'," & MakeMSSQLDate(LinePal.exp_date) & " , 0, ' ', getdate(), '" & Item.ITTIN_DEF.Item(1).TranspNumber & "', '" & Item.ITTIN_DEF.Item(1).TranspNumber & "','" & LinePal.BufferZonePlace & "','" & poddon.TheNumber & "', '" & Item.ITTIN_DEF.Item(1).Container & "', 0, " & rs!order_id & "," & rs!item_id & ",1," & MyRound2(LinePal.CaliberQuantity) & ",'" & MyRound2(LinePal.PackageWeight) & "','" & factoryname & "','" & kilplace & "','" & LinePal.made_date & "','" & countryname & "'," & sCaliber & ",'" & LinePal.made_date_to & "','" & Left(LinePal.vetsved, 50) & "')"
  Set cmd.ActiveConnection = conn
  err.Clear
  cmd.Execute
  
  If err.Number <> 0 Then
    ''' MsgBox err.Description
     RollbackMe = True
     log.Error err.Description
     GoTo finalize
  End If
  
  
  Set rsitem = ConnExec("select * from item where id =" & rs!item_id, conn)
  
  
  ' запись в history
  cmd.CommandText = "INSERT INTO HISTORY(custom_field12,code,stamp,user_name,site_id,QTY_ON_HAND, [REF_NUM],[ORD_NUM], [QTY], [UOM], [LOT_SN]," & _
  " [EXP_DATE], [UNIT_COST], [LOCATION], [PALLET], [CONTAINER], [STATUS],  [ITEM],[description],  " & _
  "custom_field1,custom_field3,custom_field4,custom_field11,custom_field5,custom_field6,custom_field9,custom_field7)" & _
  "VALUES(" & _
  "'" & sBrack & "',6,getdate(), 'sa',1,0,'" & QueryCode & "','" & QueryCode & "', " & MyRound2(netto) & ",'" & CurRow.edizm & "', '" & partia & "'," & _
  MakeMSSQLDate(LinePal.exp_date) & " , 0, '" & LinePal.BufferZonePlace & "','" & poddon.TheNumber & "', '" & Item.ITTIN_DEF.Item(1).Container & "',0," & _
  "'" & rsitem!code & "','" & rsitem!Description & " '," & _
  MyRound2(LinePal.CaliberQuantity) & ",'" & MyRound2(LinePal.PackageWeight) & "','" & factoryname & "','" & kilplace & "','" & LinePal.made_date & "','" & countryname & "','" & LinePal.made_date_to & "','" & Left(LinePal.vetsved, 50) & "' )"
  
  Set cmd.ActiveConnection = conn
  err.Clear
  cmd.Execute
  
  If err.Number <> 0 Then
    ''' MsgBox err.Description
     RollbackMe = True
     log.Error err.Description
     GoTo finalize
  End If
  
  
  ' запись в строку заказа
  cmd.CommandText = "update RECEIVING_LINE SET status=1, QTY_ALT_PREV_REC=isnull(QTY_ALT_PREV_REC,0)+" & MyRound2(LinePal.CaliberQuantity) & ", QTY_PREV_REC =isnull(QTY_PREV_REC,0)+" & MyRound2(netto) & ", MADE_DATE=" & MakeMSSQLDate(CurRow.made_date) & ", EXP_DATE=" & MakeMSSQLDate(CurRow.exp_date) & ",PROD_COUNTRY='" & countryname & "',KILL_NUMBER='" & kilplace & "',LOT_SN='" & partia & "' where order_id=" & oid & " and item_ID=" & rlID
  
  err.Clear
  Set cmd.ActiveConnection = conn
  cmd.Execute
  If err.Number <> 0 Then
    ''' MsgBox err.Description
     RollbackMe = True
     log.Error err.Description
     GoTo finalize
  End If
  
  Dim ordid As String
  
  Dim def As ITTIN_DEF
  Set def = Item.ITTIN_DEF.Item(1)
  ordid = Manager.GetIDFromXMLField(def.QryCode)
  
  
  'запись в сам заказ
  cmd.CommandText = "update RECEIVING_ORDER SET status=1, TRACK_NUMBER1='" & def.TranspNumber & "' ,COMMENT1='" & def.Container & "', ZIP='" & def.temp_in_track & "'  where ID=" & ordid
  
  err.Clear
  Set cmd.ActiveConnection = conn
  cmd.Execute
  If err.Number <> 0 Then
    ''' MsgBox err.Description
     log.Error err.Description
     RollbackMe = True
     GoTo finalize
  End If
  
finalize:
  
  ' завершение транзакции
  If RollbackMe Then
    conn.RollbackTrans
    log.message "Откат транзакции"
  Else
    conn.CommitTrans
    SaveRCVRowToCore = True
    log.message "Сохранено"
  End If
'  Unload fs
'  Set fs = Nothing
'  DoEvents
End Function
'Сохранить данные об отгрузке паллеты в Core
'Parameters:
'[IN]   ShipOrder , тип параметра: String - код заказа,
'[IN]   NewPlace , тип параметра: String - ячейка для остатков,
'[IN][OUT]   Item , тип параметра: ITTOUT.Application - заказ на отгрузку,
'[IN][OUT]   poddon , тип параметра: ITTPL_DEF- описание пллеты,
'[IN]   CurRow , тип параметра: ITTOUT_LINES- строка заказа,
'[IN][OUT]   LinePal , тип параметра: ITTOUT_PALET- паллета в заказе,
'[IN]   isFull , тип параметра: Boolean  -отгрузка полностью
'Returns:
' Boolean, семантика результата:
'   true  - удачно
'   false - ошибка
'See Also:
'  CleanRCVAtCore
'  ConnExec
'  FindPoddon
'  GetBRIEFFromXMLField
'  MyRound
'  MyRound2
'  PrintSticker
'  SaveRCVRowToCore
'  UpdateMyPalet
'Example:
' dim variable as Boolean
' variable = me.SaveShipRowToCore(...параметры...)
Public Function SaveShipRowToCore(ByVal ShipOrder As String, ByVal NewPlace As String, ByRef Item As ITTOUT.Application, ByRef poddon As ITTPL_DEF, ByVal CurRow As ITTOUT_LINES, LinePal As ITTOUT_PALET, ByVal isFull As Boolean) As Boolean
Attribute SaveShipRowToCore.VB_HelpID = 840
On Error Resume Next
  
  Dim conn As ADODB.Connection
  Set conn = GetCoreConn
  Dim cmd As ADODB.Command
  Dim rs As ADODB.Recordset
  Dim oid As String
  Dim rlID As String
  Dim palID As String
  Dim palNum As String
  
  Set conn = GetCoreConn
  If conn.State <> adStateOpen Then
    conn.open
  End If
  

  oid = Manager.GetIDFromXMLField(Item.ITTOUT_DEF.Item(1).ShipOrder)
  rlID = Manager.GetIDFromXMLField(CurRow.good_id)
  palID = LinePal.TheNumber.CorePalette_ID
  palNum = LinePal.TheNumber.TheNumber
  
  
  Dim strs As ADODB.Recordset
  Dim LCRS As ADODB.Recordset
  Dim rsitem As ADODB.Recordset
  Dim lcrscode As String
  
  ' начальные проверки
  Set strs = ConnExec("select * from STOCK where PALLET_STATUS is null and  PALLET_ID=" & palID, conn)

  If strs.EOF Then
    MsgBox "Не обнаружены данные о палете"
    Exit Function
  End If
  
  
  Set LCRS = ConnExec("select * from location where id=" & strs!location_id, conn)
  If LCRS Is Nothing Then
    Set LCRS = ConnExec("select * from location where id is null", conn)
  End If
  If Not LCRS.EOF Then
    lcrscode = LCRS!code
  Else
    lcrscode = ""
  End If
  LCRS.Close
  Set LCRS = Nothing
  
  Set rsitem = ConnExec("select * from [item] where [id]=" & strs!item_id, conn)
    
  Dim rsitemcode As String
  Dim rsitemdesc As String
  
  If Not rsitem.EOF Then
    rsitemcode = rsitem!code
    rsitemdesc = rsitem!Description
    
    rsitem.Close
    Set rsitem = Nothing
  Else
     MsgBox "Не обнаружены данные о товаре"
     Exit Function
  End If
  

  
  '
  Dim w As Double
  Dim Q As Long
  Dim netto As Double
  
  
  netto = LinePal.GoodWithPaletWeight - LinePal.FullPackageWeight - poddon.Weight
  
  
'  Dim fs As frmSaving
'  Set fs = New frmSaving
'  fs.Show
  
  
 log.message "Сохраняем отгрузку. Заказ: " & GetBRIEFFromXMLField(Item.ITTOUT_DEF.Item(1).ShipOrder) & " поддон:" & palNum
  
  ' сохраняем данные из стока в параметрах отгружаемой паллеты
  With LinePal
  
    Set .made_country = FindCountry("" & strs!custom_field6)
    
    On Error Resume Next
  
    If Not .made_country Is Nothing Then
      Set .factory = FindFactory(.made_country.id, "" & strs!custom_field4)
    End If
    
    If Not .factory Is Nothing Then
      Set .KILL_NUMBER = FindKill(.factory.id, "" & strs!custom_field11)
    End If
    
    Set .PartRef = FindPartia(GetBRIEFFromXMLField(CurRow.LineAtQuery), strs!LOT_SN)
    
    If Not IsNull(strs!custom_field5) Then .made_date = CDate(strs!custom_field5)
    
    If Not IsNull(strs!exp_date) Then .exp_date = strs!exp_date
    
    err.Clear
    .save
    If err.Number <> 0 Then
        log.Error err.Description
    Else
        log.message "Cохранен в весовом"
    End If
    
    
  End With
  
  
  
  ' начало транзакции
  
  Dim RollbackMe As Boolean
  
  ' начало транзакции
  RollbackMe = False
  conn.BeginTrans
    
    
  Set cmd = New ADODB.Command
  cmd.CommandType = adCmdText
  cmd.CommandText = "update pallet set order_id=null where id=" & palID
  Set cmd.ActiveConnection = conn
  On Error Resume Next
  err.Clear
  cmd.Execute
   If err.Number <> 0 Then
    log.Error err.Description
    RollbackMe = True
    GoTo finalize
  End If
  
  If isFull Then
  
    ' отгрузка полной паллеты
    w = MyRound(strs!QTY_ON_HAND)
    Q = MyRound("0" & strs!custom_field1)
    Set poddon = LinePal.TheNumber
    If LinePal.CaliberQuantity <> Q Or Abs(netto - w) > 0.001 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdText
        cmd.CommandText = "update STOCK set location_id = null, pallet_id =null, PALLET_STATUS=1, QTY_ON_HAND =isnull(QTY_ON_HAND,0)-" & MyRound2(netto) & ", CUSTOM_FIELD1='" & MyRound2(Q - LinePal.CaliberQuantity) & "' where PALLET_STATUS is null and PALLET_ID=" & palID
        Set cmd.ActiveConnection = conn
        On Error Resume Next
        cmd.Execute
         If err.Number <> 0 Then
          log.Error err.Description
          RollbackMe = True
           GoTo finalize
       
        End If
    Else
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdText
        cmd.CommandText = "delete from STOCK where PALLET_STATUS is null and PALLET_ID=" & palID
        Set cmd.ActiveConnection = conn
        On Error Resume Next
        cmd.Execute
         If err.Number <> 0 Then
          log.Error err.Description
          RollbackMe = True
           GoTo finalize
        End If
    End If
    
    
    
    
    If lcrscode <> "" Then
      cmd.CommandText = "INSERT INTO SHIPPING_HISTORY( [REF_NUMBER], [QTY_SHIP], [UOM], [LOT_SN], [EXP_DATE], [UNIT_PRICE], [COMMENTS], [ship_DATE], [TRACK_NUMBER],  [LOCATION_id],[LOCATION], [PALLET], [CONTAINER], [STATUS], [ORDER_ID], [ITEM_ID], [USER_ID], custom_field1,custom_field3,custom_field4,BOX_NUMBER,custom_field5,custom_field6,custom_field11,custom_field2,custom_field9,custom_field7)" & _
      "VALUES( '" & ShipOrder & "', " & MyRound2(netto) & ",'" & strs!UOM & "', '" & strs!LOT_SN & "'," & MakeMSSQLDate(strs!exp_date) & " , 0, ' ', getdate(), '" & Item.ITTOUT_DEF.Item(1).TranspNumber & "', " & IIf(IsNull(strs!location_id), "null", strs!location_id) & ",'" & lcrscode & "','" & palNum & "', '" & Item.ITTOUT_DEF.Item(1).Container & "', 0, " & Manager.GetIDFromXMLField(Item.ITTOUT_DEF.Item(1).ShipOrder) & "," & strs!item_id & ",1," & MyRound2(LinePal.CaliberQuantity) & ",'" & strs!custom_field3 & "','" & strs!custom_field4 & "','','" & strs!custom_field5 & "','" & strs!custom_field6 & "','" & strs!custom_field11 & "','" & strs!custom_field2 & "','" & strs!custom_field9 & "','" & strs!custom_field7 & "' )"
    Else
      cmd.CommandText = "INSERT INTO SHIPPING_HISTORY( [REF_NUMBER], [QTY_SHIP], [UOM], [LOT_SN], [EXP_DATE], [UNIT_PRICE], [COMMENTS], [ship_DATE], [TRACK_NUMBER],  [LOCATION_id],[LOCATION], [PALLET], [CONTAINER], [STATUS], [ORDER_ID], [ITEM_ID], [USER_ID], custom_field1,custom_field3,custom_field4,BOX_NUMBER,custom_field5,custom_field6,custom_field11,custom_field2,custom_field9,custom_field7)" & _
      "VALUES( '" & ShipOrder & "', " & MyRound2(netto) & ",'" & strs!UOM & "', '" & strs!LOT_SN & "'," & MakeMSSQLDate(strs!exp_date) & " , 0, ' ', getdate(), '" & Item.ITTOUT_DEF.Item(1).TranspNumber & "', " & IIf(IsNull(strs!location_id), "null", strs!location_id) & ",'','" & palNum & "', '" & Item.ITTOUT_DEF.Item(1).Container & "', 0, " & Manager.GetIDFromXMLField(Item.ITTOUT_DEF.Item(1).ShipOrder) & "," & strs!item_id & ",1," & MyRound2(LinePal.CaliberQuantity) & ",'" & strs!custom_field3 & "','" & strs!custom_field4 & "','','" & strs!custom_field5 & "','" & strs!custom_field6 & "','" & strs!custom_field11 & "','" & strs!custom_field2 & "','" & strs!custom_field9 & "','" & strs!custom_field7 & "' )"
    End If
    
    Set cmd.ActiveConnection = conn
    err.Clear
    cmd.Execute
    
    If err.Number <> 0 Then
      log.Error err.Description
      RollbackMe = True
       GoTo finalize
    End If
    
    ' фиксируем отгрузку в history
    If lcrscode <> "" Then
      ' отгрузка из ячейки
      cmd.CommandText = "INSERT INTO HISTORY(site_id, code, " & _
      " [REF_NUM],ord_num, [QTY], [QTY_ON_HAND], [UOM]," & _
      " [LOT_SN], [EXP_DATE]," & _
      " [UNIT_COST],  [stamp],   " & _
      " [LOCATION], [PALLET], [CONTAINER], [STATUS], " & _
      " [ITEM],[DESCRIPTION], [USER_name], " & _
      " custom_field1,custom_field3,custom_field4,custom_field5,custom_field6,custom_field11,custom_field2" & _
      ",custom_field9,custom_field7)" & _
      " VALUES(1, 5,'" & _
      ShipOrder & "','" & ShipOrder & "', -(" & MyRound2(netto) & ")," & MyRound2(strs!QTY_ON_HAND) & ",'" & strs!UOM & "', '" & _
      strs!LOT_SN & "'," & MakeMSSQLDate(strs!exp_date) & _
      " , 0, getdate()," & _
      "'" & lcrscode & "','" & palNum & "', '" & Item.ITTOUT_DEF.Item(1).Container & "', 0, " & _
      "'" & rsitemcode & "','" & rsitemdesc & "','sa'," & _
      MyRound2(LinePal.CaliberQuantity) & ",'" & strs!custom_field3 & "','" & strs!custom_field4 & "','" & strs!custom_field5 & "','" & strs!custom_field6 & "','" & strs!custom_field11 & "','" & strs!custom_field2 & "' " & _
      ",'" & strs!custom_field9 & "','" & strs!custom_field7 & "')"
    Else
      ' отгрузка с пола
      cmd.CommandText = "INSERT INTO HISTORY(site_id, code, " & _
      " [REF_NUM],ord_num, [QTY], [QTY_ON_HAND], [UOM]," & _
      " [LOT_SN], [EXP_DATE]," & _
      " [UNIT_COST],  [stamp],   " & _
      " [LOCATION], [PALLET], [CONTAINER], [STATUS], " & _
      " [ITEM],[DESCRIPTION], [USER_name], " & _
      " custom_field1,custom_field3,custom_field4,custom_field5,custom_field6,custom_field11,custom_field2" & _
      " ,custom_field9,custom_field7)" & _
      " VALUES(1, 5,'" & _
      ShipOrder & "','" & ShipOrder & "', -(" & MyRound2(netto) & ")," & MyRound2(strs!QTY_ON_HAND) & ",'" & strs!UOM & "', '" & _
      strs!LOT_SN & "'," & MakeMSSQLDate(strs!exp_date) & _
      " , 0, getdate()," & _
      "'','" & palNum & "', '" & Item.ITTOUT_DEF.Item(1).Container & "', 0, " & _
      "'" & rsitemcode & "','" & rsitemdesc & "','sa'," & _
      MyRound2(LinePal.CaliberQuantity) & ",'" & strs!custom_field3 & "','" & strs!custom_field4 & "','" & strs!custom_field5 & "','" & strs!custom_field6 & "','" & strs!custom_field11 & "','" & strs!custom_field2 & "'" & _
      ",'" & strs!custom_field9 & "','" & strs!custom_field7 & "' )"

    End If
    
    Set cmd.ActiveConnection = conn
    err.Clear
    cmd.Execute
    
    If err.Number <> 0 Then
      log.Error err.Description
      RollbackMe = True
       GoTo finalize
    End If
    
    
    
    cmd.CommandText = "update SHIPPING_LINE SET status=1,  QTY_PREV_ship = isnull(QTY_PREV_ship,0) + " & MyRound2(netto) & ",QTY_ALT_PREV_SHIP=isnull(QTY_ALT_PREV_SHIP,0)+ " & MyRound2(LinePal.CaliberQuantity) & " where order_id=" & oid & " and item_ID=" & rlID
    
    err.Clear
    Set cmd.ActiveConnection = conn
    cmd.Execute
    If err.Number <> 0 Then
      log.Error err.Description
      RollbackMe = True
       GoTo finalize
    End If
    
    cmd.CommandText = "update shipping_order set status=1 where id=" & Manager.GetIDFromXMLField(Item.ITTOUT_DEF.Item(1).ShipOrder)
    Set cmd.ActiveConnection = conn
    err.Clear
    cmd.Execute

    If err.Number <> 0 Then
      log.Error err.Description
      RollbackMe = True
       GoTo finalize
    End If

      
  Else
    ' отгрузка с остатком

    w = MyRound(strs!QTY_ON_HAND)
    Q = MyRound("0" & strs!custom_field1)
    Set poddon = LinePal.TheNumber
  
    Dim loccode As ADODB.Recordset
    Set loccode = Nothing
    If NewPlace <> "" Then
       Set loccode = ConnExec("select * from location where code='" & NewPlace & "'", conn)
     
       If Not loccode.EOF Then
          Set cmd = New ADODB.Command
          cmd.CommandType = adCmdText
          cmd.CommandText = "update pallet set location_id=" & loccode!id & " where id=" & palID
          Set cmd.ActiveConnection = conn
          On Error Resume Next
          cmd.Execute
          If err.Number <> 0 Then
            log.Error err.Description
             RollbackMe = True
             GoTo finalize
          End If
          
          Set cmd = New ADODB.Command
          cmd.CommandType = adCmdText
          cmd.CommandText = "update STOCK set  location_id=" & loccode!id & " where PALLET_STATUS is null and PALLET_ID=" & palID
          Set cmd.ActiveConnection = conn
          On Error Resume Next
          err.Clear
          cmd.Execute
          If err.Number <> 0 Then
            log.Error err.Description
            RollbackMe = True
            GoTo finalize
          End If
       End If
    
    End If
  
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdText
    cmd.CommandText = "update STOCK set QTY_ON_HAND =isnull(QTY_ON_HAND,0)-(" & MyRound2(netto) & "), CUSTOM_FIELD1='" & MyRound2(Q - LinePal.CaliberQuantity) & "' where PALLET_STATUS is null and PALLET_ID=" & palID
    Set cmd.ActiveConnection = conn
    On Error Resume Next
    cmd.Execute
    If err.Number <> 0 Then
      log.Error err.Description
      RollbackMe = True
      GoTo finalize
    End If
    
    
    If Not loccode Is Nothing Then
      cmd.CommandText = "INSERT INTO SHIPPING_HISTORY( [REF_NUMBER], [QTY_SHIP], [UOM], [LOT_SN], [EXP_DATE], [UNIT_PRICE], [COMMENTS], [ship_DATE], [TRACK_NUMBER],  [LOCATION_id],[LOCATION], [PALLET], [CONTAINER], [STATUS], [ORDER_ID], [ITEM_ID], [USER_ID], custom_field1,custom_field3,custom_field4,BOX_NUMBER,custom_field5,custom_field6,custom_field11, custom_field2" & _
      " ,custom_field9,custom_field7" & _
      ")" & _
      "VALUES( '" & ShipOrder & "', " & MyRound2(netto) & ",'" & strs!UOM & "', '" & strs!LOT_SN & "'," & MakeMSSQLDate(strs!exp_date) & " , 0, ' ', getdate(), '" & Item.ITTOUT_DEF.Item(1).TranspNumber & "', '" & loccode!id & "','" & loccode!code & "','" & palNum & "', '" & Item.ITTOUT_DEF.Item(1).Container & "', 0, " & Manager.GetIDFromXMLField(Item.ITTOUT_DEF.Item(1).ShipOrder) & "," & strs!item_id & ",1," & MyRound2(LinePal.CaliberQuantity) & ",'" & strs!custom_field3 & "','" & strs!custom_field4 & "','','" & strs!custom_field5 & "','" & strs!custom_field6 & "','" & strs!custom_field11 & "','" & strs!custom_field2 & "'" & _
      ",'" & strs!custom_field9 & "','" & strs!custom_field7 & "'" & _
      " )"
    Else
      cmd.CommandText = "INSERT INTO SHIPPING_HISTORY( [REF_NUMBER], [QTY_SHIP], [UOM], [LOT_SN], [EXP_DATE], [UNIT_PRICE], [COMMENTS], [ship_DATE], [TRACK_NUMBER],  [LOCATION_id],[LOCATION], [PALLET], [CONTAINER], [STATUS], [ORDER_ID], [ITEM_ID], [USER_ID], custom_field1,custom_field3,custom_field4,BOX_NUMBER,custom_field5,custom_field6,custom_field11, custom_field2" & _
      " ,custom_field9,custom_field7" & _
      ")" & _
      "VALUES( '" & ShipOrder & "', " & MyRound2(netto) & ",'" & strs!UOM & "', '" & strs!LOT_SN & "'," & MakeMSSQLDate(strs!exp_date) & " , 0, ' ', getdate(), '" & Item.ITTOUT_DEF.Item(1).TranspNumber & "', '" & "" & "','" & "" & "','" & palNum & "', '" & Item.ITTOUT_DEF.Item(1).Container & "', 0, " & Manager.GetIDFromXMLField(Item.ITTOUT_DEF.Item(1).ShipOrder) & "," & strs!item_id & ",1," & MyRound2(LinePal.CaliberQuantity) & ",'" & strs!custom_field3 & "','" & strs!custom_field4 & "','','" & strs!custom_field5 & "','" & strs!custom_field6 & "','" & strs!custom_field11 & "','" & strs!custom_field2 & "'" & _
      ",'" & strs!custom_field9 & "','" & strs!custom_field7 & "'" & _
      " )"
    End If
    Set cmd.ActiveConnection = conn
    err.Clear
    cmd.Execute

    If err.Number <> 0 Then
      log.Error err.Description
      RollbackMe = True
      GoTo finalize
    End If
    
    
    
     ' фиксируем перемещение
      cmd.CommandText = "INSERT INTO HISTORY(site_id, code, " & _
      " [REF_NUM],ord_num, [QTY], [QTY_ON_HAND], [UOM]," & _
      " [LOT_SN], [EXP_DATE]," & _
      " [UNIT_COST],  [stamp],   " & _
      " [LOCATION], [PALLET], [CONTAINER], [STATUS], " & _
      " [ITEM],[DESCRIPTION], [USER_name], " & _
      " custom_field1,custom_field3,custom_field4,custom_field5,custom_field6,custom_field11, custom_field2" & _
      " ,custom_field9,custom_field7" & _
      " )VALUES(1, 8,'" & _
      ShipOrder & "','" & ShipOrder & "', 0," & MyRound2(strs!QTY_ON_HAND) & ",'" & strs!UOM & "', '" & _
      strs!LOT_SN & "'," & MakeMSSQLDate(strs!exp_date) & _
      " , 0, getdate()," & _
      "'" & NewPlace & "','" & palNum & "', '" & Item.ITTOUT_DEF.Item(1).Container & "', 0, " & _
      "'" & rsitemcode & "','" & rsitemdesc & "','sa'," & _
      strs!custom_field1 & ",'" & strs!custom_field3 & "','" & strs!custom_field4 & "','" & strs!custom_field5 & "','" & strs!custom_field6 & "','" & strs!custom_field11 & "','" & strs!custom_field2 & "'" & _
      ",'" & strs!custom_field9 & "','" & strs!custom_field7 & "'" & _
      " )"
    
    Set cmd.ActiveConnection = conn
    err.Clear
    cmd.Execute
    
    If err.Number <> 0 Then
      log.Error err.Description
      RollbackMe = True
      GoTo finalize
    End If
    
    
    ' фиксируем отгрузку в history
     cmd.CommandText = "INSERT INTO HISTORY(site_id, code, " & _
    " [REF_NUM],ord_num, [QTY], [QTY_ON_HAND], [UOM]," & _
    " [LOT_SN], [EXP_DATE]," & _
    " [UNIT_COST],  [stamp],   " & _
    " [LOCATION], [PALLET], [CONTAINER], [STATUS], " & _
    " [ITEM],[DESCRIPTION], [USER_name], " & _
    " custom_field1,custom_field3,custom_field4,custom_field5,custom_field6,custom_field11,custom_field2" & _
    " ,custom_field9,custom_field7" & _
    " )VALUES(1, 5,'" & _
    ShipOrder & "','" & ShipOrder & "', -(" & MyRound2(netto) & ")," & MyRound2(strs!QTY_ON_HAND) & ",'" & strs!UOM & "', '" & _
    strs!LOT_SN & "'," & MakeMSSQLDate(strs!exp_date) & _
    " , 0, getdate()," & _
    "'" & NewPlace & "','" & palNum & "', '" & Item.ITTOUT_DEF.Item(1).Container & "', 0, " & _
    "'" & rsitemcode & "','" & rsitemdesc & "','sa'," & _
    "'" & MyRound2(LinePal.CaliberQuantity) & "','" & strs!custom_field3 & "','" & strs!custom_field4 & "','" & strs!custom_field5 & "','" & strs!custom_field6 & "','" & strs!custom_field11 & "','" & strs!custom_field2 & "' " & _
    ",'" & strs!custom_field9 & "','" & strs!custom_field7 & "'" & _
    ")"
    
    Set cmd.ActiveConnection = conn
    err.Clear
    cmd.Execute
    
    If err.Number <> 0 Then
      log.Error err.Description
      RollbackMe = True
      GoTo finalize
    End If
    
    cmd.CommandText = "update shipping_order set status=1 where id=" & Manager.GetIDFromXMLField(Item.ITTOUT_DEF.Item(1).ShipOrder)
    Set cmd.ActiveConnection = conn
    err.Clear
    cmd.Execute

    If err.Number <> 0 Then
      log.Error err.Description
      RollbackMe = True
      GoTo finalize
    End If
    
    cmd.CommandText = "update SHIPPING_LINE SET status=1,QTY_PREV_ship = isnull(QTY_PREV_ship,0) +(" & MyRound2(netto) & "),QTY_ALT_PREV_SHIP=isnull(QTY_ALT_PREV_SHIP,0)+ (" & MyRound2(LinePal.CaliberQuantity) & ") where order_id=" & oid & " and item_ID=" & rlID
    err.Clear
    Set cmd.ActiveConnection = conn
    cmd.Execute
    If err.Number <> 0 Then
      log.Error err.Description
      RollbackMe = True
      GoTo finalize
    End If
    
  End If
  
finalize:
  ' завершение транзакции
  
  If RollbackMe Then
    conn.RollbackTrans
    log.message "Откат транзакции"
  Else
    conn.CommitTrans
    SaveShipRowToCore = True
    log.message "Сохранено в CORE"
  End If
  
'  Unload fs
'  Set fs = Nothing
End Function

