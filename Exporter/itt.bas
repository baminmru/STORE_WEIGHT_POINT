Attribute VB_Name = "itt"
Option Explicit

Public Function MyRound(ByVal s As String) As Double
  Dim e As String, out As Double
  e = Replace(s, ",", ".")
  out = Val("0" & e)
  out = Round(out + 0.001, 2)
  MyRound = out

End Function

Public Function MyRound2(ByVal s As String) As String
  Dim e As String, out As Double
  e = Replace(s, ",", ".")
  out = Val("0" & e)
  out = Round(out + 0.001, 2)
  e = out
  MyRound2 = Replace(e, ",", ".")

End Function

Public Function FindPoddon(ByVal TheNumber As String) As Object
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

Private Function ORDERID(ByVal s As String) As String
Dim out() As String
On Error Resume Next
out = Split(s, " от")

ORDERID = Replace(out(0), " ", "")

End Function

Public Sub PrintSticker(ByVal poddon As ITTPL_DEF, Optional CaliberWeight As Double = -1)



If GetSetting("RBH", "ITTSETTINGS", "PSTICKER", 1) Then
  
  If MsgBox("Напечатать стикер на поддон?", vbYesNo) = vbYes Then
  
    Dim strs As ADODB.Recordset
    Dim itemrs As ADODB.Recordset
    Dim locrs As ADODB.Recordset
    Dim ordrs As ADODB.Recordset
    
    Dim conn As ADODB.Connection
    Set conn = Manager.GetCustomObjects("refref")
    If conn.State <> adStateOpen Then
      conn.Open
    End If
  
    Set strs = conn.Execute("select * from STOCK where PALLET_STATUS is null and  PALLET_ID=" & poddon.CorePalette_ID)
    If strs.EOF Then
      MsgBox "Поддон не принят на склад"
      Exit Sub
    End If
    Set locrs = conn.Execute("select * from location where id=" & strs!location_id)
    Set itemrs = conn.Execute("select * from item where id=" & strs!item_id)
    Set ordrs = conn.Execute("select Name from partner join receiving_order on partner.id = receiving_order.partner_id where receiving_order.number='" & ORDERID(strs!ORD_NUM) & "'")
      
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
    Printer.Print strs!ORD_NUM
    
    Printer.Font = "Arial CYR"
    Printer.FontSize = 32
    
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
    
    
    
    If strs!status = 101 Then
      Printer.Print "БРАК"
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
    Printer.Print strs!lot_sn & ""
    
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
'      Printer.FontSize = 72
'      Printer.Print "Поддон №"
'      Printer.Print poddon.Code
      Printer.Font = "Arial CYR"
      Printer.FontBold = False
      Printer.Print "Ячейка .№";
      Printer.FontSize = 48
       Printer.FontBold = True
      Printer.Print locrs!code
      Printer.EndDoc
    End If
    
    
   Exit For
  End If
 Next
  End If
  End If
bye2:
  
  Exit Sub
  
bye:
  If err.Number > 0 Then
    MsgBox err.Description, , "Печать документов на поддон"
  End If
End Sub


Public Sub UpdateMyPalet(ByVal Pallet As ITTIN_PALET)

  Dim conn As ADODB.Connection
  Set conn = Manager.GetCustomObjects("refref")
  If conn.State <> adStateOpen Then
      conn.Open
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
  
  s = ""
  s = s & " Update receiving_history"
  s = s & " set QTY_REC=" & MyRound2(Pallet.GoodWithPaletWeight - Pallet.FullPackageWeight - Pallet.PalWeight) & ", custom_field1='" & Pallet.CaliberQuantity & "' "
  s = s & " where pallet ='" & poddon.TheNumber & "' and order_id=" & ordid
  
  conn.Execute s
  
  s = ""
  s = s & " Update receiving_line"
  s = s & " Set QTY_PREV_REC = " & MyRound2(qline.CurValue) & " "
  s = s & " , QTY_ALT_PREV_REC = " & cnt & " "
  s = s & " Where order_id = " & ordid & " and item_id=" & good_id
  
  conn.Execute s
  
  s = ""
  s = s & " Update stock"
  s = s & " set QTY_ON_HAND=" & MyRound2(Pallet.GoodWithPaletWeight - Pallet.FullPackageWeight - poddon.Weight) & ", custom_field1='" & Pallet.CaliberQuantity & "'"
  s = s & " Where pallet_id = " & poddon.CorePalette_ID & " And pallet_status Is Null"
  
  
  conn.Execute s
  
'  s = ""
'  s = s & " Update HISTORY"
'  s = s & " set QTY=" & MyRound2(Pallet.GoodWithPaletWeight - Pallet.FullPackageWeight - poddon.Weight) & ", custom_field1='" & Pallet.CaliberQuantity & "'"
'  s = s & " Where CODE=6 and REF_NUM='" & GetBRIEFFromXMLField(qry.ITTIN_DEF.Item(1).QryCode) & "' and pallet = " & poddon.CorePalette_ID & " and item='" & qline.articul & "'"
'
'
'  conn.Execute s
  
End Sub

Public Function GetBRIEFFromXMLField(ByVal XML As String) As String
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



Public Sub CleanRCVAtCore(ByVal Item As ITTIN.Application)
  Dim conn As ADODB.Connection
  Set conn = Manager.GetCustomObjects("refref")
  Dim cmd As ADODB.Command
  Dim rs As ADODB.Recordset
  Dim rsitem As ADODB.Recordset
  Dim poddon As ITTPL_DEF
  
  Dim code As String
  Dim palID As String
  Dim oid As String
  
  code = GetBRIEFFromXMLField(Item.ITTIN_DEF.Item(1).QryCode)
  
  Set conn = Manager.GetCustomObjects("refref")
  If conn.State <> adStateOpen Then
    conn.Open
  End If
  
  Set cmd = New ADODB.Command
  cmd.CommandType = adCmdText
  cmd.CommandText = "delete from stock where ORD_NUM ='" & code & "'"
  Set cmd.ActiveConnection = conn
  On Error Resume Next
  cmd.Execute
   If err.Number <> 0 Then
    MsgBox err.Description
  End If
  
   Set cmd = New ADODB.Command
  cmd.CommandType = adCmdText
  cmd.CommandText = "delete from receiving_history where REF_NUMBER ='" & code & "'"
  Set cmd.ActiveConnection = conn
  On Error Resume Next
  cmd.Execute
   If err.Number <> 0 Then
    MsgBox err.Description
  End If
     Set cmd = New ADODB.Command
  cmd.CommandType = adCmdText
  cmd.CommandText = "delete from history where code =6 and ORD_NUM ='" & code & "'"
  Set cmd.ActiveConnection = conn
  On Error Resume Next
  cmd.Execute
   If err.Number <> 0 Then
    MsgBox err.Description
  End If

  
End Sub


Public Sub SaveRCVRowToCore(ByVal Item As ITTIN.Application, ByVal CurRow As ITTIN_QLINE, LinePal As ITTIN_PALET, NewPlace As String, QueryCode As String)
On Error Resume Next
  Dim conn As ADODB.Connection
  Set conn = Manager.GetCustomObjects("refref")
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
  Set conn = Manager.GetCustomObjects("refref")
  If conn.State <> adStateOpen Then
    conn.Open
  End If
  
  
'  Set bzrs = conn.Execute( _
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
    Set locrs = conn.Execute("select * from location where code='" & bzid & "'")
    If Not locrs.EOF Then
      bzid = locrs!id
    End If
  End If
  
  s = "insert into stock(SITE_ID,ITEM_ID,LOCATION_ID,ORDER_ID,QTY_ON_HAND," & _
  "status,UNIT_COST,UOM,LOT_SN,REF_NUM," & _
  "ORD_NUM,PALLET_ID,custom_field1,custom_field6,custom_field11,custom_field5,exp_date,custom_field3,custom_field4,custom_field12,custom_field2,custom_field9,custom_field7)" & _
  "values(" & _
  "1," & Manager.GetIDFromXMLField(CurRow.good_id) & ",'" & bzid & "',null," & MyRound2(netto) & _
   "," & sStatus & ",0,'" & CurRow.edizm & "','" & partia & "','" & QueryCode & "'," & _
  "'" & QueryCode & "'," & palID & "," & MyRound2(LinePal.CaliberQuantity) & ",'" & countryname & "','" & kilplace & "','" & LinePal.made_date & "'," & MakeMSSQLDate(LinePal.exp_date) & ",'" & MyRound2(CurRow.PackageWeight) & "','" & factoryname & "','" & sBrack & "'," & sCaliber & ",'" & LinePal.made_date_to & "','" & Left(LinePal.vetsved, 50) & "') "

  
  Set cmd = New ADODB.Command
  cmd.CommandType = adCmdText
  cmd.CommandText = s
  Set cmd.ActiveConnection = conn
  On Error Resume Next
  cmd.Execute
   If err.Number <> 0 Then
    MsgBox err.Description
  End If
      
  Set loccode = conn.Execute("select code from location where id=" & bzid)
  If Not loccode.EOF Then
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdText
    cmd.CommandText = "update pallet set location_id=" & bzid & " where id=" & palID
    Set cmd.ActiveConnection = conn
    On Error Resume Next
    cmd.Execute
     If err.Number <> 0 Then
      MsgBox err.Description
    End If
  End If
  
  
  Set rs = conn.Execute("select * from RECEIVING_LINE where order_id=" & oid & " and item_id=" & rlID)
  If rs.EOF Then Exit Sub
  
  conn.BeginTrans
  err.Clear
  
  
  cmd.CommandText = "INSERT INTO RECEIVING_HISTORY(custom_field12, [REF_NUMBER], [QTY_REC], [UOM], [LOT_SN], [EXP_DATE], [UNIT_PRICE], [COMMENTS], [REC_DATE], [TRACK_NUMBER2], [TRACK_NUMBER3], [LOCATION], [PALLET], [CONTAINER], [STATUS], [ORDER_ID], [ITEM_ID], [USER_ID], custom_field1,custom_field3,custom_field4,custom_field11,custom_field5,custom_field6,custom_field2,custom_field9,custom_field7)" & _
  "VALUES( '" & sBrack & "','" & QueryCode & "', " & MyRound2(netto) & ",'" & CurRow.edizm & "', '" & partia & "'," & MakeMSSQLDate(LinePal.exp_date) & " , 0, ' ', getdate(), '" & Item.ITTIN_DEF.Item(1).TranspNumber & "', '" & Item.ITTIN_DEF.Item(1).TranspNumber & "','" & LinePal.BufferZonePlace & "','" & poddon.TheNumber & "', '" & Item.ITTIN_DEF.Item(1).Container & "', 0, " & rs!order_id & "," & rs!item_id & ",1," & MyRound2(LinePal.CaliberQuantity) & ",'" & MyRound2(LinePal.PackageWeight) & "','" & factoryname & "','" & kilplace & "','" & LinePal.made_date & "','" & countryname & "'," & sCaliber & ",'" & LinePal.made_date_to & "','" & Left(LinePal.vetsved, 50) & "')"
  Set cmd.ActiveConnection = conn
  err.Clear
  cmd.Execute
  
  If err.Number <> 0 Then
    MsgBox err.Description
  End If
  
  Set rsitem = conn.Execute("select * from item where id =" & rs!item_id)
  
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
    MsgBox err.Description
  End If
  
  
  cmd.CommandText = "update RECEIVING_LINE SET status=1, QTY_ALT_PREV_REC=isnull(QTY_ALT_PREV_REC,0)+" & MyRound2(LinePal.CaliberQuantity) & ", QTY_PREV_REC =isnull(QTY_PREV_REC,0)+" & MyRound2(netto) & ", MADE_DATE=" & MakeMSSQLDate(CurRow.made_date) & ", EXP_DATE=" & MakeMSSQLDate(CurRow.exp_date) & ",PROD_COUNTRY='" & countryname & "',KILL_NUMBER='" & kilplace & "',LOT_SN='" & partia & "' where order_id=" & oid & " and item_ID=" & rlID
  
  err.Clear
  Set cmd.ActiveConnection = conn
  cmd.Execute
  If err.Number <> 0 Then
    MsgBox err.Description
  End If
  
  Dim ordid As String
  
  Dim def As ITTIN_DEF
  Set def = Item.ITTIN_DEF.Item(1)
  ordid = Manager.GetIDFromXMLField(def.QryCode)
  
  cmd.CommandText = "update RECEIVING_ORDER SET status=1, TRACK_NUMBER1='" & def.TranspNumber & "' ,COMMENT1='" & def.Container & "', ZIP='" & def.temp_in_track & "'  where ID=" & ordid
  
  err.Clear
  Set cmd.ActiveConnection = conn
  cmd.Execute
  If err.Number <> 0 Then
    MsgBox err.Description
  End If
  
  
  conn.CommitTrans
End Sub

