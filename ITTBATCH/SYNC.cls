VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SYNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim db As String
Dim server As String
Dim user As String
Dim pass As String
Dim syncdb As String
Dim conn As ADODB.Connection
Dim sconn As ADODB.Connection

Public Sub BeforeSync()

  Dim l As New LOGGER
  l.CreateLog "ITTBATCH.TXT", "ITTBATCH"
  l.Message "BeforSync started"
  
  
  db = GetSetting("ITT", "ITTBATCH", "DB", "")
  syncdb = GetSetting("ITT", "ITTBATCH", "SYNCDB", "")
  server = GetSetting("ITT", "ITTBATCH", "SRV", "")
  user = GetSetting("ITT", "ITTBATCH", "USER", "")
  pass = GetSetting("ITT", "ITTBATCH", "PASS", "")
  
   l.Message "srv=" & server & " db=" & db & " syncdb=" & syncdb
  
  Set conn = New ADODB.Connection
  conn.Provider = "SQLoledb"
  conn.ConnectionString = "Server=" & server & ";DataBase=" & db & ";UID=" & user & ";Pwd=" & pass & ";"
  conn.Open
  If conn.State = ADODB.adStateOpen Then
    
    ' очищаем список заказов
    'conn.Execute "delete from " & syncdb & ".dbo.qrytab"
    
    conn.Execute "update " & syncdb & ".dbo.qrytab set Dirty =-1,RowDirty=-1"
    l.Message "update " & syncdb & ".dbo.qrytab set Dirty =-1,RowDirty=-1"
    
    
    ' добавляем незакрытые заказы
    conn.Execute "delete from " & syncdb & ".dbo.qrytab where uniqueid in ( " & _
    " select RECEIVING_ORDER.ID from RECEIVING_ORDER" & _
    " join partner on RECEIVING_ORDER.partner_id=partner.id" & _
    " Where (Status = 0 Or Status = 1) and rec_date > getdate()-15 and partner.code='" & GetSetting("ITT", "ITTBATCH", "CLI", "") & "')"
    
    l.Message "deleting rows"
    
    conn.Execute "insert into " & syncdb & ".dbo.qrytab(UniqueID,Dirty,RowDirty,QRYNAME) " & _
    " select RECEIVING_ORDER.ID,0,0, convert(varchar(20),number) + ' от '+ convert(varchar(20),rec_date,103) from RECEIVING_ORDER" & _
    " join partner on RECEIVING_ORDER.partner_id=partner.id" & _
    " Where (Status = 0 Or Status = 1) and rec_date > getdate()-15 and partner.code='" & GetSetting("ITT", "ITTBATCH", "CLI", "") & "'"
    
    l.Message "insert open orders"
    
'    ' добавляем записи чобы почистить закрытые заказы
'    conn.Execute "insert into " & syncdb & ".dbo.qrytab(UniqueID,Dirty,QRYNAME)" & _
'    " select RECEIVING_ORDER.ID, -1, convert(varchar(20),number) + ' от '+ convert(varchar(20),ord_date,103) from RECEIVING_ORDER" & _
'    " join partner on RECEIVING_ORDER.partner_id=partner.id" & _
'    " Where (Status = 2)  and ord_date >getdate()-15 and partner.code='" & GetSetting("ITT", "ITTBATCH", "CLI", "") & "'"

    
    
    conn.Close
  End If
  
  
  
  Set conn = Nothing
 
  l.Message "BeforSync stopped"

End Sub

Public Sub AfterSync()

  On Error Resume Next
  
  Dim l As New LOGGER
  l.CreateLog "ITTBATCH.TXT", "ITTBATCH"
  l.Message "AfterSync started"

  db = GetSetting("ITT", "ITTBATCH", "DB", "")
  syncdb = GetSetting("ITT", "ITTBATCH", "SYNCDB", "")
  server = GetSetting("ITT", "ITTBATCH", "SRV", "")
  user = GetSetting("ITT", "ITTBATCH", "USER", "")
  pass = GetSetting("ITT", "ITTBATCH", "PASS", "")

  l.Message "srv=" & server & " db=" & db & " syncdb=" & syncdb

  Set conn = New ADODB.Connection
  conn.Provider = "SQLoledb"
  conn.ConnectionString = "Server=" & server & ";DataBase=" & db & ";UID=" & user & ";Pwd=" & pass & ";"
  conn.Open
  If conn.State = ADODB.adStateOpen Then
  
    Dim rs As ADODB.Recordset

    
    
    ' обработка  приемки товара
    
    Set rs = conn.Execute("select * from " & syncdb & ".dbo.loadtab where Dirty =0")
    If Not rs Is Nothing Then
    
      While Not rs.EOF
      
        l.Message "LoadRecord:" & rs!QRYNUM & " pallet " & rs!pallet
        
        Debug.Print rs!Uniqueid
        
        'If
        Call ProcessLoadRow(rs!QRYNUM & "", rs!pallet & "", rs!ssctop & "", rs!sscbottom & "", UCase(rs!LOT & ""))
        'Then
      '
          conn.Execute "update " & syncdb & ".dbo.loadtab set Dirty =-1,RowDirty =-1 where UniqueID=" & rs!Uniqueid
        'End If
        rs.MoveNext
      Wend
      rs.Close
      Set rs = Nothing
    
    End If
    
    
    
    ' обработка перемещения

    
    Set rs = conn.Execute("select * from " & syncdb & ".dbo.placetab where Dirty =0")
    If Not rs Is Nothing Then
    
      While Not rs.EOF
      
        Debug.Print rs!Uniqueid
        
        l.Message "PlaceRecord:" & rs!pallet & " to " & rs!location
        
        'If
        Call ProcessPlaceRow(UCase(rs!pallet & ""), UCase(rs!location & ""))
        'Then
          conn.Execute "update " & syncdb & ".dbo.placetab set Dirty =-1,RowDirty =-1 where UniqueID=" & rs!Uniqueid
        'End If
        rs.MoveNext
      Wend
      rs.Close
      Set rs = Nothing
    
    End If
    
    Err.Clear
    
    
    
    
    ' обработка инвентаризации
    Set rs = conn.Execute("select * from " & syncdb & ".dbo.invtab where Dirty =0")
    If Not rs Is Nothing Then
    
      While Not rs.EOF
        l.Message "INVentoryRecord: location=" & rs!location & " pallet=" & rs!pallet
        Debug.Print rs!Uniqueid
        'If
        Call ProcessInvRow(rs!pallet, UCase(rs!location), rs!goodq)
        'Then
          conn.Execute "update " & syncdb & ".dbo.invtab set Dirty =-1,RowDirty =-1 where UniqueID=" & rs!Uniqueid
        'End If
        rs.MoveNext
      Wend
      rs.Close
      Set rs = Nothing
    
    End If
    
        
     

    conn.Close
  End If
  
  
  
  Set conn = Nothing
  l.Message "AfterSync stopped"
End Sub



Public Sub Setup()
  Dim f As frmSetup
  Set f = New frmSetup
  f.Show vbModal
  Unload f
  Set f = Nothing
End Sub

Private Function ProcessLoadRow(qryname As String, pallet As String, ssctop As String, sscbottom As String, LOT As String) As Boolean
  
  On Error Resume Next
  
  Dim l As New LOGGER
  Dim rsq As ADODB.Recordset
  Dim rsm As ADODB.Recordset
  Dim rsp As ADODB.Recordset
  Dim rsl As ADODB.Recordset
  Dim rsitem As ADODB.Recordset
  
  Dim orderID As String
  Dim orderNum As String
  Dim ItemID As String
  Dim ItemCode As String
  Dim ItemDesc As String
  
  Dim AltCode As String
  Dim GroupCode As String
  Dim BoxQ As String
  Dim sExpDate As String
  Dim sPalDate As String
  Dim SSC As String
  Dim mult As Double
  Dim UOM As String
  Dim qs As String
  Dim palletid As String
  Dim site_id As String
  Dim dpaldate As Date
  Dim po_number As String
  Dim itemvoltage As Long
  
  
  l.CreateLog "ittbatch.txt", "ITTBATCH"
  l.Message "1"
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    Err.Clear
  End If
  
  ' находим товар по коду из SSC
  ssctop = Replace(ssctop, Chr(29), "")
  sscbottom = Replace(sscbottom, Chr(29), "")
    
  l.Message "2"
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    Err.Clear
  End If
  
  'mid("020500018607900037010890123456",3,14)
  GroupCode = Mid(ssctop, 3, 14)
  BoxQ = Mid(ssctop, 19, 4)
  sExpDate = Mid(ssctop, 25, 6) 'yymmdd
  
  l.Message "3"
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    Err.Clear
  End If
  
  'mid("00350001861030006654130105259300123456",3,14)
  SSC = Mid(sscbottom, 3, 18)
  sPalDate = Mid(sscbottom, 23, 6)  'yymmdd
  AltCode = Mid(sscbottom, 31, 8)
  
  l.Message "4"
  
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    Err.Clear
  End If

  
  ' находим заказ
  ProcessLoadRow = False
  
  Set rsq = conn.Execute("select RECEIVING_ORDER.ID,RECEIVING_ORDER.number,po_number,site_id from RECEIVING_ORDER where status<>2 and  convert(varchar(20),number) + ' от '+ convert(varchar(20),rec_date,103) ='" & qryname & "'")
  If Not rsq Is Nothing Then
      If Not rsq.EOF Then
               orderID = rsq!id
               orderNum = rsq!Number
               site_id = rsq!site_id
               po_number = rsq!po_number
      Else
        l.Message "Error: " & qryname & " not found or closed"
        Exit Function
      End If
  Else
    Exit Function
  End If

  l.Message "5"
  
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    Err.Clear
  End If


  Dim i As Integer
  Dim ac2 As String
  ac2 = ""
  For i = 1 To Len(AltCode)
    If Mid(AltCode, i, 1) <> "0" Or ac2 <> "" Then
    ac2 = ac2 & Mid(AltCode, i, 1)
    End If
  Next
  AltCode = ac2
  
  Set rsitem = conn.Execute("select item.id,item.code, item.description, item.voltage from item join partner on item.class= partner.code and partner.code='" & GetSetting("ITT", "ITTBATCH", "CLI", "") & "' where ALT_CODE='" & AltCode & "'")
  If Not rsitem Is Nothing Then
    If Not rsitem.EOF Then
      ItemID = rsitem!id
      ItemCode = rsitem!code
      ItemDesc = rsitem!Description
      itemvoltage = rsitem!voltage
      
      
      ' расчитываем количество штук
      Set rsm = conn.Execute("select u1.MULTIPLIER,item.id from item join UOM U1 on U1.code=item.sales_uom and item.UOM_subclass_id=u1.uom_subclass_id Where Item.id = " & ItemID)
      mult = 1
      If Not rsm Is Nothing Then
        If Not rsm.EOF Then
          mult = rsm!MULTIPLIER
        End If
      End If
      
      UOM = "шт."
      
    Else
      l.Message "Error: Item ALT_code='" & AltCode & "' not found"
      Exit Function
    End If
    
    l.Message "6"
  
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    Err.Clear
  End If
  
  Else
    Exit Function
  End If
  
  Set rsp = conn.Execute("select * from pallet where code ='" & pallet & "'")
  If Not rsp Is Nothing Then
    If Not rsp.EOF Then
      l.Message "Error: Pallet already registred :" & pallet & "."
      Exit Function
    End If
  End If
  
  'conn.BeginTrans
  qs = ""
  qs = "insert into pallet(CODE,SITE_ID,LOCATION_ID) values('" & pallet & "'," & site_id & ",null)"
  conn.Execute (qs)
  l.Message qs
  
  Set rsp = conn.Execute("select * from pallet where code ='" & pallet & "'")
  If rsp Is Nothing Then
    l.Message "Error: Pallet " & pallet & " not found"
    'conn.RollbackTrans
    Exit Function
  End If
  If rsp.EOF Then
    l.Message "Error: Pallet " & pallet & " not found"
    'conn.RollbackTrans
    Exit Function
  End If
  
  palletid = rsp!id
  
  
  ' проверяем нет ли уже в заказе строки для этой паллеты
  Set rsl = conn.Execute("select * from receiving_history where order_id=" & orderID & " and pallet='" & pallet & "'")
  
  If Not rsl.EOF Then
    l.Message "Error: Pallet " & pallet & " already exists in order :" & qryname
    'conn.RollbackTrans
    Exit Function
  End If
  
  
  conn.Execute ("update RECEIVING_ORDER set status=1 where id=" & orderID)
  
  
  ' находим ( или добавляем) строку в заказ
  Set rsl = conn.Execute("select * from receiving_line where order_id=" & orderID & " and item_id=" & ItemID)
  If Not rsl Is Nothing Then
    If rsl.EOF Then
      qs = ""
      qs = qs & "INSERT INTO [RECEIVING_LINE]( [ORDER_ID], [STATUS], [ITEM_ID], [QTY_ORD], [QTY_PREV_REC], [UOM],  [UNIT_PRICE],  [KILL_NUMBER], [AMOUNT], [COMMENTS],  [PRODUCER])"
      qs = qs & "VALUES("
      qs = qs & orderID & "," '<ORDER_ID,int,>,
      qs = qs & 1 & "," '<STATUS,int,>,
      qs = qs & ItemID & "," '<ITEM_ID,int,>,
      qs = qs & NoComa(Val(BoxQ) * mult) & "," '<QTY_ORD,decimal(12,2),>,
      qs = qs & NoComa(Val(BoxQ) * mult) & "," '<QTY_PREV_REC,decimal(12,2),>,
      qs = qs & "'" & UOM & "'," '<UOM,nvarchar(50),>,
      qs = qs & 0 & "," '<UNIT_PRICE,decimal(12,4),>,
      qs = qs & "''" & "," '<KILL_NUMBER,nvarchar(100),>,
      qs = qs & NoComa(Val(BoxQ) * mult) & ","  '<AMOUNT,decimal(12,4),>,
      qs = qs & "''" & "," '<COMMENTS,nvarchar(255),>,
      qs = qs & "''"  '<PRODUCER,nvarchar(100),>
      qs = qs & ")"
      Err.Clear
      l.Message qs
      conn.Execute qs
      If Err.Number <> 0 Then
        l.Message "Error:" & Err.Description
        'conn.RollbackTrans
        Exit Function
      End If
    
    Else
      qs = ""
      qs = qs & "Update [RECEIVING_LINE]"
      qs = qs & "  SET  [ORDER_ID]=" & orderID & ","  '<ORDER_ID,int,>,
      
      If rsl!QTY_ORD = rsl!QTY_PREV_REC + Val(BoxQ) * mult Then
        qs = qs & "   [STATUS]= 2,"   '<STATUS,int,>,"
      Else
        qs = qs & "   [STATUS]= 1,"   '<STATUS,int,>,"
      End If
      
      qs = qs & "   [ITEM_ID]=" & ItemID & "," '<ITEM_ID,int,>,"
      'qs = qs & "   [QTY_ORD]= QTY_ORD + " & (Val(BoxQ) * mult) & ","  '<QTY_ORD,decimal(12,2),>,
      qs = qs & "   [QTY_PREV_REC]= QTY_PREV_REC + " & NoComa(Val(BoxQ) * mult) & ","  '<QTY_PREV_REC,decimal(12,2),>,
      qs = qs & "   [UOM]= '" & UOM & "'" '<UOM,nvarchar(50),>,"
      qs = qs & "   WHERE ID=" & rsl!id
      
      Err.Clear
      l.Message qs
      conn.Execute qs
      If Err.Number <> 0 Then
        l.Message "Error:" & Err.Description
        'conn.RollbackTrans
        Exit Function
      End If
      
      
     
    End If
  
  Else
    If Err.Number <> 0 Then
      l.Message "Error:" & Err.Description
    End If
    'conn.RollbackTrans
    Exit Function
  End If
  
  
  
  ' добавляем строку в историю заказа
  
      qs = ""
  
      qs = qs & "INSERT INTO [RECEIVING_HISTORY]"
      qs = qs & "( [REF_NUMBER], [QTY_REC], [UOM], [LOT_SN], [EXP_DATE], [UNIT_PRICE], [COMMENTS], [REC_DATE], [TRACK_NUMBER2], [TRACK_NUMBER3], [LOCATION], [PALLET], [CONTAINER], [STATUS], [ORDER_ID], [ITEM_ID], [USER_ID], [CUSTOM_FIELD1], [CUSTOM_FIELD2], [CUSTOM_FIELD3], [CUSTOM_FIELD4], [CUSTOM_FIELD5], [CUSTOM_FIELD6], [CUSTOM_FIELD7], [CUSTOM_FIELD8], [CUSTOM_FIELD9], [CUSTOM_FIELD10], [CUSTOM_FIELD11], [CUSTOM_FIELD12], [CUSTOM_FIELD13], [CUSTOM_FIELD14], [CUSTOM_FIELD15], [CUSTOM_FIELD16])"
      qs = qs & "VALUES("
      qs = qs & "'" & orderNum & "'," '<REF_NUMBER,nvarchar(50),>,
      qs = qs & (Val(BoxQ) * mult) & "," '<QTY_REC,numeric(12,2),>,
      qs = qs & "'" & UOM & "',"  '<UOM,nvarchar(50),>,
      qs = qs & "'" & LOT & "',"  '<LOT_SN,nvarchar(255),>,
      If sExpDate <> "000000" Then
        qs = qs & "convert(datetime,'" & sExpDate & "',112)," '<EXP_DATE,datetime,>,
      Else
        qs = qs & "dateadd(Day," & itemvoltage & ",convert(datetime,'" & sPalDate & "',112))," '<EXP_DATE,datetime,>,
        'qs = qs & "null," '<EXP_DATE,datetime,>,
      End If
      qs = qs & "0," '<UNIT_PRICE,numeric(12,4),>,
      qs = qs & "''," '<COMMENTS,nvarchar(255),>,
      qs = qs & "getdate()," '<REC_DATE,datetime,>,
      qs = qs & "''," '<TRACK_NUMBER2,nvarchar(50),>,
      qs = qs & "''," '<TRACK_NUMBER3,nvarchar(50),>,
      qs = qs & "''," '<LOCATION,nvarchar(50),>,
      qs = qs & "'" & pallet & "'," '<PALLET,nvarchar(255),>,
      qs = qs & "''," '<CONTAINER,nvarchar(255),>,
      qs = qs & "1," '<STATUS,int,>,
      qs = qs & orderID & "," '<ORDER_ID,int,>,
      qs = qs & ItemID & "," '<ITEM_ID,int,>
      qs = qs & 1 & "," ' <USER_ID,int,>,
      qs = qs & "''," '<CUSTOM_FIELD1,nvarchar(100),>,
      qs = qs & "'" & SSC & "'," '<CUSTOM_FIELD2,nvarchar(100),>,
      qs = qs & "''," '<CUSTOM_FIELD3,nvarchar(100),>,
      qs = qs & "'" & ssctop & "|" & sscbottom & "',"   '<CUSTOM_FIELD4,nvarchar(100),>,
      qs = qs & "''," '<CUSTOM_FIELD5,nvarchar(100),>,
      qs = qs & "''," '<CUSTOM_FIELD6,nvarchar(100),>,
      qs = qs & "''," '<CUSTOM_FIELD7,nvarchar(100),>,
      qs = qs & "''," '<CUSTOM_FIELD8,nvarchar(100),>,
      qs = qs & "''," '<CUSTOM_FIELD9,nvarchar(100),>,
      qs = qs & "''," '<CUSTOM_FIELD10,nvarchar(100),>,
      qs = qs & "''," '<CUSTOM_FIELD11,nvarchar(100),>,
      qs = qs & "''," '<CUSTOM_FIELD12,nvarchar(100),>,
      qs = qs & "''," '<CUSTOM_FIELD13,nvarchar(100),>,
      qs = qs & "''," '<CUSTOM_FIELD14,nvarchar(100),>,
      qs = qs & "''," '<CUSTOM_FIELD15,nvarchar(100),>,
      qs = qs & "''" '<CUSTOM_FIELD16,nvarchar(100),>
      qs = qs & ")"
      Err.Clear
      On Error Resume Next
      l.Message qs
      conn.Execute qs
      If Err.Number <> 0 Then
        l.Message "Error:" & Err.Description
        'conn.RollbackTrans
        Exit Function
      End If
    
  ' добавляем строку в history ?
  
    
    qs = ""
    
    qs = qs & "INSERT INTO[HISTORY]( [CODE], [STAMP], [ITEM], [DESCRIPTION], [LOT_SN], [EXP_DATE], [UNIT_COST], [QTY_ON_HAND], [QTY], [UOM], [STATUS], [LOCATION], [REF_NUM], [ORD_NUM], [USER_NAME], [SITE_ID], [PALLET], [CONTAINER], [CUSTOM_FIELD1], [CUSTOM_FIELD2], [CUSTOM_FIELD3], [CUSTOM_FIELD4], [CUSTOM_FIELD5], [CUSTOM_FIELD6], [CUSTOM_FIELD7], [CUSTOM_FIELD8], [CUSTOM_FIELD9], [CUSTOM_FIELD10], [CUSTOM_FIELD11], [CUSTOM_FIELD12], [CUSTOM_FIELD13], [CUSTOM_FIELD14], [CUSTOM_FIELD15], [CUSTOM_FIELD16])"
    qs = qs & "VALUES("
    qs = qs & "6," '<CODE,int,>,
    qs = qs & "getdate()," '<STAMP,datetime,>,
    qs = qs & "'" & ItemCode & "'," '<ITEM,nvarchar(50),>,
    qs = qs & "'" & ItemDesc & "'," '<DESCRIPTION,nvarchar(255),>,
    qs = qs & "'" & LOT & "'," ' <LOT_SN,nvarchar(50),>,
     If sExpDate <> "000000" Then
      qs = qs & "convert(datetime,'" & sExpDate & "',112)," '<EXP_DATE,datetime,>,
    Else
      qs = qs & "dateadd(Day," & itemvoltage & ",convert(datetime,'" & sPalDate & "',112))," '<EXP_DATE,datetime,>,
    End If
'    qs = qs & "convert(datetime,'" & sExpDate & "',112),"  '<EXP_DATE,datetime,>,
    qs = qs & "0," '<UNIT_COST,decimal(12,4),>,
    qs = qs & "0," '<QTY_ON_HAND,decimal(12,2),>,
    qs = qs & NoComa(Val(BoxQ) * mult) & "," '<QTY,decimal(12,2),>,
    qs = qs & "'" & UOM & "'," '<UOM,nvarchar(50),>,
    qs = qs & "'Наличие'," '<STATUS,nvarchar(50),>,
    qs = qs & "''," '<LOCATION,nvarchar(50),>,
    qs = qs & "'" & po_number & "'," '<REF_NUM,nvarchar(50),>,
    qs = qs & "'" & orderNum & "'," '<ORD_NUM,nvarchar(50),>,
    qs = qs & "'BATCH TERMINAL'," '<USER_NAME,nvarchar(50),>,
    qs = qs & site_id & "," '<SITE_ID,int,>,
    qs = qs & "'" & pallet & "'," '<PALLET,nvarchar(255),>,
    qs = qs & "''," '<CONTAINER,nvarchar(255),>,
    qs = qs & "''," '<CUSTOM_FIELD1,nvarchar(100),>,
    qs = qs & "'" & SSC & "'," '<CUSTOM_FIELD2,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD3,nvarchar(100),>,
    qs = qs & "'" & ssctop & "|" & sscbottom & "'," '<CUSTOM_FIELD4,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD5,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD6,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD7,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD8,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD9,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD10,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD11,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD12,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD13,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD14,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD15,nvarchar(100),>,
    qs = qs & "''" '<CUSTOM_FIELD16,nvarchar(100),>
    qs = qs & ")"
    Err.Clear
    l.Message qs
    conn.Execute qs
    If Err.Number <> 0 Then
      l.Message "Error:" & Err.Description
      'conn.RollbackTrans
      Exit Function
    End If

  ' добавляем строку в stock ?
   
   
    qs = ""
    qs = qs & "INSERT INTO [STOCK]( [SITE_ID], [ITEM_ID], [LOCATION_ID], [UOM], [LOT_SN], [EXP_DATE], [UNIT_COST], [STATUS], [REF_NUM], [ORD_NUM], [QTY_ON_HAND], [PALLET_ID], [CONTAINER_ID], [CUSTOM_FIELD1], [CUSTOM_FIELD2], [CUSTOM_FIELD3], [CUSTOM_FIELD4], [CUSTOM_FIELD5], [CUSTOM_FIELD6], [CUSTOM_FIELD7], [CUSTOM_FIELD8], [CUSTOM_FIELD9], [CUSTOM_FIELD10], [CUSTOM_FIELD11], [CUSTOM_FIELD12], [CUSTOM_FIELD13], [CUSTOM_FIELD14], [CUSTOM_FIELD15], [CUSTOM_FIELD16])"
    qs = qs & " VALUES("
    qs = qs & site_id & "," '<SITE_ID,int,>,
    qs = qs & ItemID & "," '<ITEM_ID,int,>,
    qs = qs & "null ," '<LOCATION_ID,int,>,
    qs = qs & "'" & UOM & "'," '<UOM,nvarchar(50),>,
    qs = qs & "'" & LOT & "'," '<LOT_SN,nvarchar(50),>,
        If sExpDate <> "000000" Then
      qs = qs & "convert(datetime,'" & sExpDate & "',112)," '<EXP_DATE,datetime,>,
    Else
      qs = qs & "dateadd(Day," & itemvoltage & ",convert(datetime,'" & sPalDate & "',112))," '<EXP_DATE,datetime,>,
    End If
    'qs = qs & "convert(datetime,'" & sExpDate & "',112)," '<EXP_DATE,datetime,>,
    qs = qs & "0," '<UNIT_COST,decimal(12,4),>,
    qs = qs & "0," '<STATUS,int,>,
    'qs = qs & "'" & po_number & "'," '<REF_NUM,nvarchar(50),>,
    'qs = qs & "'" & orderNum & "',"  '<ORD_NUM,nvarchar(50),>,
    qs = qs & "''," '<REF_NUM,nvarchar(50),>,
    qs = qs & "'',"  '<ORD_NUM,nvarchar(50),>,
    qs = qs & NoComa(Val(BoxQ) * mult) & "," '<QTY_ON_HAND,decimal(12,2),>,
    qs = qs & palletid & "," '<PALLET_ID,int,>,
    qs = qs & "null," '<CONTAINER_ID,int,>,
    qs = qs & "''," '<CUSTOM_FIELD1,nvarchar(100),>,
    qs = qs & "'" & SSC & "'," '<CUSTOM_FIELD2,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD3,nvarchar(100),>,
    qs = qs & "'" & ssctop & "|" & sscbottom & "'," '<CUSTOM_FIELD4,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD5,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD6,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD7,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD8,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD9,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD10,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD11,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD12,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD13,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD14,nvarchar(100),>,
    qs = qs & "''," '<CUSTOM_FIELD15,nvarchar(100),>,
    qs = qs & "''" '<CUSTOM_FIELD16,nvarchar(100),>,
    qs = qs & ")"
    Err.Clear
    l.Message qs
    conn.Execute qs
    If Err.Number <> 0 Then
      l.Message "Error:" & Err.Description
      'conn.RollbackTrans
      Exit Function
    End If

    'conn.CommitTrans
   
    ProcessLoadRow = True
End Function

Private Function ProcessInvRow(pallet As String, place As String, Quantity As Long) As Boolean
  ProcessInvRow = True
End Function

Private Function ProcessPlaceRow(pallet As String, place As String) As Boolean
  On Error Resume Next
  Dim rspl As ADODB.Recordset
  Dim rsstok As ADODB.Recordset
  Dim rsloc As ADODB.Recordset
  Dim rsitem As ADODB.Recordset
  ProcessPlaceRow = False
  Dim palletid As String
  Dim location As String
  Dim locationid As String
  Dim oldlocation As String
  Dim oldlocationid As String
  Dim qs As String
  Dim siteid As String
  Dim l As New LOGGER
  
  
  l.CreateLog "ittbatch.txt", "ITTBATCH"
  
  
  l.Message "P1"
  
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    Err.Clear
  End If

   
  Set rspl = conn.Execute("select * from pallet where code='" & pallet & "'")
  If rspl Is Nothing Then
    l.Message "Pallet " & pallet & " not found"
    Exit Function
  End If
  
  l.Message "P2"
  
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    Err.Clear
  End If

  
  If rspl.EOF Then
    l.Message "Pallet " & pallet & " not found"
    Exit Function
  End If
  
  l.Message "P3"
  
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    Err.Clear
  End If

  
  palletid = rspl!id
  
  l.Message "P4"
  
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    Err.Clear
  End If

  
  ' находим поддон в стоке
  
  Set rsstok = conn.Execute("select *,convert(varchar(20),exp_date,112)as myed from stock where pallet_id=" & palletid & " and status=0")
  
  If rsstok Is Nothing Then
    l.Message "Pallet " & pallet & " not found at stock"
    Exit Function
  End If
  
  If rsstok.EOF Then
    l.Message "Pallet " & pallet & " not found at stock"
    Exit Function
  End If
  
  
  l.Message "P5"
  
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    Err.Clear
  End If

  
  ' определяем  его  текущее положение
  
  
  oldlocationid = rsstok!location_id & ""
  If oldlocationid <> "" Then
  Set rsloc = conn.Execute("select * from location where id = " & oldlocationid)
    If Not rsloc.EOF Then
      oldlocation = rsloc!code
    Else
      oldlocation = ""
    End If
  Else
  oldlocation = ""
  End If
  
  l.Message "P6"
  
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    Err.Clear
  End If
  
  
  Set rsitem = conn.Execute("select * from item where id=" & rsstok!ITEM_ID)
  
  l.Message "P7"
  
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    Err.Clear
  End If
  
  If rsitem.EOF Then
      l.Message "Record for ITEMid=" & rsstok!ITEM_ID & " not found in database"
      Exit Function
  End If
  
  
  
  ' находим новую ячейку
  
  Set rsloc = conn.Execute("select * from location where code ='" & place & "'")
  If rsloc Is Nothing Then
      l.Message "Location " & place & " not found in database"
      Exit Function
  End If
  
  If rsloc.EOF Then
      l.Message "Location " & place & " not found in database"
      Exit Function
  End If
  
  locationid = rsloc!id
  location = place
  siteid = rsloc!site_id
  
  'conn.BeginTrans
  
  ' делаем перемещение с
   qs = ""
   qs = qs & "INSERT INTO[HISTORY]( [CODE], [STAMP], [ITEM], [DESCRIPTION], [LOT_SN], [EXP_DATE], [UNIT_COST], [QTY_ON_HAND], [QTY], [UOM], [STATUS], [LOCATION], [REF_NUM], [ORD_NUM], [USER_NAME], [SITE_ID], [PALLET], [CONTAINER], [CUSTOM_FIELD1], [CUSTOM_FIELD2], [CUSTOM_FIELD3], [CUSTOM_FIELD4], [CUSTOM_FIELD5], [CUSTOM_FIELD6], [CUSTOM_FIELD7], [CUSTOM_FIELD8], [CUSTOM_FIELD9], [CUSTOM_FIELD10], [CUSTOM_FIELD11], [CUSTOM_FIELD12], [CUSTOM_FIELD13], [CUSTOM_FIELD14], [CUSTOM_FIELD15], [CUSTOM_FIELD16])"
   qs = qs & "VALUES("
   qs = qs & "3," '<CODE,int,>,
   qs = qs & "getdate()," '<STAMP,datetime,>,
   qs = qs & "'" & rsitem!code & "'," '<ITEM,nvarchar(50),>,
   qs = qs & "'" & rsitem!Description & "'," '<DESCRIPTION,nvarchar(255),>,
   qs = qs & "'" & rsstok!lot_SN & "'," ' <LOT_SN,nvarchar(50),>,
   If rsstok!exp_date & "" = "" Then
     qs = qs & "getdate()," '<EXP_DATE,datetime,>,
   Else
     qs = qs & "convert(datetime, '" & rsstok!myed & "',112)," '<EXP_DATE,datetime,>,
   End If
   
   
   qs = qs & NoComa(rsstok!Unit_Cost) & ","  '<UNIT_COST,decimal(12,4),>,
  qs = qs & NoComa(rsstok!QTY_ON_HAND) & "," '<QTY_ON_HAND,decimal(12,2),>,
  qs = qs & NoComa(-rsstok!QTY_ON_HAND) & "," '<QTY,decimal(12,2),>,
  qs = qs & "'" & rsstok!UOM & "'," '<UOM,nvarchar(50),>,
  qs = qs & "'Наличие'," '<STATUS,nvarchar(50),>,
  qs = qs & "'" & oldlocation & "'," '<LOCATION,nvarchar(50),>,
  qs = qs & "'" & rsstok!REF_NUM & "'," '<REF_NUM,nvarchar(50),>,
  qs = qs & "'" & rsstok!ORD_NUM & "'," '<ORD_NUM,nvarchar(50),>,
  qs = qs & "'BATCH TERMINAL'," '<USER_NAME,nvarchar(50),>,
  qs = qs & "'" & rsstok!site_id & "'," '<SITE_ID,int,>,
  qs = qs & "'" & pallet & "'," '<PALLET,nvarchar(255),>,
  qs = qs & "'" & "" & "'," '<CONTAINER,nvarchar(255),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD1 & "'," '<CUSTOM_FIELD1,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD2 & "'," '<CUSTOM_FIELD2,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD3 & "'," '<CUSTOM_FIELD3,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD4 & "'," '<CUSTOM_FIELD4,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD5 & "'," '<CUSTOM_FIELD5,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD6 & "'," '<CUSTOM_FIELD6,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD7 & "'," '<CUSTOM_FIELD7,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD8 & "'," '<CUSTOM_FIELD8,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD9 & "'," '<CUSTOM_FIELD9,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD10 & "'," '<CUSTOM_FIELD10,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD11 & "'," '<CUSTOM_FIELD11,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD12 & "'," '<CUSTOM_FIELD12,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD13 & "'," '<CUSTOM_FIELD13,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD14 & "'," '<CUSTOM_FIELD14,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD15 & "'," '<CUSTOM_FIELD15,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD16 & "'" '<CUSTOM_FIELD16,nvarchar(100),>
  qs = qs & ")"
  Err.Clear
  On Error Resume Next
  l.Message qs
  conn.Execute (qs)
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    'conn.RollbackTrans
    Exit Function
  End If
    
  
  ' делаем перемещение на

  qs = ""
  
  qs = ""
  qs = qs & "INSERT INTO[HISTORY]( [CODE], [STAMP], [ITEM], [DESCRIPTION], [LOT_SN], [EXP_DATE], [UNIT_COST], [QTY_ON_HAND], [QTY], [UOM], [STATUS], [LOCATION], [REF_NUM], [ORD_NUM], [USER_NAME], [SITE_ID], [PALLET], [CONTAINER], [CUSTOM_FIELD1], [CUSTOM_FIELD2], [CUSTOM_FIELD3], [CUSTOM_FIELD4], [CUSTOM_FIELD5], [CUSTOM_FIELD6], [CUSTOM_FIELD7], [CUSTOM_FIELD8], [CUSTOM_FIELD9], [CUSTOM_FIELD10], [CUSTOM_FIELD11], [CUSTOM_FIELD12], [CUSTOM_FIELD13], [CUSTOM_FIELD14], [CUSTOM_FIELD15], [CUSTOM_FIELD16])"
  qs = qs & "VALUES("
  qs = qs & "4," '<CODE,int,>,
  qs = qs & "getdate()," '<STAMP,datetime,>,
  qs = qs & "'" & rsitem!code & "'," '<ITEM,nvarchar(50),>,
  qs = qs & "'" & rsitem!Description & "'," '<DESCRIPTION,nvarchar(255),>,
  qs = qs & "'" & rsstok!lot_SN & "'," ' <LOT_SN,nvarchar(50),>,
  If rsstok!exp_date & "" = "" Then
    qs = qs & "getdate()," '<EXP_DATE,datetime,>,
  Else
    qs = qs & "convert(datetime, '" & rsstok!myed & "',112)," '<EXP_DATE,datetime,>,
  End If
  qs = qs & NoComa(rsstok!Unit_Cost) & ","  '<UNIT_COST,decimal(12,4),>,
  qs = qs & NoComa(rsstok!QTY_ON_HAND) & "," '<QTY_ON_HAND,decimal(12,2),>,
  qs = qs & 0 & "," '<QTY,decimal(12,2),>,
  qs = qs & "'" & rsstok!UOM & "'," '<UOM,nvarchar(50),>,
  qs = qs & "'Наличие'," '<STATUS,nvarchar(50),>,
  qs = qs & "'" & location & "'," '<LOCATION,nvarchar(50),>,
  qs = qs & "'" & rsstok!REF_NUM & "'," '<REF_NUM,nvarchar(50),>,
  qs = qs & "'" & rsstok!ORD_NUM & "'," '<ORD_NUM,nvarchar(50),>,
  qs = qs & "'BATCH TERMINAL'," '<USER_NAME,nvarchar(50),>,
  qs = qs & "'" & siteid & "'," '<SITE_ID,int,>,
  qs = qs & "'" & pallet & "'," '<PALLET,nvarchar(255),>,
  qs = qs & "'" & "" & "'," '<CONTAINER,nvarchar(255),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD1 & "'," '<CUSTOM_FIELD1,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD2 & "'," '<CUSTOM_FIELD2,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD3 & "'," '<CUSTOM_FIELD3,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD4 & "'," '<CUSTOM_FIELD4,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD5 & "'," '<CUSTOM_FIELD5,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD6 & "'," '<CUSTOM_FIELD6,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD7 & "'," '<CUSTOM_FIELD7,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD8 & "'," '<CUSTOM_FIELD8,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD9 & "'," '<CUSTOM_FIELD9,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD10 & "'," '<CUSTOM_FIELD10,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD11 & "'," '<CUSTOM_FIELD11,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD12 & "'," '<CUSTOM_FIELD12,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD13 & "'," '<CUSTOM_FIELD13,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD14 & "'," '<CUSTOM_FIELD14,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD15 & "'," '<CUSTOM_FIELD15,nvarchar(100),>,
  qs = qs & "'" & rsstok!CUSTOM_FIELD16 & "'" '<CUSTOM_FIELD16,nvarchar(100),>
  qs = qs & ")"
  Err.Clear
  l.Message qs
  conn.Execute (qs)
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    'conn.RollbackTrans
    Exit Function
  End If
  
  ' обновляем ячейку в стоке
  qs = ""
  qs = qs & "Update [STOCK]"
  qs = qs & " SET"
  qs = qs & " [LOCATION_ID]=" & locationid
  qs = qs & ",[SITE_ID]=" & siteid
  qs = qs & " WHERE [PALLET_ID]=" & palletid & " and status=0"
  
  Err.Clear
  l.Message qs
  conn.Execute (qs)
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    'conn.RollbackTrans
    Exit Function
  End If
  
  
  ' обновляем данные в паллете
  qs = ""
  qs = qs & "Update [PALLET]"
  qs = qs & " SET"
  qs = qs & " [LOCATION_ID]=" & locationid
  qs = qs & " ,[SITE_ID]=" & siteid
  qs = qs & " WHERE [ID]=" & palletid
  
  Err.Clear
  l.Message qs
  conn.Execute (qs)
  If Err.Number <> 0 Then
    l.Message "Error:" & Err.Description
    'conn.RollbackTrans
    Exit Function
  End If
  
  'conn.CommitTrans

    ProcessPlaceRow = True
End Function

Public Function NoComa(ByVal s As Variant) As String
  Dim e As String, out As Double
  e = s
  NoComa = Replace(e, ",", ".")
End Function
