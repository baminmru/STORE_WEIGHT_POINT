VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BEFORESTATUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim theItem As ITTSSCC.Application
Dim Manager As MTZManager.Main
Dim session As MTZSession.session
Dim conn As ADODB.Connection
  





Public Function Check(Item As Object, ByVal NewStatusid As String, ByVal CurrentUser As Object, ByVal Module As String) As Boolean
  Dim result As Boolean
  Dim ID As String
  Dim iop As ITTSSCC.Application
  result = True
  
  On Error Resume Next
  Set theItem = Item
  Set Manager = theItem.Manager
  Set session = theItem.MTZSession
  
 
  Dim csstr As String
  Dim i As Long
  Dim rtype As ITTD.ITTD_RULE
  

 
' состояния для типа:ITTOPT Задание на перемещения
' "{0A7FC795-E787-4D17-9689-96EFFF8F0D9D}" 'Задание исполнено
' "{300483B2-1D94-4A33-8ADF-ABF32E72E57B}" 'Оформлено
' "{C861FA15-0DF6-42D4-BCE9-2B38C3E6C0CB}" 'Оформляется
  
 
  
  
'    If NewStatusid = "{0A7FC795-E787-4D17-9689-96EFFF8F0D9D}" Then
'        Set conn = Manager.GetCustomObjects("refref")
'
'        For i = 1 To theItem.ITTOPT_MOVE.Count
'          If theItem.ITTOPT_MOVE.Item(i).ThePalletteNum <> "" Then
'
'              result = result And RegisterMove(theItem.ITTOPT_MOVE.Item(i))
''              If Not rsult Then
''                Exit For
''              End If
'          End If
'
'
'        Next
'    End If
    
    
  Check = result
End Function

'Private Function RegisterMove(ByRef MD As ITTOPT_MOVE) As Boolean
'  Dim rs As ADODB.Recordset
'  Dim strs As ADODB.Recordset
'  Dim locfrs As ADODB.Recordset
'  Dim loctrs As ADODB.Recordset
'  Dim palrs As ADODB.Recordset
'  Dim itemrs As ADODB.Recordset
'
'  Dim cmd As ADODB.Command
'  RegisterMove = True
'
'  If MD.ISMoved = Boolean_Net Then
'    Set rs = conn.Execute("select * from v_bami_stock where loc_code='" & MD.FromLocation & "' and pallet_code=" & MD.ThePalletteNum)
'    If Not rs.EOF Then
'      Set strs = conn.Execute("select * from stock where id=" & rs!Stock_ID)
'      Set loctrs = conn.Execute("select * from location where code='" & MD.ToLocation & "'")
'      Set locfrs = conn.Execute("select * from location where code='" & MD.FromLocation & "'")
'      Set itemrs = conn.Execute("select * from item where id=" & strs!item_id)
'      Set palrs = conn.Execute("select * from pallet where id=" & strs!pallet_id)
'
'
'      ' перемещение с
'      Set cmd = New ADODB.Command
'      cmd.CommandText = "INSERT INTO HISTORY(" & _
'      "stamp,Code , Item, Description, LOT_SN, EXP_DATE, UNIT_COST, QTY_ON_HAND, QTY, UOM, Status," & _
'      "LOCATION , REF_NUM, ORD_NUM, USER_NAME, SITE_ID, PALLET, Container," & _
'      "CUSTOM_FIELD1 , CUSTOM_FIELD2, CUSTOM_FIELD3, CUSTOM_FIELD4, CUSTOM_FIELD5, CUSTOM_FIELD6, CUSTOM_FIELD7, CUSTOM_FIELD8," & _
'      "CUSTOM_FIELD9, CUSTOM_FIELD10, CUSTOM_FIELD11, CUSTOM_FIELD12, CUSTOM_FIELD13, CUSTOM_FIELD14, CUSTOM_FIELD15, CUSTOM_FIELD16 " & _
'      ")VALUES( " & _
'      "getdate(),3,'" & itemrs!Code & "','" & itemrs!Description & "','" & strs!LOT_SN & "'," & MakeMSSQLDate(strs!exp_date) & "," & strs!Unit_COST & "," & MyRound2(strs!QTY_ON_HAND) & ",-" & MyRound2(strs!QTY_ON_HAND) & ",'" & strs!UOM & "',0," & _
'      "'" & locfrs!Code & "','" & strs!ref_num & "','" & strs!ord_num & "','sa',1,'" & palrs!Code & "',''," & _
'      "'-" & MyRound2(strs!custom_field1) & "','" & strs!custom_field2 & "','" & strs!custom_field3 & "','" & strs!custom_field4 & "','" & strs!custom_field5 & "','" & strs!custom_field6 & "','" & strs!custom_field7 & "','" & strs!custom_field8 & "'," & _
'      "'" & strs!custom_field9 & "','" & strs!custom_field10 & "','" & strs!custom_field11 & "','" & strs!custom_field12 & "','" & strs!custom_field13 & "','" & strs!custom_field14 & "','" & strs!custom_field15 & "','" & strs!custom_field16 & "')"
'
'
'      On Error Resume Next
'
'      Set cmd.ActiveConnection = conn
'      Err.Clear
'      cmd.Execute
'
'      If Err.Number <> 0 Then
'        MsgBox Err.Description
'        RegisterMove = False
'        Exit Function
'      End If
'
'
'      ' перемещение на
'
'      Set cmd = New ADODB.Command
'      cmd.CommandText = "INSERT INTO HISTORY(" & _
'      "Stamp, Code , Item, Description, LOT_SN, EXP_DATE, UNIT_COST, QTY_ON_HAND, QTY, UOM, Status," & _
'      "LOCATION , REF_NUM, ORD_NUM, USER_NAME, SITE_ID, PALLET, Container," & _
'      "CUSTOM_FIELD1 , CUSTOM_FIELD2, CUSTOM_FIELD3, CUSTOM_FIELD4, CUSTOM_FIELD5, CUSTOM_FIELD6, CUSTOM_FIELD7, CUSTOM_FIELD8," & _
'      "CUSTOM_FIELD9, CUSTOM_FIELD10, CUSTOM_FIELD11, CUSTOM_FIELD12, CUSTOM_FIELD13, CUSTOM_FIELD14, CUSTOM_FIELD15, CUSTOM_FIELD16 " & _
'      ")VALUES( " & _
'      "getdate(),4,'" & itemrs!Code & "','" & itemrs!Description & "','" & strs!LOT_SN & "'," & MakeMSSQLDate(strs!exp_date) & "," & strs!Unit_COST & "," & 0 & "," & MyRound2(strs!QTY_ON_HAND) & ",'" & strs!UOM & "',0," & _
'      "'" & loctrs!Code & "','" & strs!ref_num & "','" & strs!ord_num & "','sa',1,'" & palrs!Code & "',''," & _
'      "'" & strs!custom_field1 & "','" & strs!custom_field2 & "','" & strs!custom_field3 & "','" & strs!custom_field4 & "','" & strs!custom_field5 & "','" & strs!custom_field6 & "','" & strs!custom_field7 & "','" & strs!custom_field8 & "'," & _
'      "'" & strs!custom_field9 & "','" & strs!custom_field10 & "','" & strs!custom_field11 & "','" & strs!custom_field12 & "','" & strs!custom_field13 & "','" & strs!custom_field14 & "','" & strs!custom_field15 & "','" & strs!custom_field16 & "')"
'
'      Set cmd.ActiveConnection = conn
'      Err.Clear
'      cmd.Execute
'
'      If Err.Number <> 0 Then
'        MsgBox Err.Description
'        RegisterMove = False
'        Exit Function
'      End If
'
'
'      ' обновление стока
'
'      Set cmd = New ADODB.Command
'      cmd.CommandText = "Update stock set location_id=" & loctrs!ID & " where id = " & strs!ID
'
'
'      Set cmd.ActiveConnection = conn
'      Err.Clear
'      cmd.Execute
'
'      If Err.Number <> 0 Then
'        MsgBox Err.Description
'        RegisterMove = False
'        Exit Function
'      End If
'
'      MD.ISMoved = Boolean_Da
'      MD.Save
'    End If
'  End If
'
'
'
'End Function

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
