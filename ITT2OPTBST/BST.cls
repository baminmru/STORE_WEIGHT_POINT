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
Attribute VB_HelpID = 10
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim theItem As ITT2OPT.Application
Dim Manager As MTZManager.Main
Dim session As MTZSession.session
'Dim jobs As MTZ2JOB.Application
Dim task As ITTOPT.Application
Dim conn As ADODB.Connection
Dim complexShifts As BoxShifts
Dim EB As EmptyBoxes
Dim Mixed As MixedBoxes

'Обработка  смены состояния документа
'Parameters:
'[IN][OUT]  Item , тип параметра: Object -  документ,
'[IN]   NewStatusid , тип параметра: String - новое состояние,
'[IN]   CurrentUser , тип параметра: Object -  пользователь,
'[IN]   Module , тип параметра: String  - название типа документа
'Returns:
' Boolean, семантика результата:
'   true  - можно менять состояние
'   false - нельзя менять состояние
'Example:
' dim variable as Boolean
' variable = me.Check(...параметры...)
Public Function Check(Item As Object, ByVal NewStatusid As String, ByVal CurrentUser As Object, ByVal Module As String) As Boolean
Attribute Check.VB_HelpID = 15
  Dim result As Boolean
  result = True
    
  Check = result
End Function


'Запуск оптимизации
' параметры:
'Item - документ Задание на оптимизацию
Public Function RunOptimization(Item As Object) As Boolean
  Dim result As Boolean
  Dim f As frmProgress, ID As String
  Dim iop As ITTOPT.Application
  Set iop = Item
  result = True
  
  On Error Resume Next
  Set Manager = iop.Manager
  Set session = iop.MTZSession
  
 
  Dim csstr As String
  Dim i As Long
  Dim rtype As ITTD.ITTD_RULE
  Dim opttype As ITTD_OPTTYPE
  Dim kid As Integer
  Dim where As String
  Dim cnt As Long
  
  Dim rs As ADODB.Recordset
  Dim rsname As ADODB.Recordset
  Dim CS_name As String
  Dim fi As Long
  Dim pfx As String
  
  pfx = ""
  

  csstr = ""
  kid = -1
  cnt = 0
  
  Dim def  As ITTOPT_DEF
  
  Set def = iop.ITTOPT_DEF.Item(1)
  Set rtype = iop.ITTOPT_DEF.Item(1).TheRule
  Set opttype = iop.ITTOPT_DEF.Item(1).opttype
  
  
  

  csstr = csstr & "item_code"

 If rtype.UseClient = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & ","
   End If
   csstr = csstr & "Partner"
  End If
  
  If rtype.TheCountry = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & ","
   End If
   csstr = csstr & "country"
  End If
  
  If rtype.TheFactory = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & ","
   End If
   csstr = csstr & "factory"
  End If
  
  If rtype.KillPlace = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & ","
   End If
   csstr = csstr & "kill_place"
  End If
  
 If rtype.UsePartia = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & ","
   End If
   csstr = csstr & "Partia"
  End If
  
 
  
  
   If rtype.UseVetsved = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & ","
   End If
   csstr = csstr & "vetsved"
  End If
  
  
  If rtype.UsePalType = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & ","
   End If
   csstr = csstr & "pallettype"
   pfx = "_pt"
  End If
   
  If rtype.UseBrak = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & ","
   End If
   csstr = csstr & "isbrak"
  End If
     
  
  If rtype.UseExpDate = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & ","
   End If
   csstr = csstr & "month(exp_date),year(exp_date)"
  End If
   
  
  Set conn = Manager.GetCustomObjects("refref")
  
  
  
   
      If rtype Is Nothing Or opttype Is Nothing Then
        MsgBox "Необходимо задать и сохранить тип оптимизации и правило формирования пратии", vbCritical, "Оптимизация"
        RunOptimization = False
        Exit Function
      End If
    
      iop.StatusID = "{451AF4A2-1E63-42B6-9494-A4F8301AEEBB}"
      
      
      
      If opttype.Code = 2 Then
        If def.TheKamera <= 0 Then
          kid = Val("0" & InputBox("Задайте номер камеры", "Номер камеры", 1))
        Else
          kid = def.TheKamera
        End If
        
        where = where & " and T='" & kid & "'"
      Else
        
         If def.TheKamera > 0 Then
             kid = def.TheKamera
             where = where & " and T='" & kid & "'"
         End If
        
        
      End If
      
     
        
       If def.made_country <> "" Then
        
        where = where & " and "
        
        where = where & "country like '" & Trim(Replace(def.made_country, ";", " ")) & "%'"
       End If
       
       If def.Factory <> "" Then
       
        where = where & " and "
     
        where = where & "factory like '" & Trim(Replace(def.Factory, ";", " ")) & "'"
       End If
       
       If def.KILL_NUMBER <> "" Then
     
        where = where & " and "
     
        where = where & "kill_place like '" & Trim(Replace(def.KILL_NUMBER, ";", " ")) & "'"
       End If
       
       If def.good <> "" Then

        where = where & " and "
        where = where & "item_code like '%" & def.good & "%'"
       End If
       
       If def.Partia <> "" Then
       
        where = where & " and "
     
        where = where & "Partia like '" & Trim(Replace(def.Partia, ";", " ")) & "'"
       End If


       If def.TheClient <> "" Then
       
        where = where & " and "
     
        where = where & "Partner like '%" & Trim(def.TheClient) & "%'"
       End If

        
      If def.IsBrak <> "" Then
       
        where = where & " and "
     
        where = where & "IsBrak like '%" & Trim(def.IsBrak) & "%'"
       End If
 
        
       If def.vetsved <> "" Then
       
        where = where & " and "
     
        where = where & "vetsved like '%" & Trim(def.vetsved) & "%'"
       End If
 
        
       If def.PalType <> "" Then
       
        where = where & " and "
     
        where = where & "pallettype = '" & Trim(def.PalType) & "'"
       End If
        
        
      
      ' получаем список пустых ячеек
      If opttype.Code = 2 Then
        Set rs = conn.Execute("SELECT * FROM V_BAMI_LOCATION WHERE loctype <>'B' and active='Y' and T='" & kid & "' and code not in (select loc_code from  v_bami_stock_all) order by QMAX")
      Else
       Set rs = conn.Execute("SELECT * FROM V_BAMI_LOCATION WHERE loctype <>'B' and active='Y' and code not in (select loc_code from  v_bami_stock_all) order by QMAX")
      End If
      
      Set EB = New EmptyBoxes
      If Not rs Is Nothing Then
        While Not rs.EOF
          EB.Add rs!T, False, rs!Code, rs!QMAX
          rs.MoveNext
        Wend
        rs.Close
      Else
        MsgBox "Проверьте правильность установок для соединения c базы Core IMS"
        RunOptimization = False
        Exit Function
      End If
      Set rs = Nothing
      
      Set Mixed = New MixedBoxes
      
      
      
      Set rs = conn.Execute("select count(distinct checksum(" & csstr & ")) cnt From v_bami_stock_all where 1=1 " & where)
     
      If rs Is Nothing Then
        MsgBox "Проверьте правильность установок для соединения c базы Core IMS"
        RunOptimization = False
        Exit Function
      End If
      
      Set f = New frmProgress
      f.pb.Min = 0
      f.pb.Value = 0
      f.pb.Max = rs!cnt
      f.Caption = "Оптимизация " & CLng(100 * f.pb.Value / f.pb.Max) & "%"
      f.Show
      
      
      
      With def
        If opttype.Code = 2 Then
          .TheKamera = kid
        End If
        .Save
      End With
      
      If complexShifts Is Nothing Then
        Set complexShifts = New BoxShifts
      End If
      
      Set rs = conn.Execute("select distinct checksum(" & csstr & ") CS From v_bami_stock_all where 1=1 " & where)
      While Not rs.EOF
        i = i + 1
        
        CS_name = ""
        Set rsname = conn.Execute("select top 1 " & csstr & " from v_bami_stock_all  where checksum(" & csstr & ") =" & rs!CS)
        If Not rsname Is Nothing Then
          If Not rsname.EOF Then
          
            For fi = 0 To rsname.Fields.Count - 1
              If CS_name <> "" Then CS_name = CS_name & "; "
              CS_name = CS_name & rsname.Fields.Item(fi).Value & ""
            Next
          End If
          rsname.Close
        End If
        Set rsname = Nothing
        
        f.pb.Value = i
        f.Caption = "Оптимизация " & CLng(100 * f.pb.Value / f.pb.Max) & "%"
        f.Label1 = "Обработка товара:" & CS_name
        DoEvents
        cnt = cnt + BuildOptimizingTask(kid, opttype, csstr, rs!CS, CS_name, where)
        rs.MoveNext
      Wend
      
      Unload f
      Set f = Nothing
      
      Set f = New frmProgress
      f.pb.Min = 0
      f.pb.Value = 0
      f.pb.Max = complexShifts.Count
      f.Label1 = "Сохранение"
      f.Caption = "Сохранение " & CLng(100 * f.pb.Value / f.pb.Max) & "%"
      f.Show
      
      'complexShifts.Base.Sort "FromCode"
      
      For i = 1 To complexShifts.Count
        f.pb.Value = i
        f.Caption = "Сохранение " & CLng(100 * f.pb.Value / f.pb.Max) & "%"
        DoEvents
        
        With iop.ITTOPT_MOVE.Add
          .sequence = iop.ITTOPT_MOVE.Count
          .FromLocation = complexShifts.Item(i).FromCode
          .ToLocation = complexShifts.Item(i).ToCode
          .GoodInfo = complexShifts.Item(i).GoodInfo
          .ThePalletPlanned = complexShifts.Item(i).pallet
          .ptype = complexShifts.Item(i).ptype
          .Stage = complexShifts.Item(i).Stage
          
trysave:
          Err.Clear
          
          .Save
          If Err.Number <> 0 Then
            If MsgBox("Ошибка сохранения строки задания" & vbCrLf & Err.Description & vbCrLf & "Повторить попытку записи ?", vbYesNo + vbCritical, "Сохранение задания на перемещение") = vbYes Then
             GoTo trysave
            End If
          End If
        End With
        
      Next
      
      Dim cnt2 As Long
      cnt2 = 0
      For i = 1 To EB.Count
      If EB.Item(i).Used Then
        cnt2 = cnt2 + 1
      End If
      Next
      
      MsgBox "Освобождается  ячеек:" & cnt & "." & vbCrLf & "Занято свободных ячеек под разбор " & cnt2 & ".", vbOKOnly, "Оптимизация склада"
      
      ' состояния для типа:ITTOPT Задание на перемещения
      ' "{0A7FC795-E787-4D17-9689-96EFFF8F0D9D}" 'Задание исполнено
      ' "{300483B2-1D94-4A33-8ADF-ABF32E72E57B}" 'Оформлено
      ' "{C861FA15-0DF6-42D4-BCE9-2B38C3E6C0CB}" 'Оформляется
        
 
      
      
      
trystate:
          Err.Clear
          
          iop.StatusID = "{300483B2-1D94-4A33-8ADF-ABF32E72E57B}"
          If Err.Number <> 0 Then
            If MsgBox("Ошибка сохранения состояния" & vbCrLf & Err.Description & vbCrLf & "Повторить попытку записи ?", vbYesNo + vbCritical, "Сохранение задания на перемещение") = vbYes Then
             GoTo trystate
            End If
          End If
      DoEvents
      
      rs.Close
      Set rs = Nothing
      result = True
      Unload f
      Set f = Nothing

    
    
  RunOptimization = result
End Function

' оптимизация одного товара
' параметры:
'
'KameraID As Integer - номер камеры
' opttype As ITTD_OPTTYPE - тип оптимизации
' csstr As String - сборка списка полей для вычисления контрольной суммы
' CS As Long - контрольная сумма
' CS_name As String - сборка названия товара
' where As String -  доп. условие
'Returns:
'  значение типа Long кол-во пустых ячеек
Private Function BuildOptimizingTask(ByVal KameraID As Integer, ByVal opttype As ITTD_OPTTYPE, ByVal csstr As String, ByVal CS As Long, ByVal CS_name As String, ByVal where As String) As Long
  Dim rs As ADODB.Recordset
  Dim rse As ADODB.Recordset
  Dim OPt As Optimizator
  Set OPt = New Optimizator
  Dim iop As ITTOPT.Application
  Dim i As Long
  Dim cnt As Long
  Dim ost As Long
  Dim lcSelect As String
  
  lcSelect = "select LC.loc_code from  ( select distinct loc_code ,checksum(" & csstr & ") X from v_bami_stock_all with (nolock)"
  lcSelect = lcSelect & vbCrLf & "  group by loc_code, checksum(" & csstr & ") ) LC group by  LC.loc_code having count(*)>1 "
  cnt = 0
  ost = 0
  
  Set OPt.st = New Boxes
  OPt.CS_name = CS_name
  
  Dim rspal As ADODB.Recordset
  Dim curbox As Box
  
  If opttype.Code = "0" Then
  
      
    Set rs = conn.Execute("select  sum(case when pallettype ='E' then 1 else 1.25 end) qcur,LOC_CODE,qmax,sklad,T,X,Y,Z from v_bami_stock_all  where 1=1 " & where & " and  (Loc_code in (" & lcSelect & ")) and checksum(" & csstr & ") = " & CS & " group by LOC_CODE,qmax,sklad,T,X,Y,Z having count(*) < qmax")
    
    While Not rs.EOF
          Set curbox = OPt.st.Add(rs!qcur, rs!QMAX, rs!T, rs!Z, rs!Y, rs!X, rs!loc_code, CS_name)
          curbox.HaveMixedPallet = True
          Set rse = conn.Execute("select  sum(case when pallettype ='E' then 1 else 1.25 end) qcur from v_bami_stock_all  where Loc_code ='" & rs!loc_code & "'")
          If Not rse Is Nothing Then
            Mixed.Add rse!qcur, rs!loc_code
            rse.Close
            Set rse = Nothing
          End If
          Set rspal = conn.Execute("select * from v_bami_stock_all  where  Loc_code ='" & rs!loc_code & "' and checksum(" & csstr & ") = " & CS)
          If Not rspal Is Nothing Then
            While Not rspal.EOF
              If UCase(rspal!pallettype) = "E" Then
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1, "E"
              Else
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1.25, "I"
              End If
              rspal.MoveNext
            Wend
            rspal.Close
            
          End If
          Set rspal = Nothing
          cnt = cnt + 1
          rs.MoveNext
    Wend
  
  
    Set rs = conn.Execute("select  sum(case when pallettype ='E' then 1 else 1.25 end) qcur,LOC_CODE,qmax,sklad,T,X,Y,Z from v_bami_stock_all  where 1=1 " & where & " and (Loc_code not in (" & lcSelect & ")) and checksum(" & csstr & ") = " & CS & " group by LOC_CODE,qmax,sklad,T,X,Y,Z having count(*) < qmax")
    
    While Not rs.EOF
          Set curbox = OPt.st.Add(rs!qcur, rs!QMAX, rs!T, rs!Z, rs!Y, rs!X, rs!loc_code, CS_name)
          Set rspal = conn.Execute("select * from v_bami_stock_all  where  Loc_code ='" & rs!loc_code & "' and checksum(" & csstr & ") = " & CS)
          If Not rspal Is Nothing Then
            While Not rspal.EOF
              If UCase(rspal!pallettype) = "E" Then
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1, "E"
              Else
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1.25, "I"
              End If
              rspal.MoveNext
            Wend
            rspal.Close
            
          End If
          Set rspal = Nothing
          cnt = cnt + 1
          rs.MoveNext
    Wend
    If cnt > 1 Then
      Set OPt.EB = EB
      Set OPt.Mixed = Mixed
      ost = ost + OPt.OptimizeKamera(-1)
    End If
  End If
  
  If opttype.Code = "1" Then
    
    Set rs = conn.Execute("select  sum(case when pallettype ='E' then 1 else 1.25 end) qcur,LOC_CODE,qmax,sklad,T,X,Y,Z from v_bami_stock_all  where 1=1 " & where & " and (Loc_code  in (" & lcSelect & ")) and checksum(" & csstr & ") = " & CS & " group by LOC_CODE,qmax,sklad,T,X,Y,Z having count(*) < qmax")
    While Not rs.EOF
          Set curbox = OPt.st.Add(rs!qcur, rs!QMAX, rs!T, rs!Z, rs!Y, rs!X, rs!loc_code, CS_name)
          curbox.HaveMixedPallet = True
          Set rse = conn.Execute("select  sum(case when pallettype ='E' then 1 else 1.25 end) qcur from v_bami_stock_all  where Loc_code ='" & rs!loc_code & "'")
          If Not rse Is Nothing Then
            Mixed.Add rse!qcur, rs!loc_code
            rse.Close
            Set rse = Nothing
          End If
          Set rspal = conn.Execute("select * from v_bami_stock_all  where  Loc_code ='" & rs!loc_code & "' and checksum(" & csstr & ") = " & CS)
          If Not rspal Is Nothing Then
            While Not rspal.EOF
              If UCase(rspal!pallettype) = "E" Then
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1, "E"
              Else
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1.25, "I"
              End If
              rspal.MoveNext
            Wend
            rspal.Close
            
          End If
          Set rspal = Nothing
          cnt = cnt + 1
          rs.MoveNext
    Wend
    
    
    Set rs = conn.Execute("select  sum(case when pallettype ='E' then 1 else 1.25 end) qcur,LOC_CODE,qmax,sklad,T,X,Y,Z from v_bami_stock_all  where 1=1 " & where & " and (Loc_code not in (" & lcSelect & ")) and checksum(" & csstr & ") = " & CS & " group by LOC_CODE,qmax,sklad,T,X,Y,Z having count(*) < qmax")
    While Not rs.EOF
          Set curbox = OPt.st.Add(rs!qcur, rs!QMAX, rs!T, rs!Z, rs!Y, rs!X, rs!loc_code, CS_name)
          Set rspal = conn.Execute("select * from v_bami_stock_all  where  Loc_code ='" & rs!loc_code & "' and checksum(" & csstr & ") = " & CS)
          If Not rspal Is Nothing Then
            While Not rspal.EOF
              If UCase(rspal!pallettype) = "E" Then
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1, "E"
              Else
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1.25, "I"
              End If
              rspal.MoveNext
            Wend
            rspal.Close
            
          End If
          Set rspal = Nothing
          cnt = cnt + 1
          rs.MoveNext
    Wend
    If cnt > 1 Then
     Set OPt.EB = EB
     Set OPt.Mixed = Mixed
     ost = ost + OPt.OptimizeAll
    End If
  End If
  
  
  If opttype.Code = "2" Then
    
    Set rs = conn.Execute("select  sum(case when pallettype ='E' then 1 else 1.25 end) qcur,LOC_CODE,qmax,sklad,T,X,Y,Z from v_bami_stock_all  where  1=1 " & where & " and  (Loc_code in (" & lcSelect & ")) and checksum(" & csstr & ") = " & CS & " group by LOC_CODE,qmax,sklad,T,X,Y,Z having count(*) < qmax")
    While Not rs.EOF
          Set curbox = OPt.st.Add(rs!qcur, rs!QMAX, rs!T, rs!Z, rs!Y, rs!X, rs!loc_code, CS_name)
          Set rspal = conn.Execute("select * from v_bami_stock_all  where  Loc_code ='" & rs!loc_code & "' and checksum(" & csstr & ") = " & CS)
          curbox.HaveMixedPallet = True
          Set rse = conn.Execute("select  sum(case when pallettype ='E' then 1 else 1.25 end) qcur from v_bami_stock_all  where Loc_code ='" & rs!loc_code & "'")
          If Not rse Is Nothing Then
            Mixed.Add rse!qcur, rs!loc_code
            rse.Close
            Set rse = Nothing
          End If
          If Not rspal Is Nothing Then
            While Not rspal.EOF
              If UCase(rspal!pallettype) = "E" Then
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1, "E"
              Else
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1.25, "I"
              End If
              rspal.MoveNext
            Wend
            rspal.Close
            
          End If
          Set rspal = Nothing
          cnt = cnt + 1
          rs.MoveNext
    Wend
    
    
    Set rs = conn.Execute("select  sum(case when pallettype ='E' then 1 else 1.25 end) qcur,LOC_CODE,qmax,sklad,T,X,Y,Z from v_bami_stock_all  where  1=1 " & where & " and (Loc_code not in (" & lcSelect & ")) and checksum(" & csstr & ") = " & CS & " group by LOC_CODE,qmax,sklad,T,X,Y,Z having count(*) < qmax")
    While Not rs.EOF
          Set curbox = OPt.st.Add(rs!qcur, rs!QMAX, rs!T, rs!Z, rs!Y, rs!X, rs!loc_code, CS_name)
          Set rspal = conn.Execute("select * from v_bami_stock_all  where  Loc_code ='" & rs!loc_code & "' and checksum(" & csstr & ") = " & CS)
          If Not rspal Is Nothing Then
            While Not rspal.EOF
              If UCase(rspal!pallettype) = "E" Then
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1, "E"
              Else
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1.25, "I"
              End If
              rspal.MoveNext
            Wend
            rspal.Close
            
          End If
          Set rspal = Nothing
          cnt = cnt + 1
          rs.MoveNext
    Wend
    If cnt > 1 Then
     Set OPt.EB = EB
     Set OPt.Mixed = Mixed
     ost = ost + OPt.OptimizeKamera(KameraID)
    End If
  End If
  
  
  
  If opttype.Code = "3" Then
     Set rs = conn.Execute("select  sum(case when pallettype ='E' then 1 else 1.25 end) qcur,LOC_CODE,qmax,sklad,T,X,Y,Z from v_bami_stock_all  where  1=1 " & where & " and  (Loc_code in (" & lcSelect & ")) and checksum(" & csstr & ") = " & CS & " group by LOC_CODE,qmax,sklad,T,X,Y,Z  having count(*) < qmax")
    While Not rs.EOF
          Set curbox = OPt.st.Add(rs!qcur, rs!QMAX, rs!T, rs!Z, rs!Y, rs!X, rs!loc_code, CS_name)
          curbox.HaveMixedPallet = True
           Set rse = conn.Execute("select  sum(case when pallettype ='E' then 1 else 1.25 end) qcur from v_bami_stock_all  where Loc_code ='" & rs!loc_code & "'")
          If Not rse Is Nothing Then
            Mixed.Add rse!qcur, rs!loc_code
            rse.Close
            Set rse = Nothing
          End If
          Set rspal = conn.Execute("select * from v_bami_stock_all  where  Loc_code ='" & rs!loc_code & "' and checksum(" & csstr & ") = " & CS)
          If Not rspal Is Nothing Then
            While Not rspal.EOF
              If UCase(rspal!pallettype) = "E" Then
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1, "E"
              Else
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1.25, "I"
              End If
              rspal.MoveNext
            Wend
            rspal.Close
            
          End If
          Set rspal = Nothing
          cnt = cnt + 1
          rs.MoveNext
    Wend
    
    
    Set rs = conn.Execute("select  sum(case when pallettype ='E' then 1 else 1.25 end) qcur,LOC_CODE,qmax,sklad,T,X,Y,Z from v_bami_stock_all  where  1=1 " & where & " and  (Loc_code not in (" & lcSelect & ")) and checksum(" & csstr & ") = " & CS & " group by LOC_CODE,qmax,sklad,T,X,Y,Z  having count(*) < qmax")
    While Not rs.EOF
          Set curbox = OPt.st.Add(rs!qcur, rs!QMAX, rs!T, rs!Z, rs!Y, rs!X, rs!loc_code, CS_name)
          Set rspal = conn.Execute("select * from v_bami_stock_all  where  Loc_code ='" & rs!loc_code & "' and checksum(" & csstr & ") = " & CS)
          If Not rspal Is Nothing Then
            While Not rspal.EOF
              If UCase(rspal!pallettype) = "E" Then
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1, "E"
              Else
                curbox.Pallets.Add rspal!pallet_code, CS_name, 1.25, "I"
              End If
              rspal.MoveNext
            Wend
            rspal.Close
            
          End If
          Set rspal = Nothing
          cnt = cnt + 1
          rs.MoveNext
    Wend
    If cnt > 1 Then
     Set OPt.EB = EB
     Set OPt.Mixed = Mixed
     ost = ost + OPt.OptimizeKamera(-1)
    End If
  End If
  
  ' create new optimization task
  If OPt.shifts.Count > 0 Then
   
    
    For i = 1 To OPt.shifts.Count
      complexShifts.Add OPt.shifts.Item(i).ToCode, OPt.shifts.Item(i).FromCode, OPt.shifts.Item(i).pallet, CS_name, OPt.shifts.Item(i).ptype, OPt.shifts.Item(i).Stage
    Next
  
  End If
  
   BuildOptimizingTask = ost
End Function



