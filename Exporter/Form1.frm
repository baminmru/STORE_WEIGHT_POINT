VERSION 5.00
Object = "{1801C003-859D-471D-BF31-D4428050324B}#2.1#0"; "MTZ_PANEL.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmWizPoddons 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Взвешивание поддонов"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   3120
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Следующий поддон"
      Default         =   -1  'True
      Height          =   735
      Left            =   5280
      TabIndex        =   14
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Закончить"
      Height          =   735
      Left            =   2880
      TabIndex        =   13
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Шаг 1 - выбор заказа"
      Height          =   2655
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7815
      Begin VB.TextBox txtQryCode 
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Код заказа"
         Top             =   720
         Width           =   6015
      End
      Begin VB.TextBox txtTheClient 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   6615
      End
      Begin MTZ_PANEL.DropButton cmdQryCode 
         Height          =   300
         Left            =   6240
         TabIndex        =   10
         Tag             =   "refopen.ico"
         ToolTipText     =   "Код заказа"
         Top             =   720
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.Label lblQryCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Код заказа:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3000
      End
      Begin VB.Label Label14 
         Caption         =   "Клиент"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Шаг 2 - Взвешивание поддона"
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.CommandButton cmd3ClearW 
         Caption         =   "x"
         Height          =   375
         Left            =   5400
         TabIndex        =   4
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton cmd3ClearNum 
         Caption         =   "x"
         Height          =   375
         Left            =   5400
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txt3Weight 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1680
         Width           =   5055
      End
      Begin VB.TextBox txt3Poddon 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "Вес поддона"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Номер поддона"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2655
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   840
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      Handshaking     =   2
   End
End
Attribute VB_Name = "frmWizPoddons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StepNo As Integer
Dim XMLQryCode As String
Dim XMLTheClient As String
Dim Item As ITTIN.Application
Dim conn As ADODB.Connection
Private curQRow As ITTIN.ITTIN_QLINE
Dim LinePal As ITTIN_PALET
Dim pal As ITTPL_DEF
Private StopWeighting As Boolean
Private wave As MTZMCI.WavePlayer
Private emu As Boolean
Private port As String
Private psetup As String
Private poddon As ITTPL_DEF

Private Sub cmd3ClearW_Click()
txt3Weight = 0
End Sub

Private Sub cmdCancel_Click()
StepNo = 3
ProcessStatus
End Sub

Private Sub cmdNext_Click()
If StepNo = 1 And txtQryCode.Text = "" Then
  MsgBox "Надо выбрать заказ"
Else
  If StepNo < 2 Then
    StepNo = StepNo + 1
  Else
    If Val("0" & txt3Weight) > 0 Then
      SavePoddon
    Else
      MsgBox "Дождитесь получения веса с электронных весов"
      Exit Sub
    End If
  End If

  ProcessStatus
End If
End Sub

Private Sub SavePoddon()
On Error Resume Next
  If txt3Poddon <> "" And txt3Weight <> "" Then
    Set poddon = FindPoddon(txt3Poddon)
    If Not poddon Is Nothing Then
    ' состояния для типа:ITTPL Палетта
' "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}" 'Взвешена
' "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" 'На складе с грузом
' "{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}" 'Отправлена с грузом
' "{E9BFB749-A606-4DEF-A429-07D636F108C6}" 'Пустая
' "{588C5203-1E59-408E-92A1-B3DFED8C19FA}" 'Списана

      Dim Obj As ITTIN_EPL
      Dim conn As ADODB.Connection
      Dim cmd As ADODB.Command
      Dim rs As ADODB.Recordset
      
      If poddon.Application.StatusID <> "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" Then
      
        
        Set rs = Session.GetData("select instance.name from ITTIN_EPL join instance on " & _
        " ITTIN_EPL.InstanceID = instance.InstanceID" & _
        " and instance.status in ('{EB3A7D03-EB3F-4541-AD93-D55C92BE02AC}','{49A919F7-94A6-49DE-9280-1EEAC973647B}')" & _
        " where ITTIN_EPL.TheNumber ='" & poddon.id & "'")
        
        If Not rs Is Nothing Then
         If rs.EOF Then
          
            Set Obj = Nothing
            For i = 1 To Item.ITTIN_EPL.Count
               If poddon.id = Item.ITTIN_EPL.Item(i).TheNumber.id Then
                Set Obj = Item.ITTIN_EPL.Item(i).TheNumber
               End If
            Next
            If Obj Is Nothing Then
              Set Obj = Item.ITTIN_EPL.Add
            End If
            With Obj
              Set .TheNumber = poddon
              .PalWeight = MyRound(txt3Weight)
              .save
            End With
            
            poddon.Application.StatusID = "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}"
            poddon.Weight = MyRound(txt3Weight)
            poddon.WDate = Date
            poddon.save
            
            Set conn = Manager.GetCustomObjects("refref")
            If conn.State <> adStateOpen Then
              conn.Open
            End If
            Set cmd = New ADODB.Command
            cmd.CommandText = "update pallet_weight set weight =" & MyRound2(txt3Weight) & " ,date_weight=getdate() where code =" & poddon.TheNumber
            If conn.State <> adStateOpen Then
              conn.Open
            End If
            Set cmd.ActiveConnection = conn
    '        cmd.Execute
            err.Clear
            Set cmd.ActiveConnection = conn
            cmd.Execute
            If err.Number <> 0 Then
              MsgBox err.Description
            End If
            
            Me.Caption = "Взвешено поддонов к заказу:" & Item.ITTIN_EPL.Count
          Else
            MsgBox "Поддон закреплен за другим заказом:<" & rs!Name & ">, который сейчас оформляется."
          End If
        End If
      Else
        MsgBox "Поддон находится в состоянии <" & poddon.Application.StatusName & "> и не может быть добавлен к заказу"
      End If
      DoEvents
    End If
  End If
End Sub


Private Sub cmdQryCode_Click()
  On Error Resume Next

  Dim pars As New NamedValues
  Dim res As NamedValues
  If (txtQryCode.Tag = "") Then
    ' call MsgBox("Нет данных для запроса")
  Else
    txtQryCode.Tag = Replace(txtQryCode.Tag, "%ID%", " 1=1 ")
    Call pars.Add("xml", txtQryCode.Tag)
  End If
  If Manager.GetCustomObjects("cliFilter").Name <> "" Then
    Call pars.Add("filter", " and " & (Manager.GetCustomObjects("cliFilter").Name))
  End If
  Set res = Manager.GetSQLDataDialog(pars)
  If (Not res Is Nothing) Then
    Dim resStr As String
    resStr = res.Item("RESULT").Value
    If (resStr = "OK") Then
      txtQryCode.Tag = res.Item("xml").Value
      If (txtQryCode.Text <> res.Item("brief").Value) Then
        txtQryCode.Text = res.Item("brief").Value
        'mIDQryCode = res.Item("ID").Value
        Call txtQryCode_Change
        MakeItem
      End If
    Else
      Dim errStr As String
      errStr = res.Item("ErrorDescription").Value
      If (errStr <> vbNullString) Then
       Call MsgBox("Ошибка исполнения: " & errStr, vbOKOnly + vbCritical)
     End If
    End If
  End If
End Sub

Private Sub MakeItem()
On Error Resume Next
'Найти заказ у в нашей базе
  Dim rs As ADODB.Recordset
  Dim id As String
  Dim qID As String
  qID = Manager.GetIDFromXMLField(txtQryCode.Tag)
  id = ""
  Set rs = Session.GetData("select instanceid from ITTIN_DEF where QryCode like '%<ID>" & qID & "</ID>%'")
  If Not rs Is Nothing Then
    If Not rs.EOF Then
      id = rs!InstanceID
    End If
  End If
  rs.Close
  
  'Если нет заказа, то сформировать новый
  If id = "" Then
    id = CreateGUID2
    Manager.NewInstance id, "ITTIN", txtQryCode
    Set Item = Manager.GetInstanceObject(id)
    
    If conn.State <> ADODB.adStateOpen Then
      conn.Open
    End If
    
    Set rs = conn.Execute("select * from receiving_order where id=" & Manager.GetIDFromXMLField(txtQryCode.Tag))
    If rs.EOF Then Exit Sub
    
    
    With Item.ITTIN_DEF.Add
      .ProcessDate = Date
      .QryCode = txtQryCode.Tag
      .TheClient = txtTheClient.Tag
      .Supplier = rs!street1
      .TTN = rs!ACCOUNT_NUMBER
      .TTNDate = Date
      .TranspNumber = rs!Comment1
      .Container = rs!TRACK_NUMBER1
      .Track_time_in = Now
      .track_time_out = DateAdd("h", 4, Now)
      .temp_in_track = -1
      .save
    End With
    
    
    Dim XMLQRY_NUM As String
    Dim XMLLineAtQuery As String
    Dim XMLgood_ID As String
    
    Set rs = conn.Execute("select A.*, B.DESCRIPTION  BRIEF, B.code ARTICUL from receiving_line A join item B on A.item_id =B.id where (a.PARENT_ID  is null or a.parent_id=0) and a.order_id='" & qID & "'")
    While Not rs.EOF
      Set curQRow = Item.ITTIN_QLINE.Add
      With curQRow
        
        .edizm = "" & rs!UOM
        .articul = "" & rs!articul
        
        '.made_country = "" & rs!prod_country
        '.KILL_NUMBER = "" & rs!KILL_NUMBER
        
        If Not IsNull(rs!made_date) Then .made_date = rs!made_date
        If Not IsNull(rs!exp_date) Then .exp_date = rs!exp_date
        
        
        
        XMLLineAtQuery = "<SQLData>"
        XMLLineAtQuery = XMLLineAtQuery & "<connectionstring>ref</connectionstring>"
        XMLLineAtQuery = XMLLineAtQuery & "<connectionprovider>ref</connectionprovider>"
        XMLLineAtQuery = XMLLineAtQuery & "<query>select A.ID [Код], A.ORDER_ID [Код Заказа], A.QTY_ORD [Количество], B.DESCRIPTION [Наименование]  from receiving_line A join item B on A.item_id =B.id </query>"
        XMLLineAtQuery = XMLLineAtQuery & "<IDFieldName>Код</IDFieldName>"
        XMLLineAtQuery = XMLLineAtQuery & "<BriefFields>Наименование</BriefFields>"
        XMLLineAtQuery = XMLLineAtQuery & "<Brief>" & rs!brief & "</Brief>"
        XMLLineAtQuery = XMLLineAtQuery & "<ID>" & rs!id & "</ID>"
        XMLLineAtQuery = XMLLineAtQuery & "</SQLData>"
        
        .LineAtQuery = XMLLineAtQuery
        
        
        
        
        XMLQRY_NUM = "<SQLData>"
        XMLQRY_NUM = XMLQRY_NUM & "<connectionstring>ref</connectionstring>"
        XMLQRY_NUM = XMLQRY_NUM & "<connectionprovider>ref</connectionprovider>"
        XMLQRY_NUM = XMLQRY_NUM & "<query>select  QTY_ORD from receiving_line where ID='%LineAtQueryID%'</query>"
        XMLQRY_NUM = XMLQRY_NUM & "<IDFieldName>QTY_ORD</IDFieldName>"
        XMLQRY_NUM = XMLQRY_NUM & "<BriefFields>QTY_ORD</BriefFields>"
        XMLQRY_NUM = XMLQRY_NUM & "<ID>" & rs!QTY_ORD & "</ID>"
        XMLQRY_NUM = XMLQRY_NUM & "<Brief>" & rs!QTY_ORD & "</Brief>"
        XMLQRY_NUM = XMLQRY_NUM & "<LineAtQueryID>" & rs!id & "</LineAtQueryID>"
        XMLQRY_NUM = XMLQRY_NUM & "</SQLData>"
              
        .QRY_NUM = XMLQRY_NUM
         
        XMLgood_ID = "<SQLData>"
        XMLgood_ID = XMLgood_ID & "<connectionstring>ref</connectionstring>"
        XMLgood_ID = XMLgood_ID & "<connectionprovider>ref</connectionprovider>"
        XMLgood_ID = XMLgood_ID & "<query>select  item_id from RECEIVING_LINE where ID='%LineAtQueryID%'</query>"
        XMLgood_ID = XMLgood_ID & "<IDFieldName>ITEM_ID</IDFieldName>"
        XMLgood_ID = XMLgood_ID & "<BriefFields>ITEM_ID</BriefFields>"
        XMLgood_ID = XMLgood_ID & "<Brief>" & rs!item_id & "</Brief>"
        XMLgood_ID = XMLgood_ID & "<ID>" & rs!item_id & "</ID>"
        XMLgood_ID = XMLgood_ID & "<LineAtQueryID>" & rs!id & "</LineAtQueryID>"
        XMLgood_ID = XMLgood_ID & "</SQLData>"
        
        .good_id = XMLgood_ID
        
        .save
      End With
      
      Call GetNumValue(curQRow, "sequence", "{E7F3EE01-4EC4-41D2-8657-BA22089DE0E5}", Now, "IN%P", "")
      rs.MoveNext
    Wend
    
    
    Set rs = Session.GetData("select * from ITTCS_DEF where clientcode like '%<ID>" & Manager.GetIDFromXMLField(txtTheClient.Tag) & "</ID>%'")
    Dim srvid As String
    Dim srvObj As ITTCS.Application
    Dim srv As ITTD_SRV
    srvid = rs!InstanceID
    Set srvObj = Manager.GetInstanceObject(srvid)
    Dim i As Long
    For i = 1 To srvObj.ITTCS_LIN.Count
       Set srv = srvObj.ITTCS_LIN.Item(i).srv
       If srv.ForReceiving = Boolean_Da Then
          If srvObj.ITTCS_LIN.Item(i).UseSrv = Boolean_Da Then
            With Item.ITTIN_SRV.Add
               Set .srv = srv
               .Quantity = 0
               .save
            End With
          End If
       End If
    Next
  Else
    Set Item = Manager.GetInstanceObject(id)
  End If
End Sub

Private Sub LoadHeader(Item As Object)
'  txtSupplier = Item.Supplier
'  txtTTN = Item.TTN
'  dtpTTNDate = Date
'  If Item.TTNDate <> 0 Then
'   dtpTTNDate = Item.TTNDate
'  Else
'   dtpTTNDate.Value = Null
'  End If
'  txtTranspNumber = Item.TranspNumber
'  txtContainer = Item.Container
'  txtStampNumber = Item.StampNumber
'  txtStampStatus = Item.StampStatus
'  dtpTrack_time_in = Now
'  If Item.Track_time_in <> 0 Then
'   dtpTrack_time_in = Item.Track_time_in
'  Else
'   dtpTrack_time_in.Value = Null
'  End If
'  dtptrack_time_out = Now
'  If Item.track_time_out <> 0 Then
'   dtptrack_time_out = Item.track_time_out
'  Else
'   dtptrack_time_out.Value = Null
'  End If
'  txttemp_in_track = Item.temp_in_track

End Sub

Private Sub ProcessStatus()
On Error Resume Next
  Frame1.Visible = False
  Frame2.Visible = False
  

  Select Case StepNo
  Case 0
    cmdNext.Caption = "Начать процесс"
    
  Case 1
  
    Before1
    Frame1.Visible = True
  

    
  Case 2
    cmdNext.Caption = "Следующий поддон"
    Before2
    Frame2.Visible = True
  
    
  Case 3
  If MsgBox("Напечатать акт о весе поддонов ?", vbYesNo) = vbYes Then
    On Error Resume Next
    Set repShowINEPL = Nothing
    Set repShowINEPL = New ReportShow
    repShowINEPL.ReportSource = "V_viewITTIN_ITTIN_EPL"
    repShowINEPL.ReportFilter = " instanceid='" & Item.id & "'"
    repShowINEPL.ReportPath = App.Path & "\in_epl.rpt"
    repShowINEPL.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
    repShowINEPL.Run True
    Set repShowINEPL = Nothing
   End If
   Unload Me
   
  End Select
  

End Sub


Private Sub Before1()
On Error Resume Next
    txtQryCode.Text = ""
    txtQryCode.Tag = XMLQryCode
    LoadBtnPictures cmdQryCode, cmdQryCode.Tag
    cmdQryCode.RemoveAllMenu
    txtTheClient.Text = ""
    txtTheClient.Tag = XMLTheClient
End Sub



Private Sub Form_Load()
On Error Resume Next

    StepNo = 0
    XMLQryCode = "<SQLData>"
    XMLQryCode = XMLQryCode & "<connectionstring>ref</connectionstring>"
    XMLQryCode = XMLQryCode & "<connectionprovider>ref</connectionprovider>"
    XMLQryCode = XMLQryCode & "<query>select A.ID [КОД] , convert(varchar(30),A.NUMBER) +'  от ' + convert(varchar(30),A.ORD_DATE,111)  [Название], PARTNER.Name [Клиент]  from RECEIVING_ORDER A left join PARTNER  on A.PARTNER_ID=PARTNER.ID where (a.STATUS = 1 or a.status =0) </query>"
    XMLQryCode = XMLQryCode & "<IDFieldName>КОД</IDFieldName>"
    XMLQryCode = XMLQryCode & "<BriefFields>Название</BriefFields>"
    XMLQryCode = XMLQryCode & "</SQLData>"
    
  
    XMLTheClient = "<SQLData>"
    XMLTheClient = XMLTheClient & "<connectionstring>ref</connectionstring>"
    XMLTheClient = XMLTheClient & "<connectionprovider>ref</connectionprovider>"
    XMLTheClient = XMLTheClient & "<query>select partner.ID, partner.Name from RECEIVING_ORDER join partner on RECEIVING_ORDER.partner_id=partner.id where RECEIVING_ORDER.ID='%QryCodeID%'</query>"
    XMLTheClient = XMLTheClient & "<IDFieldName>ID</IDFieldName>"
    XMLTheClient = XMLTheClient & "<BriefFields>Name</BriefFields>"
    XMLTheClient = XMLTheClient & "</SQLData>"
    
    
    
    
    ProcessStatus
    Set conn = Manager.GetCustomObjects("refref")
    If GetSetting("RBH", "ITTSETTINGS", "SOUND", "False") <> "False" Then
      Set wave = New MTZMCI.WavePlayer
      wave.OpenDevice
    End If
    
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
   If UnloadMode <> 1 Then
     Cancel = -1
   Else
   wave.StopPlaying
   Set wave = Nothing
   Timer1.Enabled = False
   If MSComm1.PortOpen Then
      MSComm1.PortOpen = False
    End If
    
   End If
      
End Sub

Private Sub Timer1_Timer()
  
  Dim w As Double
  On Error Resume Next
  If StepNo = 2 Then
    
    If txt3Poddon = "" Then
      txt3Poddon.SetFocus
    End If
    
    If txt3Weight = "0" Or Not IsNumeric(txt3Weight) Then
      w = GetWeight
      If w > 0 Then
        txt3Weight = Round(w + 0.001, 1)
        MyBeep "Poddon"
      End If
    End If
  End If
End Sub

Private Sub txtQryCode_Change()
On Error Resume Next
If (txtQryCode.Text = "") Then
  ' Убрать Brief и ID
  If (txtQryCode.Tag <> "") Then
    Dim XMLDoc As New DOMDocument
    Call XMLDoc.loadXML(txtQryCode.Tag)
    Dim Node As MSXML2.IXMLDOMNode
    For Each Node In XMLDoc.childNodes.Item(0).childNodes
     If (Node.baseName = "ID") Then
       Node.Text = ""
     End If
     If (Node.baseName = "Brief") Then
       Node.Text = ""
     End If
    Next
    txtQryCode.Tag = XMLDoc.XML
  End If
End If

cmdTheClient_Click

End Sub


Private Sub Before2()

  On Error Resume Next
  txt3Poddon = ""
  txt3Weight = "0"
  
  If GetSetting("RBH", "ITTSETTINGS", "SOUND", "False") <> "False" Then
    Set wave = New MTZMCI.WavePlayer
    wave.OpenDevice
  End If

  emu = Not (GetSetting("RBH", "ITTSETTINGS", "EMULATOR", "False") = "False")
  psetup = GetSetting("RBH", "ITTSETTINGS", "WSETUP", "4800,e,8,1")
  port = GetSetting("RBH", "ITTSETTINGS", "WPORT", 1)

  If Not emu Then
    If MSComm1.PortOpen Then
      MSComm1.PortOpen = False
    End If
      
    MSComm1.Handshaking = comNone
    MSComm1.DTREnable = False
    MSComm1.EOFEnable = False
      
    MSComm1.Settings = psetup
    MSComm1.CommPort = port
    MSComm1.PortOpen = True
  End If
  
  
End Sub


Private Sub MyBeep(ByVal BeepType As String)
On Error Resume Next
      If Not wave Is Nothing Then
        On Error Resume Next
        wave.OpenFile App.Path & "\" & BeepType & ".wav"
        wave.Play
      End If
End Sub

Private Sub cmdTheClient_Click()
On Error Resume Next
  On Error Resume Next
  Dim pars As New NamedValues
  Dim res As NamedValues
  If (txtTheClient.Tag = "") Then
    ' call MsgBox("Нет данных для запроса")
  Else
    Call pars.Add("permanent", "true")
    txtTheClient.Tag = AddSQLRefIds(txtTheClient.Tag, "QryCode", txtQryCode.Tag)
    txtTheClient.Tag = Replace(txtTheClient.Tag, "%ID%", " 1=1 ")
    Call pars.Add("xml", txtTheClient.Tag)
  End If
  Set res = Manager.GetSQLDataDialog(pars)
  If (Not res Is Nothing) Then
    Dim resStr As String
    resStr = res.Item("RESULT").Value
    If (resStr = "OK") Then
      txtTheClient.Tag = res.Item("xml").Value
      If (txtTheClient.Text <> res.Item("brief").Value) Then
        txtTheClient.Text = res.Item("brief").Value
'        mIDTheClient = res.Item("ID").Value
        'Call txtTheClient_Change
      End If
    Else
      Dim errStr As String
      errStr = res.Item("ErrorDescription").Value
      If (errStr <> vbNullString) Then
       Call MsgBox("Ошибка исполнения: " & errStr, vbOKOnly + vbCritical)
     End If
    End If
  End If
End Sub



Public Function GetWeight4() As Double
  On Error Resume Next
    Dim ws As String
    Dim ch As String
    Dim start As Single
    Dim ws1 As String
    Dim ws2 As String
    GetWeight4 = 0
    
    MSComm1.output = Chr(68)
    start = Timer   ' Set start time.
    Do While Timer < start + 0.2
    Loop
    
    If MSComm1.InBufferCount > 0 Then GoTo answer_s1
    start = Timer   ' Set start time
    Do While Timer < start + 0.5
       If MSComm1.InBufferCount > 0 Then GoTo answer_s1
    Loop
    
    GetWeight4 = 0  ' не дождались ответа
    Exit Function
    
answer_s1:
    
    ws = MSComm1.Input
    ' первый раз вес стабилен
    If Asc(Mid(ws, 1, 1)) >= 128 Then
    
      ''''''''''''''''''''''''''''''''''''
      'ЗАДЕРЖКА !!!
      '
      ' ждем чтобы исключить дребезг
      start = Timer   ' Set start time.
      Do While Timer < start + 0.3
      Loop
      
      ' спрашиваем еще раз
      MSComm1.output = Chr(68)
      
      
      start = Timer   ' Set start time.
      Do While Timer < start + 0.2
      Loop
      
      If MSComm1.InBufferCount > 0 Then GoTo answer_s2
      start = Timer   ' Set start time
      Do While Timer < start + 0.5
         If MSComm1.InBufferCount > 0 Then GoTo answer_s2
      Loop
      
    End If
    
    GetWeight4 = 0 ' нет второго ответа
    Exit Function
    
answer_s2:

    ws = MSComm1.Input
    
    ' второй раз вес тоже стабилен
    If Asc(Mid(ws, 1, 1)) >= 128 Then
      MSComm1.output = Chr(69)
      start = Timer   ' Set start time.
      Do While Timer < start + 0.2
      Loop
      If MSComm1.InBufferCount > 0 Then GoTo answer_w1
      start = Timer   ' Set start time
      Do While Timer < start + 0.5
       If MSComm1.InBufferCount > 0 Then GoTo answer_w1
      Loop
    End If
    
    GetWeight4 = 0 ' вес не стабилен, или нет ответа
    Exit Function
    
answer_w1:

    ' прочли показания веса
    ws1 = MSComm1.Input
    
    
    ''''''''''''''''''''''''''''''''''''
    'ЗАДЕРЖКА !!!
    '
    ' ждем чтобы исключить дребезг
    start = Timer   ' Set start time.
    Do While Timer < start + 0.3
    Loop
    
    ' спрашиваем вес еще раз
    MSComm1.output = Chr(69)
    start = Timer   ' Set start time.
    Do While Timer < start + 0.2
    Loop
    
    If MSComm1.InBufferCount > 0 Then GoTo answer_w2
    start = Timer   ' Set start time
    Do While Timer < start + 0.5
       If MSComm1.InBufferCount > 0 Then GoTo answer_w2
    Loop
    
    GetWeight4 = 0 '  нет ответа
    Exit Function
      
answer_w2:
    ws = MSComm1.Input
  
    If ws1 = ws Then
      GetWeight4 = (Asc(Mid(ws, 2, 1)) * 256 + Asc(Mid(ws, 1, 1))) / 10
    Else
      GetWeight4 = 0 ' вес не стабилен, отличаются показания
    End If
  
End Function

Private Function GetWeight() As Double
On Error Resume Next
  If emu Then
    If StepNo = 6 Then
      GetWeight = Rnd(Second(Now)) * 1000 + MyRound("0" & txt3Weight)
    Else
      GetWeight = Rnd(Second(Now)) * 40
    End If
  Else
    GetWeight = GetWeight4
  End If
End Function
