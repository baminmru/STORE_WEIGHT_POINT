Attribute VB_Name = "Module1"
Attribute VB_HelpID = 1305
Option Explicit
'����������� ������

Public Manager As MTZManager.Main
Attribute Manager.VB_VarHelpID = 1355
Public Session As MTZSession.Session
Attribute Session.VB_VarHelpID = 1460
Public UsersID As String
Attribute UsersID.VB_VarHelpID = 1485
Public UserName As String
Attribute UserName.VB_VarHelpID = 1475
Public UserPassword As String
Attribute UserPassword.VB_VarHelpID = 1480
Public PrivateStoreID As String
Attribute PrivateStoreID.VB_VarHelpID = 1375
Public SysStoreID As String
Attribute SysStoreID.VB_VarHelpID = 1470
Public Site As String
Attribute Site.VB_VarHelpID = 1465
Public LastChat As Date
Attribute LastChat.VB_VarHelpID = 1345
Public NextReminder As Date
Attribute NextReminder.VB_VarHelpID = 1365
Public DeltaReminder As String
Attribute DeltaReminder.VB_VarHelpID = 1315
Public usr As MTZUsers.Application
Attribute usr.VB_VarHelpID = 1490
Public MyUser As MTZUsers.Users
Attribute MyUser.VB_VarHelpID = 1360
Public LastMagic As String

' ������� ��� ����������� �������� ����
Public repShowOL As ReportShow
Attribute repShowOL.VB_VarHelpID = 1395
Public repShowSRVOUT As ReportShow
Attribute repShowSRVOUT.VB_VarHelpID = 1405
Public repShowKL As ReportShow
Attribute repShowKL.VB_VarHelpID = 1385
Public repShowSRVIN As ReportShow
Attribute repShowSRVIN.VB_VarHelpID = 1400
Public repShowINEPL As ReportShow
Attribute repShowINEPL.VB_VarHelpID = 1380
Public RptShowSRVALL As ReportShow
Attribute RptShowSRVALL.VB_VarHelpID = 1430
Public RptStickers As ReportShow
Attribute RptStickers.VB_VarHelpID = 1435
Public RptWrongLocation As ReportShow
Attribute RptWrongLocation.VB_VarHelpID = 1455
Public repShowMoves As ReportShow
Attribute repShowMoves.VB_VarHelpID = 1390
Public RptNedostacha As ReportShow
Attribute RptNedostacha.VB_VarHelpID = 1420
Public RptActVes As ReportShow
Attribute RptActVes.VB_VarHelpID = 1410
Public RptVimorozka As ReportShow
Attribute RptVimorozka.VB_VarHelpID = 1445
Public RptVimorozka2 As ReportShow
Attribute RptVimorozka2.VB_VarHelpID = 1450
Public RptHran As ReportShow
Attribute RptHran.VB_VarHelpID = 1415
Public RptStok103 As ReportShow
Attribute RptStok103.VB_VarHelpID = 1440
Public RptOtobrano As ReportShow
Attribute RptOtobrano.VB_VarHelpID = 1425

' ����������
Public ITTDic As ITTD.Application
Attribute ITTDic.VB_VarHelpID = 1340

' �������� ��������� � ���� SQL Refrence
'Parameters:
'[IN]   strTo , ��� ���������: String - ���� ��������,
'[IN]   fldName , ��� ���������: String - �� ������ ����,
'[IN]   strFrom , ��� ���������: String  - ������ ��������
'Returns:
'  �������� ���� String ����� �������� ����� SQL Reference
'See Also:
'  DeltaReminder
'  FindCountry
'  FindFactory
'  FindKill
'  FindPartia
'  ITTDic
'  LastChat
'  Main
'  Manager
'  MyUser
'  NextReminder
'  PrintGrid
'  PrivateStoreID
'  repShowINEPL
'  repShowKL
'  repShowMoves
'  repShowOL
'  repShowSRVIN
'  repShowSRVOUT
'  RptActVes
'  RptHran
'  RptNedostacha
'  RptOtobrano
'  RptShowSRVALL
'  RptStickers
'  RptStok103
'  RptVimorozka
'  RptVimorozka2
'  RptWrongLocation
'  Session
'  Site
'  SysStoreID
'  UserName
'  UserPassword
'  UsersID
'  usr
'Example:
' dim variable as String
' variable = me.AddSQLRefIds(...���������...)
Public Function AddSQLRefIds(ByVal strTo As String, ByVal fldName As String, ByVal strFrom As String) As String
Attribute AddSQLRefIds.VB_HelpID = 1310
  Dim XMLDocFrom As New DOMDocument
  Dim XMLDocTo As New DOMDocument
  AddSQLRefIds = strTo
  On Error GoTo err
  Call XMLDocTo.loadXML(strTo)
  Call XMLDocFrom.loadXML(strFrom)
  Dim Node As MSXML2.IXMLDOMNode
  Dim id As String
  For Each Node In XMLDocFrom.childNodes.Item(0).childNodes
    If (Node.baseName = "ID") Then
      id = Node.Text
      Dim NodeTO As MSXML2.IXMLDOMNode
      Dim bAdded As Boolean
      bAdded = False
      For Each NodeTO In XMLDocTo.childNodes.Item(0).childNodes
       If (NodeTO.baseName = fldName & "ID") Then
         NodeTO.Text = id
         bAdded = True
         Exit For
       End If
      Next
      If (Not bAdded) Then
       Dim newNode As MSXML2.IXMLDOMNode
       Set newNode = XMLDocTo.createNode(MSXML2.NODE_ELEMENT, fldName & "ID", XMLDocTo.namespaceURI)
        newNode.Text = id
       Call XMLDocTo.childNodes.Item(0).appendChild(newNode)
      End If
      AddSQLRefIds = XMLDocTo.XML
      Exit For
    End If
  Next
err:
End Function

'������� ���������� ���
'Parameters:
' ���������� ���
'Returns:
' ���
Sub Main()
Attribute Main.VB_HelpID = 1350
Dim par() As String
Dim i As Long
Dim tst As Long
Dim UserPassword As String
Set Manager = New MTZManager.Main

tst = 0
'  ������ ��������� ������
  If Command$ <> "" Then
        par() = Split(Command, " ")
        For i = LBound(par) To UBound(par)
          If UCase(Left(par(i), 4)) = "USR:" Then
            UserName = Right(par(i), Len(par(i)) - 4)
            tst = tst + 1
          End If
          
          If UCase(Left(par(i), 4)) = "PWD:" Then
            UserPassword = Right(par(i), Len(par(i)) - 4)
            tst = tst + 1
          End If
          
          If UCase(Left(par(i), 4)) = "APP:" Then
            Site = Right(par(i), Len(par(i)) - 4)
            tst = tst + 1
          End If

        Next
        If tst = 3 Then
          Set Session = Manager.GetSession(Site)
          If Session Is Nothing Then
            GoTo useForm
          End If
          
          If Not Session.Login(UserName, UserPassword) Then
            Set Session = Nothing
            GoTo useForm
          End If
        Else
         GoTo useForm
        End If
  Else
  
'  ����������� ����� ������
useForm:
    Dim f As frmLogin
    Set f = New frmLogin

again:
    Set Session = Nothing
    Set Manager = Nothing
    Set Manager = New MTZManager.Main
    
    f.Show vbModal
    If Not f.OK Then
      Unload f
      Set f = Nothing
      Set Manager = Nothing
      Exit Sub
    End If
    Site = f.txtSite
    
    Set Session = Manager.GetSession(Site)
    If Session Is Nothing Then
      MsgBox "�� ��������� ���� � ����� ������", vbCritical, "������"
      GoTo again
    End If
    
    
'    �������� �����
    If Not Session.Login(f.txtUserName, f.txtPassword) Then
      Set Session = Nothing
      MsgBox "�������� ������ �����������", vbCritical, "������"
      GoTo again
    End If
    UserName = f.txtUserName
    UserPassword = f.txtPassword
    Unload f
    Set f = Nothing
 
 End If
 
'  ��������� �������� ������������
  Dim rs As ADODB.Recordset
  Set rs = Manager.ListInstances(Site, "MTZUsers")
  Set usr = Manager.GetInstanceObject(rs!InstanceID)
  Manager.LockInstanceObject usr.id
  
  
'  ��������� ��������� - ����� ���������
  Dim fn As ITTFN.Application
  Set rs = Manager.ListInstances(Site, "ITTFN")
  Set fn = Manager.GetInstanceObject(rs!InstanceID)
   Manager.LockInstanceObject fn.id
  
'  ���������� ������ �������� ���������� � ��������� ����������� ��������
  If fn.ITTFN_INFO.Count = 1 Then
    Manager.AddCustomObjects fn.ITTFN_INFO.Item(1), "PFX"
  End If
  
  Set rs = Nothing
  Set MyUser = usr.FindRowObject("Users", Session.GetSessionUserID())
  Set rs = Nothing
  
'  ����� ����
  Set MyRole = ChooseRole()
  If MyRole Is Nothing Then
      Session.Logout
     Set Manager = Nothing
     Exit Sub
  End If
  
  Manager.LockInstanceObject MyRole.id
  
'  ����������� ����� ����
  frmSplash.Show
  frmSplash.lblWarning = "�������� ���������"
  DoEvents
  
   
  
  Dim orgid As String
  
    
'  �������� ��������
  frmSplash.lblWarning = "�������� ��������"
  DoEvents
  On Error Resume Next
   Dim intFile As Integer
   intFile = FreeFile
   Open App.Path & "\Licenses.txt" For Input As #intFile
   Dim strKey As String, strprogid As String
   ' On the client machine, read the license key from the file.
   
   
   While Not EOF(intFile)
    strprogid = ""
    strKey = ""
    Input #intFile, strprogid, strKey
    If strprogid <> "" Then
      Licenses.Add strprogid, strKey
    Else
      GoTo closefile
    End If
   Wend

closefile:
   Close #intFile
   
   
'  ����������� ���������� ��� ������ � ������ MDI
  frmSplash.lblWarning = "����������� ����������"
  DoEvents
  
  RegisterMDIGUI
  
'  �������� ���������� � CORE IMS
  frmSplash.lblWarning = "����������� � Core IMS"
  DoEvents
  Dim conn As ADODB.Connection
  Set conn = GetCoreConn()
  
  
'  �������� � ����� ������� �����
  frmSplash.lblWarning = "������������� ����"
  DoEvents
  Load frmMain
  
  
  
  
  Unload frmSplash
  
  frmMain.Show
  
End Sub

'��������� ���������� � CORE ��� ������������� ��������� �����������
'���������
'  Check - boolean  - ������ ����������� ������
'���������
'  ������ -Connection
Public Function GetCoreConn(Optional ByVal Check As Boolean = True) As ADODB.Connection
  
  Dim conn As ADODB.Connection
  
  ' �������� ����������� ������ - ����������
  Set conn = Manager.GetCustomObjects("refref")
  
  ' ��������� ���������
  If conn Is Nothing Then
'  ������� �����
    Set conn = New ADODB.Connection
    conn.Provider = "SQLoledb"
    
'    ��������� ����������
    conn.ConnectionString = "Server=" & GetSetting("RBH", "ITTSETTINGS", "CORESRV", "") & ";DataBase=" & GetSetting("RBH", "ITTSETTINGS", "COREDB", "") & ";UID=" & GetSetting("RBH", "ITTSETTINGS", "COREUSR", "") & ";Pwd=" & GetSetting("RBH", "ITTSETTINGS", "COREPASS", "") & ";"
    conn.open
    If conn.State = adStateOpen Then
        Manager.AddCustomObjects conn, "refref"
    Else
        MsgBox "���������� ����������� � CORE IMS"
    End If
  Else
    If conn.State = adStateClosed Then
        conn.Provider = "SQLoledb"
        conn.ConnectionString = "Server=" & GetSetting("RBH", "ITTSETTINGS", "CORESRV", "") & ";DataBase=" & GetSetting("RBH", "ITTSETTINGS", "COREDB", "") & ";UID=" & GetSetting("RBH", "ITTSETTINGS", "COREUSR", "") & ";Pwd=" & GetSetting("RBH", "ITTSETTINGS", "COREPASS", "") & ";"
        '    ��������� ����������
        conn.open
        Manager.RemoveCustomObjects "refref"
        Manager.AddCustomObjects conn, "refref"
    End If
  End If
  
  If Not conn Is Nothing Then
'   ������ ����������� ������
   If Check Then
   
    Dim rs As ADODB.Recordset
    On Error Resume Next
    err.Clear
    Call rs.open("SELECT 'OK' SRV_TEST", conn)
    If err.Number <> 0 Then
'            �������� ����������� ����������
            conn.Close
            conn.Provider = "SQLoledb"
            conn.ConnectionString = "Server=" & GetSetting("RBH", "ITTSETTINGS", "CORESRV", "") & ";DataBase=" & GetSetting("RBH", "ITTSETTINGS", "COREDB", "") & ";UID=" & GetSetting("RBH", "ITTSETTINGS", "COREUSR", "") & ";Pwd=" & GetSetting("RBH", "ITTSETTINGS", "COREPASS", "") & ";"
            conn.open
            Manager.RemoveCustomObjects "refref"
            Manager.AddCustomObjects conn, "refref"
    Else
        If rs!SRV_TEST <> "OK" Then
        '            �������� ����������� ����������
            conn.Close
            conn.Provider = "SQLoledb"
            conn.ConnectionString = "Server=" & GetSetting("RBH", "ITTSETTINGS", "CORESRV", "") & ";DataBase=" & GetSetting("RBH", "ITTSETTINGS", "COREDB", "") & ";UID=" & GetSetting("RBH", "ITTSETTINGS", "COREUSR", "") & ";Pwd=" & GetSetting("RBH", "ITTSETTINGS", "COREPASS", "") & ";"
            conn.open
            Manager.RemoveCustomObjects "refref"
            Manager.AddCustomObjects conn, "refref"
        Else
            'conn - OK!
        End If
    End If
    
   End If
    
  End If
  
  
  Set GetCoreConn = conn
End Function




'������ �������
'Parameters:
'[IN][OUT]  gr , ��� ���������: Object  - �������
'See Also:
'  AddSQLRefIds
'  DeltaReminder
'  FindCountry
'  FindFactory
'  FindKill
'  FindPartia
'  ITTDic
'  LastChat
'  Main
'  Manager
'  MyUser
'  NextReminder
'  PrivateStoreID
'  repShowINEPL
'  repShowKL
'  repShowMoves
'  repShowOL
'  repShowSRVIN
'  repShowSRVOUT
'  RptActVes
'  RptHran
'  RptNedostacha
'  RptOtobrano
'  RptShowSRVALL
'  RptStickers
'  RptStok103
'  RptVimorozka
'  RptVimorozka2
'  RptWrongLocation
'  Session
'  Site
'  SysStoreID
'  UserName
'  UserPassword
'  UsersID
'  usr
'Example:
'  call me.PrintGrid(...���������...)
Public Sub PrintGrid(gr As Object)
Attribute PrintGrid.VB_HelpID = 1370
  
  Dim r As RECT
  Dim ph As Long, pw As Long
  Dim i As Long, j As Long
  Dim ColPerPage() As Long, HorPages As Long, curw As Long
  Dim CurRow As Long, CurCol As Long, FirstRow As Long, CellTop As Long
  Dim dx As Double, dy As Double, pcnt As Long

  ph = Printer.ScaleHeight - 1000
  pw = Printer.ScaleWidth - 200
  dx = 1.1
  dy = 1.1
  pcnt = 0

  ' ������� ������� ������� ���� �� ������
  curw = 0
  HorPages = 1
  ReDim ColPerPage(HorPages)
  ColPerPage(HorPages) = 0
  For i = 0 To gr.Cols - 1
    If gr.ColWidth(i) > 0 Then curw = curw + gr.ColWidth(i) * dx

    ' ������ ��������� ������ ��������
    If curw > pw Then
      HorPages = HorPages + 1
      ReDim Preserve ColPerPage(HorPages)
      ColPerPage(HorPages) = IIf(i - 1 < 1, 1, i - 1)
      curw = gr.ColWidth(i) * dx
    End If

    ' ���� ������� ����� ������� �� �������� �� � ��������� ��������
    If i > 0 And curw > pw Then
      HorPages = HorPages + 1
      ReDim Preserve ColPerPage(HorPages)
      ColPerPage(HorPages) = i
      curw = 0
    End If
  Next
  ReDim Preserve ColPerPage(HorPages + 1)
  ColPerPage(HorPages + 1) = gr.Cols

  CurCol = 0
  CurRow = 0
  FirstRow = 0
  Printer.Font.Name = gr.Font.Name
  Printer.Font.Bold = gr.Font.Bold
  Printer.Font.Charset = gr.Font.Charset
  Printer.Font.Italic = gr.Font.Italic
  Printer.Font.Strikethrough = gr.Font.Strikethrough
  Printer.Font.Underline = gr.Font.Underline
  Printer.Font.Weight = gr.Font.Weight
  Printer.Font.Size = gr.Font.Size

  ' ���� �� ������������ ������
  While FirstRow < gr.Rows

    ' �������������� ���� �������
    For i = 1 To HorPages
      curw = 0

      ' ������� ��� ������ �� �������
      For j = ColPerPage(i) To ColPerPage(i + 1) - 1

        ' ������ ������� �������
        If gr.ColWidth(j) > 0 Then
          CellTop = 0
          CurRow = FirstRow

          ' ����������� �� ������ �����
          While CellTop <= ph

              ' �� �������� �� ������ �����
              If CellTop + gr.RowHeight(CurRow) * dy > ph Then
                If gr.RowHeight(CurRow) * dy > ph Then
                  ' ���� ������ ������� ����� ������ �� ������ �� �� �������
                  gr.RowHeight(CurRow) = ph / dy
                  GoTo nxtcol
                Else
                  GoTo nxtcol
                End If
              End If

              ' ������������� ������������� ��� ��������� ������
              r.Left = curw / Printer.TwipsPerPixelX + 2
              r.Right = IIf((curw + gr.ColWidth(j) * dx) > pw, pw, curw + gr.ColWidth(j) * dx) _
                / Printer.TwipsPerPixelX - 2
              r.Top = CellTop / Printer.TwipsPerPixelY + 2
              r.Bottom = (CellTop + gr.RowHeight(CurRow) * dy) / Printer.TwipsPerPixelY - 2

              ' ������ ������ �������� ������ ������
              If CurRow = 0 Then
                Printer.Line (curw, (CellTop + gr.RowHeight(CurRow) * dy) - 20)- _
                  (IIf((curw + gr.ColWidth(j) * dx) > pw, pw, curw + gr.ColWidth(j) * dx), _
                  (CellTop + gr.RowHeight(CurRow) * dy)), , BF
              End If


              ' ������� �������
              Printer.Line (curw, CellTop)- _
                (IIf((curw + gr.ColWidth(j) * dx) > pw, pw, curw + gr.ColWidth(j) * dx), _
                (CellTop + gr.RowHeight(CurRow) * dy)), , B


              ' ������� ����� � ������������� ������� (� ��������� ����)
              DrawText Printer.hdc, gr.TextMatrix(CurRow, j), Len(gr.TextMatrix(CurRow, j)), r, &H10 + &H100

              ' �������� ������� ��� ��������� ������
              CellTop = CellTop + gr.RowHeight(CurRow) * dy

              ' ��������� � ��������� ������
              CurRow = CurRow + 1
              If CurRow >= gr.Rows Then GoTo nxtcol

          Wend
nxtcol:
          ' ��������� ������ � ��������� � ��������� �������
          curw = curw + gr.ColWidth(j) * dx
        End If
      Next ' ���� �� ��������


      ' �������� ����� ��������
      Printer.Line (0, ph - 20)-(Printer.ScaleWidth, ph), , B
      Printer.CurrentX = Printer.ScaleWidth / 3
      Printer.CurrentY = ph + 100
      pcnt = pcnt + 1
      Printer.Print "�������� �" & pcnt
      ' �� �������� �������� ����� ���������� �����
      If CurRow < gr.Rows Or i < HorPages Then Printer.NewPage
    Next
    ' ��������� � ������ ����� �������������� �������
    FirstRow = CurRow
  Wend
  Printer.EndDoc
End Sub





' ����������� ���������� ��� ������ � ������ MDI Child
'���������� ���
'����������� ���
Private Sub RegisterMDIGUI()
 Dim g As gui
Set g = New gui
g.INIT "ITTIN"
Manager.RegisterGUI g, "ITTIN"
Set g = New gui
g.INIT "ITTOUT"
Manager.RegisterGUI g, "ITTOUT"
Set g = New gui
g.INIT "ITTFN"
Manager.RegisterGUI g, "ITTFN"
Set g = New gui
g.INIT "ITTCS"
Manager.RegisterGUI g, "ITTCS"
Set g = New gui
g.INIT "ITTPL"
Manager.RegisterGUI g, "ITTPL"
Set g = New gui
g.INIT "ITTD"
Manager.RegisterGUI g, "ITTD"
Set g = New gui
g.INIT "ITTOP"
Manager.RegisterGUI g, "ITTOP"
Set g = New gui
g.INIT "ITTOPT"
Manager.RegisterGUI g, "ITTOPT"

Set g = New gui
g.INIT "ITT2OPT"
Manager.RegisterGUI g, "ITT2OPT"

Set g = New gui
g.INIT "ITTNO"
Manager.RegisterGUI g, "ITTNO"

End Sub

' ����� ������ �� �����
'Parameters:
'[IN]   LineName , ��� ���������: String - �����,
'[IN][OUT]   Name , ��� ���������: String  - ������
'Returns:
'  ������ - ������ � ������
'  ,��� Nothing
'See Also:
'  AddSQLRefIds
'  DeltaReminder
'  FindCountry
'  FindFactory
'  FindKill
'  ITTDic
'  LastChat
'  Main
'  Manager
'  MyUser
'  NextReminder
'  PrintGrid
'  PrivateStoreID
'  repShowINEPL
'  repShowKL
'  repShowMoves
'  repShowOL
'  repShowSRVIN
'  repShowSRVOUT
'  RptActVes
'  RptHran
'  RptNedostacha
'  RptOtobrano
'  RptShowSRVALL
'  RptStickers
'  RptStok103
'  RptVimorozka
'  RptVimorozka2
'  RptWrongLocation
'  Session
'  Site
'  SysStoreID
'  UserName
'  UserPassword
'  UsersID
'  usr
'Example:
' dim variable as Object
' Set variable = me.FindPartia(...���������...)
Public Function FindPartia(ByVal LineName As String, Name As String) As Object
Attribute FindPartia.VB_HelpID = 1335

  Dim rs As ADODB.Recordset
  Set rs = Session.GetData("select * from v_AUTOITTD_PART where ITTD_PART_Name ='" & Name & "' and ITTD_PART_TheGood='" & LineName & "'")
  If Not rs.EOF Then
    Set FindPartia = MyUser.Application.FindRowObject("ITTD_PART", rs!id)
  End If

End Function

'����� ������ �� �����
'Parameters:
'[IN]   Name , ��� ���������: String  - �������� ������
'Returns:
'  ������ - ������ � ������
'  ,��� Nothing
'See Also:
'  AddSQLRefIds
'  DeltaReminder
'  FindFactory
'  FindKill
'  FindPartia
'  ITTDic
'  LastChat
'  Main
'  Manager
'  MyUser
'  NextReminder
'  PrintGrid
'  PrivateStoreID
'  repShowINEPL
'  repShowKL
'  repShowMoves
'  repShowOL
'  repShowSRVIN
'  repShowSRVOUT
'  RptActVes
'  RptHran
'  RptNedostacha
'  RptOtobrano
'  RptShowSRVALL
'  RptStickers
'  RptStok103
'  RptVimorozka
'  RptVimorozka2
'  RptWrongLocation
'  Session
'  Site
'  SysStoreID
'  UserName
'  UserPassword
'  UsersID
'  usr
'Example:
' dim variable as Object
' Set variable = me.FindCountry(...���������...)
Public Function FindCountry(ByVal Name As String) As Object
Attribute FindCountry.VB_HelpID = 1320
  Dim rs As ADODB.Recordset
  
  Set rs = Session.GetData("select * from ITTD_COUNTRY where name ='" & Name & "'")
  If Not rs.EOF Then
    Set FindCountry = MyUser.Application.FindRowObject("ITTD_COUNTRY", rs!ITTD_COUNTRYID)
  End If

End Function

'����� ����� �� �����
'Parameters:
'[IN]   countryID , ��� ���������: String,
'[IN]   Name , ��� ���������: String  - ��� ������
'Returns:
'  ������  - ������ � ������
'  ,��� Nothing
'See Also:
'  AddSQLRefIds
'  DeltaReminder
'  FindCountry
'  FindKill
'  FindPartia
'  ITTDic
'  LastChat
'  Main
'  Manager
'  MyUser
'  NextReminder
'  PrintGrid
'  PrivateStoreID
'  repShowINEPL
'  repShowKL
'  repShowMoves
'  repShowOL
'  repShowSRVIN
'  repShowSRVOUT
'  RptActVes
'  RptHran
'  RptNedostacha
'  RptOtobrano
'  RptShowSRVALL
'  RptStickers
'  RptStok103
'  RptVimorozka
'  RptVimorozka2
'  RptWrongLocation
'  Session
'  Site
'  SysStoreID
'  UserName
'  UserPassword
'  UsersID
'  usr
'Example:
' dim variable as Object
' Set variable = me.FindFactory(...���������...)
Public Function FindFactory(ByVal countryID As String, ByVal Name As String) As Object
Attribute FindFactory.VB_HelpID = 1325
  Dim rs As ADODB.Recordset
  
  Set rs = Session.GetData("select * from ITTD_FACTORY where name ='" & Name & "' and Country='" & countryID & "'")
  If Not rs.EOF Then
    Set FindFactory = MyUser.Application.FindRowObject("ITTD_FACTORY", rs!ITTD_FACTORYID)
  End If
End Function

'����� ����� �� �����
'Parameters:
'[IN]   FactoryID , ��� ���������: String - ID ������,
'[IN]   Name , ��� ���������: String  - �������� �����
'Returns:
'  ������ ������ ������ Visual Basic
'  ,��� Nothing
'See Also:
'  AddSQLRefIds
'  DeltaReminder
'  FindCountry
'  FindFactory
'  FindPartia
'  ITTDic
'  LastChat
'  Main
'  Manager
'  MyUser
'  NextReminder
'  PrintGrid
'  PrivateStoreID
'  repShowINEPL
'  repShowKL
'  repShowMoves
'  repShowOL
'  repShowSRVIN
'  repShowSRVOUT
'  RptActVes
'  RptHran
'  RptNedostacha
'  RptOtobrano
'  RptShowSRVALL
'  RptStickers
'  RptStok103
'  RptVimorozka
'  RptVimorozka2
'  RptWrongLocation
'  Session
'  Site
'  SysStoreID
'  UserName
'  UserPassword
'  UsersID
'  usr
'Example:
' dim variable as Object
' Set variable = me.FindKill(...���������...)
Public Function FindKill(ByVal FactoryID As String, ByVal Name As String) As Object
Attribute FindKill.VB_HelpID = 1330
Dim rs As ADODB.Recordset
  
  Set rs = Session.GetData("select * from ITTD_KILLPLACE where name ='" & Name & "' and Factory='" & FactoryID & "'")
  If Not rs.EOF Then
    Set FindKill = MyUser.Application.FindRowObject("ITTD_KILLPLACE", rs!ITTD_KILLPLACEID)
  End If


End Function


'�������� ���������� �����
'��������
'  d -����
'��������
'  ���������� �����
Public Function GetMagicWord(ByVal d As Date) As String
Dim magicWord(0 To 43) As String
Dim idx As Long

idx = Day(d) + Month(d)

magicWord(0) = "���"
magicWord(1) = "���"
magicWord(2) = "���"
magicWord(3) = "���"
magicWord(4) = "���"
magicWord(5) = "���"
magicWord(6) = "���"
magicWord(7) = "���"
magicWord(8) = "���"
magicWord(9) = "���"
magicWord(10) = "���"
magicWord(11) = "���"
magicWord(12) = "���"
magicWord(13) = "���"
magicWord(14) = "���"
magicWord(15) = "���"
magicWord(16) = "���"
magicWord(17) = "���"
magicWord(18) = "���"
magicWord(19) = "���"
magicWord(20) = "���"
magicWord(21) = "���"
magicWord(22) = "���"
magicWord(23) = "���"
magicWord(24) = "���"
magicWord(25) = "���"
magicWord(26) = "���"
magicWord(27) = "���"
magicWord(28) = "���"
magicWord(29) = "���"
magicWord(30) = "���"
magicWord(31) = "���"
magicWord(32) = "���"
magicWord(33) = "���"
magicWord(34) = "���"
magicWord(35) = "���"
magicWord(36) = "���"
magicWord(37) = "���"
magicWord(38) = "���"
magicWord(39) = "���"
magicWord(40) = "���"
magicWord(41) = "���"
magicWord(42) = "���"
magicWord(43) = "���"

GetMagicWord = magicWord(idx)

End Function


'���������� ���� ��������� � ������������ ����� ����������� �����
'��������
'  message - string  - ��������� �� ������
'���������
'  ����� ������������ �� �����
Public Function MagicMessageBox(message As String) As Boolean

Dim frm As frmMagicBOX
Set frm = New frmMagicBOX
frm.txtMessage = message
frm.Show vbModal
MagicMessageBox = frm.OK

Unload frm
Set frm = Nothing

End Function
