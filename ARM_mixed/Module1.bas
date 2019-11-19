Attribute VB_Name = "Module1"
Attribute VB_HelpID = 1305
Option Explicit
'Запускающий модуль

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

' объекты для отображения отчетных форм
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

' справочник
Public ITTDic As ITTD.Application
Attribute ITTDic.VB_VarHelpID = 1340

' добавить параметры в блок SQL Refrence
'Parameters:
'[IN]   strTo , тип параметра: String - куда добавить,
'[IN]   fldName , тип параметра: String - из какого поля,
'[IN]   strFrom , тип параметра: String  - откуда добавить
'Returns:
'  значение типа String новое значение блока SQL Reference
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
' variable = me.AddSQLRefIds(...параметры...)
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

'Главная процедуура АРМ
'Parameters:
' параметров нет
'Returns:
' нет
Sub Main()
Attribute Main.VB_HelpID = 1350
Dim par() As String
Dim i As Long
Dim tst As Long
Dim UserPassword As String
Set Manager = New MTZManager.Main

tst = 0
'  разбор командной строки
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
  
'  отображение формы логина
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
      MsgBox "Не определен сайт с таким именем", vbCritical, "Ошибка"
      GoTo again
    End If
    
    
'    открытие сесии
    If Not Session.Login(f.txtUserName, f.txtPassword) Then
      Set Session = Nothing
      MsgBox "Неверные данные регистрации", vbCritical, "Ошибка"
      GoTo again
    End If
    UserName = f.txtUserName
    UserPassword = f.txtPassword
    Unload f
    Set f = Nothing
 
 End If
 
'  получение текущего пользователя
  Dim rs As ADODB.Recordset
  Set rs = Manager.ListInstances(Site, "MTZUsers")
  Set usr = Manager.GetInstanceObject(rs!InstanceID)
  Manager.LockInstanceObject usr.id
  
  
'  получение документа - общие настройки
  Dim fn As ITTFN.Application
  Set rs = Manager.ListInstances(Site, "ITTFN")
  Set fn = Manager.GetInstanceObject(rs!InstanceID)
   Manager.LockInstanceObject fn.id
  
'  сохранение данных префикса документов в коллекции именованных объектов
  If fn.ITTFN_INFO.Count = 1 Then
    Manager.AddCustomObjects fn.ITTFN_INFO.Item(1), "PFX"
  End If
  
  Set rs = Nothing
  Set MyUser = usr.FindRowObject("Users", Session.GetSessionUserID())
  Set rs = Nothing
  
'  выбор роли
  Set MyRole = ChooseRole()
  If MyRole Is Nothing Then
      Session.Logout
     Set Manager = Nothing
     Exit Sub
  End If
  
  Manager.LockInstanceObject MyRole.id
  
'  отображение сплэш окна
  frmSplash.Show
  frmSplash.lblWarning = "Загрузка умолчаний"
  DoEvents
  
   
  
  Dim orgid As String
  
    
'  загрузка лицензий
  frmSplash.lblWarning = "Загрузка лицензий"
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
   
   
'  регистрация документов для работы в режиме MDI
  frmSplash.lblWarning = "Подключение документов"
  DoEvents
  
  RegisterMDIGUI
  
'  открытие соединения с CORE IMS
  frmSplash.lblWarning = "Подключение к Core IMS"
  DoEvents
  Dim conn As ADODB.Connection
  Set conn = GetCoreConn()
  
  
'  загрузка и показ главной формы
  frmSplash.lblWarning = "Инициализация меню"
  DoEvents
  Load frmMain
  
  
  
  
  Unload frmSplash
  
  frmMain.Show
  
End Sub

'получение соединения с CORE при необходимости повторное подключение
'параметры
'  Check - boolean  - выдать проверочный запрос
'результат
'  Объект -Connection
Public Function GetCoreConn(Optional ByVal Check As Boolean = True) As ADODB.Connection
  
  Dim conn As ADODB.Connection
  
  ' получаем сохраненный объект - соединение
  Set conn = Manager.GetCustomObjects("refref")
  
  ' проверяем состояние
  If conn Is Nothing Then
'  создаем новое
    Set conn = New ADODB.Connection
    conn.Provider = "SQLoledb"
    
'    открываем соединение
    conn.ConnectionString = "Server=" & GetSetting("RBH", "ITTSETTINGS", "CORESRV", "") & ";DataBase=" & GetSetting("RBH", "ITTSETTINGS", "COREDB", "") & ";UID=" & GetSetting("RBH", "ITTSETTINGS", "COREUSR", "") & ";Pwd=" & GetSetting("RBH", "ITTSETTINGS", "COREPASS", "") & ";"
    conn.open
    If conn.State = adStateOpen Then
        Manager.AddCustomObjects conn, "refref"
    Else
        MsgBox "Невозможно соединиться с CORE IMS"
    End If
  Else
    If conn.State = adStateClosed Then
        conn.Provider = "SQLoledb"
        conn.ConnectionString = "Server=" & GetSetting("RBH", "ITTSETTINGS", "CORESRV", "") & ";DataBase=" & GetSetting("RBH", "ITTSETTINGS", "COREDB", "") & ";UID=" & GetSetting("RBH", "ITTSETTINGS", "COREUSR", "") & ";Pwd=" & GetSetting("RBH", "ITTSETTINGS", "COREPASS", "") & ";"
        '    открываем соединение
        conn.open
        Manager.RemoveCustomObjects "refref"
        Manager.AddCustomObjects conn, "refref"
    End If
  End If
  
  If Not conn Is Nothing Then
'   делаем проверочный запрос
   If Check Then
   
    Dim rs As ADODB.Recordset
    On Error Resume Next
    err.Clear
    Call rs.open("SELECT 'OK' SRV_TEST", conn)
    If err.Number <> 0 Then
'            пытаемся переоткрыть соединение
            conn.Close
            conn.Provider = "SQLoledb"
            conn.ConnectionString = "Server=" & GetSetting("RBH", "ITTSETTINGS", "CORESRV", "") & ";DataBase=" & GetSetting("RBH", "ITTSETTINGS", "COREDB", "") & ";UID=" & GetSetting("RBH", "ITTSETTINGS", "COREUSR", "") & ";Pwd=" & GetSetting("RBH", "ITTSETTINGS", "COREPASS", "") & ";"
            conn.open
            Manager.RemoveCustomObjects "refref"
            Manager.AddCustomObjects conn, "refref"
    Else
        If rs!SRV_TEST <> "OK" Then
        '            пытаемся переоткрыть соединение
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




'Печать таблицы
'Parameters:
'[IN][OUT]  gr , тип параметра: Object  - таблица
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
'  call me.PrintGrid(...параметры...)
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

  ' считаем сколько страниц надо по ширине
  curw = 0
  HorPages = 1
  ReDim ColPerPage(HorPages)
  ColPerPage(HorPages) = 0
  For i = 0 To gr.Cols - 1
    If gr.ColWidth(i) > 0 Then curw = curw + gr.ColWidth(i) * dx

    ' ширина превысила размер страницы
    If curw > pw Then
      HorPages = HorPages + 1
      ReDim Preserve ColPerPage(HorPages)
      ColPerPage(HorPages) = IIf(i - 1 < 1, 1, i - 1)
      curw = gr.ColWidth(i) * dx
    End If

    ' если колонка очень широкая то запихаем ее в отдельную страницу
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

  ' цикл по вертикальным блокам
  While FirstRow < gr.Rows

    ' Горизонтальный блок страниц
    For i = 1 To HorPages
      curw = 0

      ' колонки для каждой из страниц
      For j = ColPerPage(i) To ColPerPage(i + 1) - 1

        ' только видимые колонки
        If gr.ColWidth(j) > 0 Then
          CellTop = 0
          CurRow = FirstRow

          ' ограничение по высоте листа
          While CellTop <= ph

              ' не проходим по высоте листа
              If CellTop + gr.RowHeight(CurRow) * dy > ph Then
                If gr.RowHeight(CurRow) * dy > ph Then
                  ' если высота колонки очень велика то меняем ее на меньшую
                  gr.RowHeight(CurRow) = ph / dy
                  GoTo nxtcol
                Else
                  GoTo nxtcol
                End If
              End If

              ' пересчитываем прямоугольник для отрисовки текста
              r.Left = curw / Printer.TwipsPerPixelX + 2
              r.Right = IIf((curw + gr.ColWidth(j) * dx) > pw, pw, curw + gr.ColWidth(j) * dx) _
                / Printer.TwipsPerPixelX - 2
              r.Top = CellTop / Printer.TwipsPerPixelY + 2
              r.Bottom = (CellTop + gr.RowHeight(CurRow) * dy) / Printer.TwipsPerPixelY - 2

              ' Первую строку отделяем жирной линией
              If CurRow = 0 Then
                Printer.Line (curw, (CellTop + gr.RowHeight(CurRow) * dy) - 20)- _
                  (IIf((curw + gr.ColWidth(j) * dx) > pw, pw, curw + gr.ColWidth(j) * dx), _
                  (CellTop + gr.RowHeight(CurRow) * dy)), , BF
              End If


              ' выводим рамочку
              Printer.Line (curw, CellTop)- _
                (IIf((curw + gr.ColWidth(j) * dx) > pw, pw, curw + gr.ColWidth(j) * dx), _
                (CellTop + gr.RowHeight(CurRow) * dy)), , B


              ' выводим текст в прямоугольную область (с переносом слов)
              DrawText Printer.hdc, gr.TextMatrix(CurRow, j), Len(gr.TextMatrix(CurRow, j)), r, &H10 + &H100

              ' изменяем позицию для следующей строки
              CellTop = CellTop + gr.RowHeight(CurRow) * dy

              ' готовимся к следующей сторке
              CurRow = CurRow + 1
              If CurRow >= gr.Rows Then GoTo nxtcol

          Wend
nxtcol:
          ' учитываем ширину и переходим к следующей колонке
          curw = curw + gr.ColWidth(j) * dx
        End If
      Next ' цикл по колонкам


      ' печатаем номер страницы
      Printer.Line (0, ph - 20)-(Printer.ScaleWidth, ph), , B
      Printer.CurrentX = Printer.ScaleWidth / 3
      Printer.CurrentY = ph + 100
      pcnt = pcnt + 1
      Printer.Print "Страница №" & pcnt
      ' не отбиваем страницу после последнего листа
      If CurRow < gr.Rows Or i < HorPages Then Printer.NewPage
    Next
    ' готовимся к новому блоку горизонтальных страниц
    FirstRow = CurRow
  Wend
  Printer.EndDoc
End Sub





' Регистрация документов для вывода в режиме MDI Child
'параметров нет
'результатов нет
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

' найти партию по имени
'Parameters:
'[IN]   LineName , тип параметра: String - товар,
'[IN][OUT]   Name , тип параметра: String  - партия
'Returns:
'  объект - запись о партии
'  ,или Nothing
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
' Set variable = me.FindPartia(...параметры...)
Public Function FindPartia(ByVal LineName As String, Name As String) As Object
Attribute FindPartia.VB_HelpID = 1335

  Dim rs As ADODB.Recordset
  Set rs = Session.GetData("select * from v_AUTOITTD_PART where ITTD_PART_Name ='" & Name & "' and ITTD_PART_TheGood='" & LineName & "'")
  If Not rs.EOF Then
    Set FindPartia = MyUser.Application.FindRowObject("ITTD_PART", rs!id)
  End If

End Function

'найти страну по имени
'Parameters:
'[IN]   Name , тип параметра: String  - название страны
'Returns:
'  объект - запись о стране
'  ,или Nothing
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
' Set variable = me.FindCountry(...параметры...)
Public Function FindCountry(ByVal Name As String) As Object
Attribute FindCountry.VB_HelpID = 1320
  Dim rs As ADODB.Recordset
  
  Set rs = Session.GetData("select * from ITTD_COUNTRY where name ='" & Name & "'")
  If Not rs.EOF Then
    Set FindCountry = MyUser.Application.FindRowObject("ITTD_COUNTRY", rs!ITTD_COUNTRYID)
  End If

End Function

'найти завод по имени
'Parameters:
'[IN]   countryID , тип параметра: String,
'[IN]   Name , тип параметра: String  - имя завода
'Returns:
'  объект  - запись о заводе
'  ,или Nothing
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
' Set variable = me.FindFactory(...параметры...)
Public Function FindFactory(ByVal countryID As String, ByVal Name As String) As Object
Attribute FindFactory.VB_HelpID = 1325
  Dim rs As ADODB.Recordset
  
  Set rs = Session.GetData("select * from ITTD_FACTORY where name ='" & Name & "' and Country='" & countryID & "'")
  If Not rs.EOF Then
    Set FindFactory = MyUser.Application.FindRowObject("ITTD_FACTORY", rs!ITTD_FACTORYID)
  End If
End Function

'Найти бойню по имени
'Parameters:
'[IN]   FactoryID , тип параметра: String - ID завода,
'[IN]   Name , тип параметра: String  - название бойни
'Returns:
'  объект любого класса Visual Basic
'  ,или Nothing
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
' Set variable = me.FindKill(...параметры...)
Public Function FindKill(ByVal FactoryID As String, ByVal Name As String) As Object
Attribute FindKill.VB_HelpID = 1330
Dim rs As ADODB.Recordset
  
  Set rs = Session.GetData("select * from ITTD_KILLPLACE where name ='" & Name & "' and Factory='" & FactoryID & "'")
  If Not rs.EOF Then
    Set FindKill = MyUser.Application.FindRowObject("ITTD_KILLPLACE", rs!ITTD_KILLPLACEID)
  End If


End Function


'получить магическое слово
'параметр
'  d -дата
'резуьтат
'  магическое слово
Public Function GetMagicWord(ByVal d As Date) As String
Dim magicWord(0 To 43) As String
Dim idx As Long

idx = Day(d) + Month(d)

magicWord(0) = "бар"
magicWord(1) = "бок"
magicWord(2) = "бор"
magicWord(3) = "бук"
magicWord(4) = "вес"
magicWord(5) = "гад"
magicWord(6) = "газ"
magicWord(7) = "где"
magicWord(8) = "гид"
magicWord(9) = "год"
magicWord(10) = "два"
magicWord(11) = "док"
magicWord(12) = "дуб"
magicWord(13) = "жук"
magicWord(14) = "зуб"
magicWord(15) = "зуд"
magicWord(16) = "код"
magicWord(17) = "кол"
magicWord(18) = "кот"
magicWord(19) = "кто"
magicWord(20) = "куб"
magicWord(21) = "лес"
magicWord(22) = "лоб"
magicWord(23) = "луг"
magicWord(24) = "лук"
magicWord(25) = "май"
magicWord(26) = "мир"
magicWord(27) = "нос"
magicWord(28) = "пир"
magicWord(29) = "пул"
magicWord(30) = "раз"
magicWord(31) = "рим"
magicWord(32) = "рог"
magicWord(33) = "рот"
magicWord(34) = "сми"
magicWord(35) = "сок"
magicWord(36) = "сор"
magicWord(37) = "сук"
magicWord(38) = "тон"
magicWord(39) = "тор"
magicWord(40) = "тот"
magicWord(41) = "три"
magicWord(42) = "ухо"
magicWord(43) = "что"

GetMagicWord = magicWord(idx)

End Function


'отобразить окно сообщения с возможностью ввода магического слова
'параметр
'  message - string  - сообщение об ошибке
'результат
'  выбор пользователя на форме
Public Function MagicMessageBox(message As String) As Boolean

Dim frm As frmMagicBOX
Set frm = New frmMagicBOX
frm.txtMessage = message
frm.Show vbModal
MagicMessageBox = frm.OK

Unload frm
Set frm = Nothing

End Function
