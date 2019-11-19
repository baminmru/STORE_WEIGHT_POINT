Attribute VB_Name = "Module1"
Option Explicit
Public Manager As MTZManager.Main
Public Session As MTZSession.Session
Public UsersID As String
Public UserName As String
Public UserPassword As String
Public PrivateStoreID As String
Public SysStoreID As String
Public Site As String
Public LastChat As Date
Public NextReminder As Date
Public DeltaReminder As String
Public usr As MTZUsers.Application
Public MyUser As MTZUsers.Users
Public Action As String


Public repShowOL As ReportShow
Public repShowSRVOUT As ReportShow
Public repShowKL As ReportShow
Public repShowSRVIN As ReportShow
Public repShowINEPL As ReportShow
Public RptShowSRVALL As ReportShow
Public RptStickers As ReportShow
Public RptWrongLocation As ReportShow
Public repShowMoves As ReportShow
Public RptNedostacha As ReportShow
Public RptActVes As ReportShow
Public RptVimorozka As ReportShow
Public RptVimorozka2 As ReportShow
Public RptHran As ReportShow
Public RptStok103 As ReportShow
Public RptOtobrano As ReportShow
Public ITTDic As ITTD.Application

 







Public Function AddSQLRefIds(ByVal strTo As String, ByVal fldName As String, ByVal strFrom As String) As String
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








Sub Main()
Dim par() As String
Dim i As Long
Dim tst As Long
Dim UserPassword As String
Set Manager = New MTZManager.Main

tst = 0
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
          
          If UCase(Left(par(i), 4)) = "ACT:" Then
            Action = Right(par(i), Len(par(i)) - 4)
            tst = tst + 1
          End If

        Next
        If tst >= 3 Then
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
 
  
  Dim rs As ADODB.Recordset
  Set rs = Manager.ListInstances(Site, "MTZUsers")
  Set usr = Manager.GetInstanceObject(rs!InstanceID)
  Manager.LockInstanceObject usr.id
  
  
  Set rs = Nothing
  Set MyUser = usr.FindRowObject("Users", Session.GetSessionUserID())
  Set rs = Nothing
  
  If Action = "" Then
  
  Set MyRole = ChooseRole()
  If MyRole Is Nothing Then
      Session.Logout
     Set Manager = Nothing
     Exit Sub
  End If
  
  Manager.LockInstanceObject MyRole.id
  'frmSplash.lblWarning = "Подключение к Core IMS"
  End If
  
  Dim conn As ADODB.Connection
  Set conn = New ADODB.Connection
  conn.Provider = "SQLoledb"
  conn.ConnectionString = "Server=" & GetSetting("RBH", "ITTSETTINGS", "CORESRV", "") & ";DataBase=" & GetSetting("RBH", "ITTSETTINGS", "COREDB", "") & ";UID=" & GetSetting("RBH", "ITTSETTINGS", "COREUSR", "") & ";Pwd=" & GetSetting("RBH", "ITTSETTINGS", "COREPASS", "") & ";"
  conn.Open
  Manager.AddCustomObjects conn, "refref"
  Load frmMain
  
  If Action = "" Then
    frmMain.Show
  Else
    frmMain.DoAction Action
    Unload frmMain
    
  End If
  
End Sub


Public Sub PrintGrid(gr As Object)
  
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









Public Function FindPartia(ByVal LineName As String, Name As String) As Object

  Dim rs As ADODB.Recordset
  Set rs = Session.GetData("select * from v_AUTOITTD_PART where ITTD_PART_Name ='" & Name & "' and ITTD_PART_TheGood='" & LineName & "'")
  If Not rs.EOF Then
    Set FindPartia = MyUser.Application.FindRowObject("ITTD_PART", rs!id)
  End If

End Function

Public Function FindCountry(ByVal Name As String) As Object
  Dim rs As ADODB.Recordset
  
  Set rs = Session.GetData("select * from ITTD_COUNTRY where name ='" & Name & "'")
  If Not rs.EOF Then
    Set FindCountry = MyUser.Application.FindRowObject("ITTD_COUNTRY", rs!ITTD_COUNTRYID)
  End If

End Function


Public Function FindFactory(ByVal countryID As String, ByVal Name As String) As Object
  Dim rs As ADODB.Recordset
  
  Set rs = Session.GetData("select * from ITTD_FACTORY where name ='" & Name & "' and Country='" & countryID & "'")
  If Not rs.EOF Then
    Set FindFactory = MyUser.Application.FindRowObject("ITTD_FACTORY", rs!ITTD_FACTORYID)
  End If
End Function

Public Function FindKill(ByVal FactoryID As String, ByVal Name As String) As Object
Dim rs As ADODB.Recordset
  
  Set rs = Session.GetData("select * from ITTD_KILLPLACE where name ='" & Name & "' and Factory='" & FactoryID & "'")
  If Not rs.EOF Then
    Set FindKill = MyUser.Application.FindRowObject("ITTD_KILLPLACE", rs!ITTD_KILLPLACEID)
  End If


End Function


