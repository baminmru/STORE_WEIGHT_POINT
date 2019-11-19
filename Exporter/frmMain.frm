VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Главное окно"
   ClientHeight    =   7170
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9555
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "out"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0624
            Key             =   "toQry"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":093E
            Key             =   "in"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C58
            Key             =   "auto"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F72
            Key             =   "div"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":128C
            Key             =   "add"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   9495
      TabIndex        =   0
      Top             =   6195
      Visible         =   0   'False
      Width           =   9555
      Begin RichTextLib.RichTextBox rtf 
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   393217
         TextRTF         =   $"frmMain.frx":1BDA
      End
   End
   Begin VB.Timer MenuTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2355
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   1665
      Top             =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1080
      Top             =   840
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   240
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Файл"
      Begin VB.Menu mnuSetings 
         Caption         =   "Настройка"
         Begin VB.Menu mnuCoreSetup 
            Caption         =   "Настройка соединения"
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Выход"
      End
   End
   Begin VB.Menu mnuDictionary 
      Caption         =   "Справочники"
      Begin VB.Menu mnuITTFN 
         Caption         =   "Настройки системы"
      End
      Begin VB.Menu mnuITTOP 
         Caption         =   "Операторы и кладовщики"
      End
      Begin VB.Menu mnuITTD 
         Caption         =   "Справочник"
      End
      Begin VB.Menu mnuITTNO 
         Caption         =   "Настройки оптимизатора"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Отчеты"
      Begin VB.Menu mnuVimorozka 
         Caption         =   "Отчет по выморозке"
      End
      Begin VB.Menu mnuRptHran 
         Caption         =   "Отчет по объему услуг хранения"
      End
      Begin VB.Menu mnuOtbor 
         Caption         =   "Отчет по отбору"
      End
      Begin VB.Menu mnuRptOtobrano 
         Caption         =   "Объемы отбора товара"
      End
      Begin VB.Menu mnuRpt103 
         Caption         =   "Заблокировано на выморозку"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'%%%JOURNALPLACEHOLDER%%%
'sample: Dim WithEvents fDog As frmJournalShow2

'%%%REPORTPLACEHOLDER%%%
'sample: Public RptResult As ReportShow

'Dim ObjectToReport As Object

Dim frmFind As Form
Dim frmFindFT As Form

Dim inTimer1 As Boolean
Dim inTimer2 As Boolean
Dim OnLoad As Boolean
Dim DelayedCommand As String


Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)



Public Sub DoAction(ByVal Action As String)
Dim conn As ADODB.Connection

Select Case UCase(Action)
Case "RPT1":
    
    Set RptOtobrano = New ReportShow
    RptOtobrano.ReportPath = App.Path & "\otobrano.rpt"
    RptOtobrano.ReportSource = "v_bami_stockblocked"
    Set conn = Manager.GetCustomObjects("refref")
    
    'Call RptOtobrano.Run(False, conn)
    
    RptOtobrano.ExportPDF App.Path & "\otobrano.pdf", conn
    MailThisFile "Отбор выморозки", "Отчет по отбору выморозки на  " & Now & ".", App.Path & "\" & "otobrano.pdf"
    Set RptOtobrano = Nothing
    
Case "RPT2":

    Set RptStok103 = New ReportShow
    RptStok103.ReportPath = App.Path & "\stok103.rpt"
    RptStok103.ReportSource = "v_bami_stock103"
    Set conn = Manager.GetCustomObjects("refref")
    
    'Call RptStok103.Run(False, conn)
    RptStok103.ExportPDF App.Path & "\" & "Stock103.pdf", conn
     
    MailThisFile "Заблокированные поддоны", "Отчет по заблокированным поддонам на  " & Now & ".", App.Path & "\" & "Stock103.pdf"
    Set RptStok103 = Nothing
    
End Select

End Sub


Private Sub MDIForm_Load()
  On_Load


 Dim o As ITTOP.Application
 Dim rs  As ADODB.Recordset
 Dim id As String
 Dim cliFilter As String
 Dim camFilter As String
 Dim i As Long, j As Long
  Set rs = Manager.ListInstances("", "ITTOP")
  If Not rs.EOF Then
    id = rs!InstanceID
  Else
    id = CreateGUID2
    Manager.NewInstance id, "ITTOP", "Операторы и кладовщики"
End If
Set o = Manager.GetInstanceObject(id)


For i = 1 To o.ITTOP_KLNK.Count
  If o.ITTOP_KLNK.Item(i).TheUser Is MyUser Then
    For j = 1 To o.ITTOP_KLNK.Item(i).ITTOP_KCLI.Count
      If cliFilter <> "" Then
      cliFilter = cliFilter & ","
      End If
      cliFilter = cliFilter & "'" & GetBRIEFFromXMLField(o.ITTOP_KLNK.Item(i).ITTOP_KCLI.Item(j).TheClient) & "'"
    Next
    For j = 1 To o.ITTOP_KLNK.Item(i).ITTOP_KCAM.Count
     If camFilter <> "" Then
      camFilter = camFilter & " or "
      End If
      camFilter = camFilter & " location.code like '" & o.ITTOP_KLNK.Item(i).ITTOP_KCAM.Item(j).TheKamera.CameraMask & "' "
    Next
  
  End If
Next


For i = 1 To o.ITTOP_OPLNK.Count
  If o.ITTOP_OPLNK.Item(i).TheUser Is MyUser Then
   For j = 1 To o.ITTOP_OPLNK.Item(i).ITTOP_OPKAM.Count
     If camFilter <> "" Then
      camFilter = camFilter & " or "
      End If
      camFilter = camFilter & " location.code like '" & o.ITTOP_OPLNK.Item(i).ITTOP_OPKAM.Item(j).TheKamera.CameraMask & "' "
    Next
  End If
Next
 
 If camFilter <> "" Then
     camFilter = " and (" & camFilter & ") "
 End If

 If cliFilter <> "" Then
      cliFilter = " partner.code in ( " & cliFilter & ") "
 End If
 Dim Obj As DBuffer
 Set Obj = New DBuffer
 Obj.Name = camFilter
 Manager.AddCustomObjects Obj, "camFilter"
 
 Set Obj = New DBuffer
 Obj.Name = cliFilter
 Manager.AddCustomObjects Obj, "cliFilter"


  Set rs = Manager.ListInstances("", "ITTD")
  If Not rs.EOF Then
    id = rs!InstanceID
    Set ITTDic = Manager.GetInstanceObject(id)
  End If
 
 

End Sub

Private Sub mdiForm_Unload(Cancel As Integer)
On Error Resume Next

' whait for finalize timer loops
inTimer1 = True
Me.Timer1.Enabled = False

inTimer2 = True
Me.Timer2.Enabled = False


Timer1.Enabled = False
Timer2.Enabled = False

On Error Resume Next

' unload all dynamically created journals and reports
UnloadObjects

If Not frmFind Is Nothing Then
  Unload frmFind
End If
Set frmFind = Nothing

If Not frmFindFT Is Nothing Then
  Unload frmFindFT
End If
Set frmFindFT = Nothing



Dim f As Form
For Each f In Forms
  If f.MDIChild = True Then
    On Error Resume Next
    'Call f.Controls.Item(0).object.Init(Nothing, Nothing, False, Nothing)
    Unload f
  End If
Next

  For Each f In Forms
      On Error Resume Next
      Debug.Print f.Name
  Next
  
  
  Set MyRole = Nothing
  Set MyUser = Nothing
  Set usr = Nothing


  Session.Logout
  Set Session = Nothing
  Manager.CloseClass
  Set Manager = Nothing

  If Command$ <> "DEBUG" Then
   TerminateProcess GetCurrentProcess, 0
  'Else
  ' End
  End If
End Sub









Private Sub mnuCoreSetup_Click()
  Dim f As frmCoreSetup
  Set f = New frmCoreSetup
  f.Show vbModal
  Unload f
  Set f = Nothing
End Sub



 










Private Sub mnuOtbor_Click()
  

      Dim conn As ADODB.Connection
      Dim rs As ADODB.Recordset
      
      Set conn = Manager.GetCustomObjects("refref")
      Set rs = conn.Execute("select * from v_bami_vimorozka_rpt2 ")
      
      Set RptVimorozka2 = New ReportShow
      RptVimorozka2.ReportPath = App.Path & "\Otbor.rpt"
      Call RptVimorozka2.RunDirectRS(rs, False)
  
End Sub


Private Sub mnuExit_Click()
  Unload Me
End Sub




Private Sub mnuRpt103_Click()
    Dim conn As ADODB.Connection
    
    Set RptStok103 = New ReportShow
    RptStok103.ReportPath = App.Path & "\stok103.rpt"
    RptStok103.ReportSource = "v_bami_stock103"
    Set conn = Manager.GetCustomObjects("refref")
    
    Call RptStok103.Run(False, conn)
'    RptStok103.ExportPDF App.Path & "\" & "Stock103.pdf", conn
'
'    MailThisFile "Заблокированные поддоны", "Отчет по заблокированным поддонам на  " & Now & ".", App.Path & "\" & "Stock103.pdf"
    
End Sub

Private Sub mnuRptHran_Click()
 Dim conn As ADODB.Connection
    
    Set RptHran = New ReportShow
    RptHran.ReportPath = App.Path & "\Objem.rpt"
    RptHran.ReportSource = "v_bami_hranenie"
    Set conn = Manager.GetCustomObjects("refref")
    
    Call RptHran.Run(False, conn)
End Sub

Private Sub mnuRptOtobrano_Click()
 Dim conn As ADODB.Connection
    
    Set RptOtobrano = New ReportShow
    RptOtobrano.ReportPath = App.Path & "\otobrano.rpt"
    RptOtobrano.ReportSource = "v_bami_stockblocked"
    Set conn = Manager.GetCustomObjects("refref")
    
    Call RptOtobrano.Run(False, conn)
    
'    RptOtobrano.ExportPDF App.Path & "\otobrano.pdf", conn
'    MailThisFile "Отбор выморозки", "Отчет по отбору выморозки на  " & Now & ".", App.Path & "\" & "otobrano.pdf"
    
End Sub







Private Sub mnuVimorozka_Click()
    Dim f As frmDate
    Set f = New frmDate
    f.Show vbModal
    If f.OK Then
      Dim conn As ADODB.Connection
      Dim rs As ADODB.Recordset
      
      Set conn = Manager.GetCustomObjects("refref")
      Set rs = conn.Execute("select * from v_bami_vimorozka_rpt union all " & _
      "select partner_code ,item_id, item_code, description,  qin, qout, vimorozka * datediff(d,getdate()," & MakeMSSQLDate(f.dtpDate.Value) & "), pogreshnost,0, 0,0 from v_bami_stokmorozdayly")
      
      Set RptVimorozka = New ReportShow
      RptVimorozka.ReportPath = App.Path & "\Vimorozka.rpt"
      Call RptVimorozka.RunDirectRS(rs, False)
    End If
    Unload f
    Set f = Nothing
End Sub






Private Sub Timer2_Timer()
  If inTimer2 Then Exit Sub
  inTimer2 = True
  On Error Resume Next
  Call Session.Exec("SessionTouch", Nothing)
  inTimer2 = False
End Sub


Private Function NoTabs(ByVal s As String) As String
  NoTabs = Replace(Replace(Replace(Replace(s, vbTab, " "), vbCrLf, " "), vbCr, " "), vbLf, " ")
End Function



Public Function SynchronizeARMDescription()
    Dim objARM As Object
    Dim objMenuItem As Menu
    Dim ObjItem As Object

    Set objARM = Manager.GetInstanceObject(ARMID)
    
    Dim i As Long
    Dim objRS As ADODB.Recordset
    Dim objEntryPoint As Object
    
    For i = 0 To Me.Controls.Count - 1
        Set ObjItem = Me.Controls(i)
        If UCase(TypeName(ObjItem)) = UCase("menu") Then
            If ObjItem.Caption <> "-" Then
              Debug.Print "Found menu " + ObjItem.Caption + "-" + ObjItem.Name
              
              Set objRS = Session.GetRowsEx("EntryPoints", ARMID, , "Caption='" + ObjItem.Caption + "' or Name='" & ObjItem.Name & "'")
              If objRS.EOF And objRS.BOF Then
                  Set objEntryPoint = objARM.EntryPoints.Add
                  objEntryPoint.Caption = ObjItem.Caption
                  objEntryPoint.Name = ObjItem.Name
                  objEntryPoint.AsToolbarItem = Boolean_Net
                  objEntryPoint.ActionType = 0 'MenuActionType_Nicego_ne_delat_
                  objEntryPoint.save
                  If err.Number <> 0 Then
                    MsgBox err.Description
                  End If
                  err.Clear
              Else
                  Set objEntryPoint = objARM.FindRowObject("EntryPoints", objRS!Entrypointsid)
                  If Not objEntryPoint Is Nothing Then
                    objEntryPoint.Caption = ObjItem.Caption
                    objEntryPoint.Name = ObjItem.Name
                    objEntryPoint.AsToolbarItem = Boolean_Net
                    objEntryPoint.save
                  End If
              End If
              objRS.Close
            End If
        End If
    Next
End Function






Public Sub On_Load()
If Action = "" Then
   Me.Caption = App.FileDescription & " (" & Site & "\" & MyRole.Name & "\" & MyUser.Brief & ")"
   On Error Resume Next
   'If command$ <> "DEBUG" Then
     Dim c As Control
     For Each c In Me.Controls
      If TypeName(c) = "Menu" Then
         
        If CheckMenu(c.Name) = RoleMenuStatus_Hidden Then
          c.Visible = False
      
        End If
      End If
     Next
  'End If
  End If
   Manager.FreeAllInstanses
End Sub




'
'Private Sub mnuAbout_Click()
'frmAbout.Show vbModal, Me
'End Sub









Private Sub mnuITTOP_Click()
 Dim o As Object
 Dim rs  As ADODB.Recordset
 Dim id As String
  Set rs = Manager.ListInstances("", "ITTOP")
  If Not rs.EOF Then
    id = rs!InstanceID
  Else
    id = CreateGUID2
    Manager.NewInstance id, "ITTOP", "Операторы и кладовщики"
  End If
    Set o = Manager.GetInstanceObject(id)
    If IsDocDenied(o) Then
      MsgBox "Не разрешен доступ к документам такого типа"
      Exit Sub
    End If

    Dim g  As Object
    Set g = Manager.GetInstanceGUI(o.id)
    If Not g Is Nothing Then
      g.Show GetDocumentMode(o), o, False
    End If
  Set rs = Nothing
End Sub


Private Sub mnuITTD_Click()
 Dim o As Object
 Dim rs  As ADODB.Recordset
 Dim id As String
  Set rs = Manager.ListInstances("", "ITTD")
  If Not rs.EOF Then
    id = rs!InstanceID
  Else
    id = CreateGUID2
    Manager.NewInstance id, "ITTD", "Справочник"
  End If
    Set o = Manager.GetInstanceObject(id)
    If IsDocDenied(o) Then
      MsgBox "Не разрешен доступ к документам такого типа"
      Exit Sub
    End If

    Dim g  As Object
    Set g = Manager.GetInstanceGUI(o.id)
    If Not g Is Nothing Then
      g.Show GetDocumentMode(o), o, False
    End If
  Set rs = Nothing
End Sub



Private Sub mnuITTNO_Click()
 Dim o As Object
 Dim rs  As ADODB.Recordset
 Dim id As String
  Set rs = Manager.ListInstances("", "ITTNO")
  If Not rs.EOF Then
    id = rs!InstanceID
  Else
    id = CreateGUID2
    Manager.NewInstance id, "ITTNO", "Настройки оптмизатора"
  End If
    Set o = Manager.GetInstanceObject(id)
    If IsDocDenied(o) Then
      MsgBox "Не разрешен доступ к документам такого типа"
      Exit Sub
    End If

    Dim g  As Object
    Set g = Manager.GetInstanceGUI(o.id)
    If Not g Is Nothing Then
      g.Show GetDocumentMode(o), o, False
    End If
  Set rs = Nothing
End Sub

















Private Sub UnloadObjects()

On Error Resume Next



Set repShowOL = Nothing
Set repShowSRVOUT = Nothing
Set repShowKL = Nothing
Set repShowSRVIN = Nothing
Set repShowINEPL = Nothing
Set RptShowSRVALL = Nothing
Set RptStickers = Nothing
Set RptWrongLocation = Nothing
Set repShowMoves = Nothing
Set RptNedostacha = Nothing
Set RptActVes = Nothing
Set RptVimorozka = Nothing
Set RptVimorozka2 = Nothing
Set RptHran = Nothing
Set RptStok103 = Nothing
Set RptOtobrano = Nothing



End Sub

