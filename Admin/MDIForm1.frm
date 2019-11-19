VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Администратор."
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6075
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   2400
      Top             =   0
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFILE 
      Caption         =   "Файл"
      Begin VB.Menu mnuRoles 
         Caption         =   "Роли"
      End
      Begin VB.Menu mnuDictionaries 
         Caption         =   "Справочники"
      End
      Begin VB.Menu mnuJournals 
         Caption         =   "Журналы"
      End
      Begin VB.Menu mnuBrowser 
         Caption         =   "Обозреватель объектов"
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Пользователи"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "Активность пользователей"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDocuments 
         Caption         =   "Новый документ"
      End
      Begin VB.Menu mnuS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Выход"
      End
   End
   Begin VB.Menu mnuDocument 
      Caption         =   "Документ"
      Visible         =   0   'False
      Begin VB.Menu mnuDocLock 
         Caption         =   "Заблокировать"
      End
      Begin VB.Menu mnuDocUnlock 
         Caption         =   "Разблокировать"
      End
      Begin VB.Menu mnuDocSaveXML 
         Caption         =   "Сохранить в файл"
      End
      Begin VB.Menu mnuDocLoadXML 
         Caption         =   "Загрузить из файла"
      End
      Begin VB.Menu mnuDocSecure 
         Caption         =   "Установить права"
      End
      Begin VB.Menu mnuGetID 
         Caption         =   "Получить идентификатор"
      End
      Begin VB.Menu mnuDocRename 
         Caption         =   "Переименовать"
      End
      Begin VB.Menu mnuDocDelete 
         Caption         =   "Удалить"
      End
   End
   Begin VB.Menu mnutoolS 
      Caption         =   "Инструменты"
      Begin VB.Menu mnuMergeDocs 
         Caption         =   "Замена ссылки на документ"
      End
      Begin VB.Menu mnuMergeRow 
         Caption         =   "Замена ссылки"
      End
      Begin VB.Menu mnuSaveDocs 
         Caption         =   "Сохранить документы"
      End
      Begin VB.Menu mnuSetupJ 
         Caption         =   "Настройка журналов"
      End
      Begin VB.Menu mnuS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuADminUnlock 
         Caption         =   "Разблокировать"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Удалить настройки форм и журналов"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu nmuFind 
      Caption         =   "Поиск"
      Visible         =   0   'False
      Begin VB.Menu mnuFullText 
         Caption         =   "Полнотекстовый поиск"
      End
      Begin VB.Menu mnuFindAttr 
         Caption         =   "Поиск по атрибутам"
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "Окно"
      WindowList      =   -1  'True
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "Каскад"
      End
      Begin VB.Menu mnuTileVert 
         Caption         =   "Разложить вертикально"
      End
      Begin VB.Menu mnuTileHor 
         Caption         =   "Разложить горизонтально"
      End
      Begin VB.Menu mnuArrangeIcon 
         Caption         =   "Разложить иконки"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ObjectToReport As Object

Dim frmFind As Form
Dim frmFindFT As Form

Dim inTimer1 As Boolean
Dim inTimer2 As Boolean
Dim DelayedCommand As String
Dim WithEvents fRole As frmJournalShow
Attribute fRole.VB_VarHelpID = -1

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)


Private Sub fRole_OnAdd(usedefaut As Boolean, Refesh As Boolean)
Dim objGui  As Object
  Dim o As Object
  Dim ID As String
  ID = CreateGUID2
  Manager.NewInstance ID, "ROLES", "Роль " & Now, site
  Set o = Manager.GetInstanceObject(ID)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If
  Set objGui = Manager.GetInstanceGUI(o.ID)
  If Not objGui Is Nothing Then
    objGui.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

Private Sub fRole_OnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
usedefaut = False
If MsgBox("Актуализировать описания меню?", vbYesNo) = vbYes Then
  Dim objRole As ROLES.Application
  Set objRole = Manager.GetInstanceObject(fRole.jv.RowInstanceID(fRole.jv.Row))
  Dim objRWP As ROLES_WP
    Dim i As Long
    For i = 1 To objRole.ROLES_WP.Count
        Set objRWP = objRole.ROLES_WP.item(i)
        Dim objWP As MTZwp.Application
        Set objWP = Manager.GetInstanceObject(objRWP.WP.ID)
        If Not objWP Is Nothing Then
            If Not objWP.WorkPlace.item(1) Is Nothing Then
                'If objWP.WorkPlace.item(1).EntryPoints.Count <> objRWP.ROLES_ACT.Count Then
                    ' Загружаем меню
                    LoadMenus objRWP, objWP
                'End If
            End If
  
        End If
    Next
    MsgBox "Описание меню для роли актуализировано по описанию АРМ"
  End If
End Sub


Private Sub LoadLevelEP(objWPEP As MTZwp.EntryPoints_COL, objREP As ROLES.ROLES_ACT_COL)
Dim i As Long
Dim objEP As ROLES.ROLES_ACT
Dim bChanged As Boolean
    bChanged = False
    
    For i = 1 To objWPEP.Count
        objREP.Filter = "EntryPoints='" + objWPEP.item(i).ID + "'"
        objREP.Refresh
    
        If Not objREP.Count > 0 Then
            Set objEP = objREP.Add(CreateGUID2)
            Set objEP.EntryPoints = objWPEP.item(i) '.ID
            objEP.Accesible = YesNo_Da
            objEP.Save
            bChanged = True
        Else
            Set objEP = objREP.item(1)
        End If
        
        If Not objEP Is Nothing Then
            If objWPEP.item(i).EntryPoints.Count > 0 Then
                LoadLevelEP objWPEP.item(i).EntryPoints, objEP.ROLES_ACT
            End If
        End If
    Next
    
    objREP.Filter = ""
    objREP.Refresh
    For i = objREP.Count To 1 Step -1
        If objREP.item(i).EntryPoints Is Nothing Then
            objREP.item(i).Delete
            bChanged = True
        ElseIf objWPEP.item(objREP.item(i).EntryPoints.ID) Is Nothing Then
            objREP.item(i).Delete
            bChanged = True
        End If
    Next
    If bChanged Then
        objREP.Application.Save
        objREP.Application.BatchUpdate
    End If
    
End Sub

Private Sub LoadMenus(objRWP As ROLES.ROLES_WP, objWP As MTZwp.Application)
Dim i As Long
Dim objEP As ROLES.ROLES_ACT
Dim objWP2 As MTZwp.Application
Dim bChanged As Boolean
    bChanged = False
    LoadLevelEP objWP.EntryPoints, objRWP.ROLES_ACT

End Sub

Private Sub mdiform_load()
  'CreateIcon
  DeltaReminder = GetSetting("MTZ", "CONFIG", "REMINDER", "00:15:00")
  Me.Caption = Me.Caption & " (" & site & "\" & MyUser.Brief & ")"
  
   Dim c As Control
    For Each c In Me.Controls
     If TypeName(c) = "Menu" Then
        
       If CheckMenu(c.Name) = RoleMenuStatus_Hidden Then
         c.Visible = False
       End If
     End If
    Next
End Sub

Private Sub mdiForm_Unload(Cancel As Integer)
On Error Resume Next


inTimer2 = True
Me.Timer2.Enabled = False

ReminderVisible = True
Timer2.Enabled = False


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
    Call f.Controls.item(0).object.Init(Nothing, Nothing, False, Nothing)
    Unload f
  End If
Next


Set MyUser = Nothing
Set usr = Nothing
Set model = Nothing
Erase Data



Session.Logout
Set Session = Nothing
Manager.CloseClass
Set Manager = Nothing

If Command$ <> "DEBUG" Then
 TerminateProcess GetCurrentProcess, 0
End If

End Sub



Private Sub mnuAbout_Click()
  frmAbout.Show vbModal
  'SynchronizeARMDescription
End Sub

Private Sub mnuADminUnlock_Click()
  If MsgBox("Будут отменены все блокировки документов." & vbCrLf & "Разблокировать документы ?", vbYesNo + vbQuestion, "ВНИМАНИЕ") = vbYes Then
    On Error GoTo bye
    Dim v As NamedValues
    Set v = New NamedValues
    Call Session.Exec("AdminUnlockAll", v)
  End If
  Exit Sub
bye:
  MsgBox Err.Description
End Sub

Private Sub mnuArrangeIcon_Click()
Me.Arrange vbArrangeIcons
End Sub


Private Sub mnuBrowser_Click()
CallSys
End Sub

Private Sub mnuCascade_Click()
Me.Arrange vbCascade
End Sub



Private Sub mnuDelete_Click()
  Dim tmpdir As String, tmpfile As String
  
  tmpdir = GetSetting("MTZ", "CONFIG", "LAYOUTS", App.path & "\LAYOUTS\")
  tmpfile = Dir(tmpdir & "*.*", vbNormal)
  
  Do While tmpfile <> ""   ' Start the loop.
     
        If (GetAttr(tmpdir & tmpfile) And vbDirectory) = 0 Then
           'Debug.Print tmpfile   ' Display entry only if it
           
           On Error Resume Next
           Kill tmpdir & tmpfile
           Debug.Print Err.Description
           
        End If   ' it represents a directory.
     
     tmpfile = Dir   ' Get next entry.
  Loop
  MsgBox "Настройки форм и журналов удалены"
  
End Sub

Private Sub mnuDocRename_Click()
On Error Resume Next
  Dim item As Object
  Set item = GetActiveItem()
  If item Is Nothing Then Exit Sub
  Dim s As String
  s = InputBox("Введите новое название документа", "Переименовать документ", item.Name)
  If s <> "" Then
    item.Name = s
    item.Save
  End If
  

End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub



Private Sub mnuGetID_Click()
On Error GoTo bye
  Dim item As Object
  Set item = GetActiveItem()
  If item Is Nothing Then Exit Sub
  Clipboard.Clear
  Clipboard.SetText item.ID, vbCFText
  frmShowID.Label1 = "Идентификатор документа :" & vbCrLf & item.Brief
  frmShowID.Text1 = item.ID
  frmShowID.Show vbModal
  Exit Sub
bye:

End Sub

Private Sub mnuLog_Click()
frmLog.Show
End Sub


Private Sub mnuMergeDocs_Click()
  frmMergeObjTool.Show vbModal
End Sub

Private Sub mnuMergeRow_Click()
  frmMergeRowTool.Show vbModal
End Sub

'Private Sub mnuMetaModel_Click()
'  Dim o As Object, g As Object
'  Dim rs As ADODB.Recordset
'
'  Set rs = Manager.ListInstances(site, "MTZMetaModel")
'  Set o = Manager.GetInstanceObject(rs!InstanceID)
'
'
'  Set g = Manager.GetInstanceGUI(rs!InstanceID)
'  g.Show "", o, False
'  Set rs = Nothing
'
'End Sub


Private Sub mnuPWU_Click()
Dim o As Object, g As Object
  Dim rs As ADODB.Recordset
  Dim ID As String
  Set rs = Manager.ListInstances(site, "PWU")
  If rs.EOF Then
    ID = CreateGUID2
    Manager.NewInstance ID, "PWU", "Пользователи сайта"
  Else
    ID = rs!InstanceID
  End If
  Set o = Manager.GetInstanceObject(ID)
  
  
  Set g = Manager.GetInstanceGUI(ID)
  g.Show "", o, False
  Set rs = Nothing
End Sub

Private Sub mnuRoles_Click()
    Dim ID As String
    
    Set journal = model.Manager.GetInstanceObject("{DB8F8C01-D05A-44B6-B80C-16A6B7AA65D6}")
    If Not journal Is Nothing Then
      Manager.LockInstanceObject journal.ID
      
      Set fRole = New frmJournalShow
      Set fRole.jv.journal = journal
      fRole.jv.AllowAdd = True
      fRole.jv.AllowDel = True
      fRole.jv.AllowFilter = False
      
      fRole.Caption = journal.Name
      fRole.Show
      fRole.jv.Refresh
    
    End If
End Sub

Private Sub mnuSaveDocs_Click()
frmSaveTool.Show
End Sub



Private Sub mnuTileHor_Click()
Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVert_Click()
Me.Arrange vbTileVertical
End Sub



Private Sub mnuUsers_Click()
  Dim o As Object, g As Object
  Dim rs As ADODB.Recordset

  Set rs = Manager.ListInstances(site, "MTZUsers")
  Set o = Manager.GetInstanceObject(rs!InstanceID)
  
  
  Set g = Manager.GetInstanceGUI(rs!InstanceID)
  g.Show "", o, False
  Set rs = Nothing
End Sub




Private Sub Timer2_Timer()
If inTimer2 Then Exit Sub
inTimer2 = True
On Error Resume Next
Session.SessionTouch
If Not GetActiveItem() Is Nothing Then
  mnuDocument.Visible = True
Else
  mnuDocument.Visible = False
End If

inTimer2 = False

End Sub



Private Function NoTabs(ByVal s As String) As String
  NoTabs = Replace(Replace(Replace(Replace(s, vbTab, " "), vbCrLf, " "), vbCr, " "), vbLf, " ")
End Function



Private Sub mnuDictionaries_Click()
  Set frmDicList.model = model
  frmDicList.Show vbModal
  If frmDicList.OK Then
    Dim ot As OBJECTTYPE
    Set ot = model.FindRowObject("OBJECTTYPE", frmDicList.result)
    Dim o1 As Object, o2 As Object, ID As String
    Dim rs As ADODB.Recordset
    Set rs = Manager.ListInstances(site, ot.Name)
    If rs.EOF Then
      ID = CreateGUID2
      Manager.NewInstance ID, ot.Name, ot.the_comment, site
    Else
      ID = rs!InstanceID
    End If
    
    Set o1 = Manager.GetInstanceObject(ID, site)
    If o1 Is Nothing Then
      MsgBox "Отсутствует объектная библиотека для типа:" & ot.the_comment, vbCritical + vbOKOnly
      Exit Sub
    End If
    Set o2 = Manager.GetInstanceGUI(o1.ID)
    If o2 Is Nothing Then
      MsgBox "Отсутствует интерфейсный компонент для типа:" & ot.the_comment, vbCritical + vbOKOnly
      Exit Sub
    End If
    o2.Show "", o1, False
  End If
End Sub

Private Sub mnuDocuments_Click()
  Set frmDocList.model = model
  frmDocList.Show vbModal
  If frmDocList.OK Then
    Dim ot As OBJECTTYPE
    Set ot = model.FindRowObject("OBJECTTYPE", frmDocList.result)
    Dim o1 As Object, o2 As Object, ID As String
    Dim rs As ADODB.Recordset
    ID = CreateGUID2
    Manager.NewInstance ID, ot.Name, ot.the_comment & " " & Now, site
    Set o1 = Manager.GetInstanceObject(ID, site)
    If o1 Is Nothing Then
      MsgBox "Отсутствует объектная библиотека для типа:" & ot.the_comment, vbCritical + vbOKOnly
      Exit Sub
    End If
    Set o2 = Manager.GetInstanceGUI(o1.ID)
    If o2 Is Nothing Then
      MsgBox "Отсутствует интерфейсный компонент для типа:" & ot.the_comment, vbCritical + vbOKOnly
      Exit Sub
    End If
    o2.Show "", o1, False
  End If
End Sub

Private Sub mnuJournals_Click()
  Set frmJouralList.model = model
  frmJouralList.Show vbModal
  If frmJouralList.OK Then
    Set journal = model.Manager.GetInstanceObject(frmJouralList.result)
    Dim f As frmJournalShow
    Set f = New frmJournalShow
    Set f.jv.journal = journal
    f.jv.AllowAdd = False
    f.jv.AllowDel = False
    f.jv.AllowFilter = False
    
    f.Caption = journal.Name
    f.Show
    f.jv.Refresh
  End If
End Sub

Private Sub mnuSetupJ_Click()
  Set frmJouralList.model = model
  frmJouralList.Show vbModal
  If frmJouralList.OK Then
    Set journal = model.Manager.GetInstanceObject(frmJouralList.result)
    Set frmJournalConfig.JournalDef1.model = model
    Set frmJournalConfig.JournalDef1.journal = journal
    frmJournalConfig.Show vbModal
    Unload frmJournalConfig
  End If
End Sub

Private Sub mnuDocDelete_Click()
On Error GoTo bye
  Dim item As Object
  Set item = GetActiveItem()
  If item Is Nothing Then Exit Sub
  If MsgBox("Удалить документ?", vbQuestion + vbYesNo) = vbYes Then
  
    item.UnLockResource
    item.WorkOffline = False
    item.Manager.DeleteInstance item.ID
    item.Manager.FreeInstanceObject item.ID
    Unload Me.ActiveForm
  End If
  Exit Sub
bye:
   MsgBox Err.Description, vbCritical, "Ошибка при удалении"
End Sub

Private Sub mnuDocLoadXML_Click()
 On Error Resume Next
 Dim item As Object
  Set item = GetActiveItem()
  If item Is Nothing Then Exit Sub
  If item.Application.MTZSession.CheckRight(item.SecureStyleID, "XMLLOAD") Then
  
  On Error GoTo bye
  Dim fn As String
  cdlg.CancelError = True
  cdlg.Filter = "Документ XML |*.XML"
  cdlg.DefaultExt = "XML"
  cdlg.FileName = App.path & "\" & item.ID & ".xml"
  cdlg.Flags = cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNFileMustExist
  cdlg.ShowOpen
  fn = cdlg.FileName
  
  Dim xdom As MSXML2.DOMDocument
  Set xdom = New MSXML2.DOMDocument
  xdom.Load fn
  item.XMLLoad xdom.lastChild, 1
  item.WorkOffline = False
  item.BatchUpdate
  Set xdom = Nothing
  
 End If
bye:
End Sub

Private Sub mnuDocLock_Click()
  On Error Resume Next
  Dim item As Object
  Set item = GetActiveItem()
  If item Is Nothing Then Exit Sub
  item.LockResource True
  Me.ActiveForm.TestLock
End Sub

Private Sub mnuDocSaveXML_Click()
 On Error Resume Next
 Dim item As Object
 Set item = GetActiveItem()
 If item Is Nothing Then Exit Sub
 
 If item.Application.MTZSession.CheckRight(item.SecureStyleID, "XMLSAVE") Then
 
  On Error GoTo bye
  Dim fn As String
  cdlg.CancelError = True
  cdlg.Filter = "Документ XML|*.XML"
  cdlg.DefaultExt = "XML"
  cdlg.FileName = App.path & "\" & item.ID & ".xml"
  cdlg.Flags = cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNOverwritePrompt 'cdlOFNFileMustExist
  cdlg.ShowSave
  fn = cdlg.FileName
   item.LockResource True
   item.LoadAll
   item.WorkOffline = True
   Dim xdom As MSXML2.DOMDocument
   Set xdom = New MSXML2.DOMDocument
   xdom.loadXML "<root></root>"
   item.XMLSave xdom.lastChild, xdom
   xdom.Save fn
   item.WorkOffline = False
 End If
bye:
End Sub

Private Sub mnuDocSecure_Click()
On Error Resume Next
  Dim item As Object
  Set item = GetActiveItem()
  If item Is Nothing Then Exit Sub
  item.Application.Manager.ShowSecurityDialog item
End Sub

Private Sub mnuDocUnlock_Click()
On Error Resume Next
  Dim item As Object
  Set item = GetActiveItem()
  If item Is Nothing Then Exit Sub
  If item.IsLocked Then
    item.UnLockResource
  Else
  MsgBox "Объект не заблокирован", vbInformation
  End If
  Me.ActiveForm.TestLock
End Sub

Private Function GetActiveItem() As Object
On Error Resume Next
If TypeName(Me.ActiveForm) = "frmObj" Then
  Set GetActiveItem = Me.ActiveForm.item
End If
End Function

Private Sub KillTypes(item As MTZMetaModel.MTZAPP)
  Dim ot As Object, i As Long
  On Error GoTo bye
tryagain:
  item.Application.OBJECTTYPE.Refresh
  For i = 1 To item.Application.OBJECTTYPE.Count
    Set ot = item.Application.OBJECTTYPE.item(i)
    If ot.Package.ID = item.ID Then
      item.Application.OBJECTTYPE.Delete ot.ID
      GoTo tryagain
    End If
  Next
bye:
End Sub

Private Function SynchronizeARMDescription()
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
                  objEntryPoint.Save
              Else
                  Set objEntryPoint = objARM.FindRowObject("EntryPoints", objRS!Entrypointsid)
                  If Not objEntryPoint Is Nothing Then
                    objEntryPoint.Caption = ObjItem.Caption
                    objEntryPoint.Name = ObjItem.Name
                    objEntryPoint.AsToolbarItem = Boolean_Net
                    objEntryPoint.Save
                  End If
              End If
              objRS.Close
            End If
        End If
    Next
End Function
