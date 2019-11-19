VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{5859921D-9E04-476B-B692-7FBC18917E50}#4.1#0"; "ROLESGUI.ocx"
Begin VB.Form frmRoles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Роли"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   Icon            =   "frmRoles.frx":0000
   LinkTopic       =   "Роли"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cbRefresh 
      Height          =   330
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "refresh.ico"
      Top             =   120
      Width           =   330
   End
   Begin ROLESGUI.vpnROLES_REPORTS_ vpnROLES_REPORTS_1 
      Height          =   5415
      Left            =   3120
      TabIndex        =   10
      Top             =   480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   9551
   End
   Begin ROLESGUI.vpnROLES_DOC_ vpnROLES_DOC_1 
      Height          =   5415
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   9551
   End
   Begin ROLESGUI.vpnROLES_MAP_ vpnROLES_MAP_1 
      Height          =   5415
      Left            =   3120
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   9551
   End
   Begin VB.CommandButton cbProp 
      Height          =   330
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "PROP.ico"
      ToolTipText     =   "Свойства"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cbDelete 
      Height          =   330
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "DELETE.ico"
      Top             =   120
      Width           =   330
   End
   Begin ROLESGUI.vpnROLES_WP_ vpnROLES_WP_1 
      Height          =   5415
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   9551
   End
   Begin VB.CommandButton cbAddNew 
      Height          =   330
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "NEW.ico"
      Top             =   120
      Width           =   330
   End
   Begin ROLESGUI.vpnROLES_DEF_ vpnROLES_DEF_1 
      Height          =   5415
      Left            =   3120
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   9551
   End
   Begin ROLESGUI.vpnROLES_USER_ vpnROLES_USER_1 
      Height          =   5415
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   9551
   End
   Begin GridEX20.GridEX GridEXRole 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   9763
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      ItemCount       =   0
      DataMode        =   99
      ColumnHeaderHeight=   285
      ColumnsCount    =   1
      Column(1)       =   "frmRoles.frx":000C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmRoles.frx":011C
      FormatStyle(2)  =   "frmRoles.frx":0278
      FormatStyle(3)  =   "frmRoles.frx":0328
      FormatStyle(4)  =   "frmRoles.frx":03DC
      FormatStyle(5)  =   "frmRoles.frx":04B4
      FormatStyle(6)  =   "frmRoles.frx":056C
      ImageCount      =   0
      PrinterProperties=   "frmRoles.frx":064C
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5895
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   10398
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Группы"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Пользователи"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Приложения"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Документы"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Отчёты"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset
Public Manager As MTZManager.Main
Public Session As MTZSession.Session
Public IsFirstPaint As Boolean
Private objWPS As Collection
Private Roles As Collection


Private Sub cbAddNew_Click()
Dim ID As String
    ID = CreateGUID2
    Call Manager.NewInstance(ID, "ROLES", "Описание ролей")
    RefreshGrid
End Sub

Private Sub cbProp_Click()
Dim objRole As Roles.Application
Dim objRolesGUI As ROLESGUI.GUI
    If Roles.Count > 0 Then
        Set objRole = Roles.item("Key" + CStr(GridEXRole.Row))
        Set objRolesGUI = Manager.GetInstanceGUI(objRole.ID, "")
        Call objRolesGUI.Show("", objRole)
    End If
End Sub

Private Sub cbRefresh_Click()
  Dim objRole As Roles.Application

    If Roles.Count > 0 Then
        Set objRole = Roles.item("Key" + CStr(GridEXRole.Row))
        On Error Resume Next
        On Error GoTo 0

        'objRole.ROLES_WP.item(ID).ROLES_ACT.Filter = "WorkPlaceid='" + WPID + "'"
        vpnROLES_USER_1.OnInit objRole
        vpnROLES_USER_1.OnClick objRole, Me
        
        vpnROLES_MAP_1.OnInit objRole
        vpnROLES_MAP_1.OnClick objRole, Me
        
        vpnROLES_DOC_1.OnInit objRole
        vpnROLES_DOC_1.OnClick objRole, Me
        
        vpnROLES_DEF_1.OnInit objRole
        vpnROLES_DEF_1.OnClick objRole, Me
        
        vpnROLES_WP_1.OnInit objRole
        vpnROLES_WP_1.OnClick objRole, Me
        
        vpnROLES_REPORTS_1.OnInit objRole
        vpnROLES_REPORTS_1.OnClick objRole, Me
        
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
        
    End If
End Sub

Private Sub Form_Load()
    IsFirstPaint = True
    Set Roles = New Collection
    If IsFirstPaint Then
        IsFirstPaint = False
        LoadBtnPictures cbDelete, cbDelete.Tag
        LoadBtnPictures cbAddNew, cbAddNew.Tag
        LoadBtnPictures cbProp, cbProp.Tag
        LoadBtnPictures cbRefresh, cbRefresh.Tag
        RefreshGrid
    End If
End Sub

Private Sub RefreshGrid()
    If Manager Is Nothing Or Session Is Nothing Then Exit Sub
    Dim i As Long
    Dim ID As String
    i = 0
    GridEXRole.ItemCount = 0
    On Error Resume Next
    Set rs = Manager.ListInstances("", "Roles")
    rs.MoveFirst
    Set Roles = Nothing
    Set Roles = New Collection
    On Error GoTo 0
    While Not rs.EOF
        i = i + 1
        Dim objRoles As Roles.Application
        ID = rs!InstanceID
        Set objRoles = Manager.GetInstanceObject(ID)
        Roles.Add objRoles, "Key" + CStr(i)
        rs.MoveNext
    Wend
    GridEXRole.ItemCount = i
    GridEXRole.Rebind
    DoEvents
    GridEXRole_Click
    ChekSelected
End Sub

Private Sub Form_Paint()
'    If IsFirstPaint Then
'        IsFirstPaint = False
'        LoadBtnPictures cbDelete, cbDelete.Tag
'        LoadBtnPictures cbAddNew, cbAddNew.Tag
'        LoadBtnPictures cbProp, cbProp.Tag
'        RefreshGrifd
'    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set Roles = Nothing
End Sub

Private Sub GridEXRole_AfterColEdit(ByVal ColIndex As Integer)
    GridEXRole_EndCustomEdit ColIndex
End Sub

Private Sub GridEXRole_Click()

Dim objRole As Roles.Application

    If Roles.Count > 0 Then
        Set objRole = Roles.item("Key" + CStr(GridEXRole.Row))
        On Error Resume Next
        On Error GoTo 0

        'objRole.ROLES_WP.item(ID).ROLES_ACT.Filter = "WorkPlaceid='" + WPID + "'"
        vpnROLES_USER_1.OnInit objRole
        vpnROLES_USER_1.OnClick objRole, Me
        
        vpnROLES_MAP_1.OnInit objRole
        vpnROLES_MAP_1.OnClick objRole, Me
        
        vpnROLES_DOC_1.OnInit objRole
        vpnROLES_DOC_1.OnClick objRole, Me
        
        vpnROLES_DEF_1.OnInit objRole
        vpnROLES_DEF_1.OnClick objRole, Me
        
        vpnROLES_WP_1.OnInit objRole
        vpnROLES_WP_1.OnClick objRole, Me
        
        vpnROLES_REPORTS_1.OnInit objRole
        vpnROLES_REPORTS_1.OnClick objRole, Me
        
''''        Dim objRWP As ROLES_WP
''''        Dim i As Long
''''        For i = 1 To objRole.ROLES_WP.Count
''''            Set objRWP = objRole.ROLES_WP.item(i)
''''            Dim objWP As MTZwp.Application
''''            Set objWP = Manager.GetInstanceObject(objRWP.WP.ID)
''''            If Not objWP Is Nothing Then
''''                If Not objWP.WorkPlace.item(1) Is Nothing Then
''''                    'If objWP.WorkPlace.item(1).EntryPoints.Count <> objRWP.ROLES_ACT.Count Then
''''                        ' Загружаем меню
''''                        LoadMenus objRWP, objWP
''''                    'End If
''''                End If
''''
''''            End If
''''        Next
        
    End If
End Sub

Private Sub LoadLevelEP(objWPEP As MTZwp.EntryPoints_COL, objREP As Roles.ROLES_ACT_COL)
Dim i As Long
Dim objEP As Roles.ROLES_ACT
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

Private Sub LoadMenus(objRWP As Roles.ROLES_WP, objWP As MTZwp.Application)
Dim i As Long
Dim objEP As Roles.ROLES_ACT
Dim objWP2 As MTZwp.Application
Dim bChanged As Boolean
    bChanged = False
    LoadLevelEP objWP.EntryPoints, objRWP.ROLES_ACT
'    For i = 1 To objWP.WorkPlace.item(1).EntryPoints.Count
'        objRWP.ROLES_ACT.Filter = "EntryPoints='" + objWP.WorkPlace.item(1).EntryPoints.item(i).ID + "'"
'        objRWP.ROLES_ACT.Refresh
'        If Not objRWP.ROLES_ACT.Count > 0 Then
'            Set objEP = objRWP.ROLES_ACT.Add(CreateGUID2)
'            Set objEP.EntryPoints = objWP.WorkPlace.item(1).EntryPoints.item(i) '.ID
'            objEP.Accesible = YesNo_Da
'            objEP.Save
'            bChanged = True
'        Else
'            Set objEP = objRWP.ROLES_ACT.item(1)
'        End If
'        If Not objEP Is Nothing Then
'            If objWP.WorkPlace.item(1).EntryPoints.item(i).EntryPoints.Count > 0 Then
'                LoadLevelEP objWP.WorkPlace.item(1).EntryPoints.item(i).EntryPoints, objEP
'            End If
'        End If
'    Next
'    objRWP.ROLES_ACT.Filter = ""
'    objRWP.ROLES_ACT.Refresh
'    For i = objRWP.ROLES_ACT.Count To 1 Step -1
'        If objWP.WorkPlace.item(1).EntryPoints.item(objRWP.ROLES_ACT.item(i).EntryPoints.ID) Is Nothing Then
'            objRWP.ROLES_ACT.item(i).Delete
'            bChanged = True
'        End If
'    Next
'    If bChanged Then
'        objRWP.Save
'        objRWP.BatchUpdate
'    End If
End Sub

Private Sub GridEXRole_DblClick()
'
'Dim objRole As Roles.Application
'
'    If Roles.Count > 0 Then
'        Set objRole = Roles.item("Key" + CStr(GridEXRole.Row))
'        Dim objRolesGUI As ROLESGUI.GUI
'        Set objRolesGUI = Manager.GetInstanceGUI(objRole.ID, "")
'        On Error Resume Next
'        Call objRolesGUI.Show("", objRole)
'        On Error GoTo 0
'        vpnROLES_USER_1.OnClick objRole, Me
'    End If
End Sub

Private Function GEditItem()
Dim objRole As Roles.Application
On Error GoTo ErrorExit
    If Roles.Count > 0 Then
        Set objRole = Roles.item("Key" + CStr(GridEXRole.Row))
        If objRole.Name <> GridEXRole.Value(GridEXRole.col) Then
            objRole.Name = GridEXRole.Value(GridEXRole.col)
            If objRole.ROLES_DEF.Count = 0 Then
                objRole.ROLES_DEF.Add CreateGUID2
            End If
            objRole.ROLES_DEF.item(1).Name = objRole.Name
            objRole.ROLES_DEF.item(1).Save
            objRole.Save
        End If
    End If
ErrorExit:
End Function

Private Sub GridEXRole_EndCustomEdit(ByVal ColIndex As Integer)
    GEditItem
End Sub

Private Sub GridEXRole_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal objRoles As GridEX20.JSRowData)
    Dim objRole As Roles.Application
    If Roles.Count > 0 Then
        Set objRole = Roles.item("Key" + CStr(RowIndex))
        objRole.ROLES_DEF.Refresh
        If objRole.ROLES_DEF.Count > 0 Then
            objRoles.Value(1) = objRole.ROLES_DEF.item(1).Name
        Else
            objRoles.Value(1) = objRole.Brief
        End If
        On Error Resume Next
        
        Debug.Print CStr(objRole.ROLES_DEF.Count)
        On Error GoTo 0
    End If
End Sub

Private Sub ChekSelected()
    Select Case TabStrip1.SelectedItem.Index
        Case 1:
            vpnROLES_MAP_1.Visible = True
            vpnROLES_USER_1.Visible = False
            vpnROLES_DOC_1.Visible = False
            vpnROLES_DEF_1.Visible = False
            vpnROLES_WP_1.Visible = False
            vpnROLES_REPORTS_1.Visible = False
        Case 2:
            vpnROLES_MAP_1.Visible = False
            vpnROLES_USER_1.Visible = True
            vpnROLES_DOC_1.Visible = False
            vpnROLES_DEF_1.Visible = False
            vpnROLES_WP_1.Visible = False
            vpnROLES_REPORTS_1.Visible = False
        Case 3:
            vpnROLES_MAP_1.Visible = False
            vpnROLES_USER_1.Visible = False
            vpnROLES_DOC_1.Visible = False
            vpnROLES_DEF_1.Visible = False
            vpnROLES_WP_1.Visible = True
            vpnROLES_REPORTS_1.Visible = False
        Case 4:
            vpnROLES_MAP_1.Visible = False
            vpnROLES_USER_1.Visible = False
            vpnROLES_DOC_1.Visible = True
            vpnROLES_DEF_1.Visible = False
            vpnROLES_WP_1.Visible = False
            vpnROLES_REPORTS_1.Visible = False
        Case 5:
            vpnROLES_MAP_1.Visible = False
            vpnROLES_USER_1.Visible = False
            vpnROLES_DOC_1.Visible = False
            vpnROLES_DEF_1.Visible = False
            vpnROLES_WP_1.Visible = False
            vpnROLES_REPORTS_1.Visible = True
    End Select
End Sub

Private Sub GridEXRole_Validate(Cancel As Boolean)
    GEditItem
End Sub

Private Sub TabStrip1_Click()
    ChekSelected
End Sub
