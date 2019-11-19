VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{1801C003-859D-471D-BF31-D4428050324B}#2.1#0"; "MTZ_PANEL.ocx"
Begin VB.UserControl vpnITTD_SRV_admi 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer menuTimerITTD_SRV 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   1560
   End
   Begin VB.Frame pnlITTD_SRV 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      Begin VB.CommandButton cmdITTD_SRVAcc 
         Height          =   330
         Left            =   3075
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "security.ico"
         ToolTipText     =   "����� ������"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdITTD_SRVFnd 
         Height          =   330
         Left            =   2325
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "find.ico"
         ToolTipText     =   "�����"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdITTD_SRVCfg 
         Height          =   330
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "config.ico"
         ToolTipText     =   "��������� �������� ����"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdITTD_SRVPrn 
         Height          =   330
         Left            =   1575
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "PrintPreview.ico"
         ToolTipText     =   "������"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdITTD_SRVRef 
         Height          =   330
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "Refresh.ico"
         ToolTipText     =   "��������"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdITTD_SRVDel 
         Height          =   330
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "DELETE.ico"
         ToolTipText     =   "�������"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdITTD_SRVEdit 
         Height          =   330
         Left            =   450
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "PROP.ico"
         ToolTipText     =   "��������"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdITTD_SRVAdd 
         Height          =   330
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "NEW.ico"
         ToolTipText     =   "��������"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin GridEX20.GridEX gridITTD_SRV 
         Height          =   480
         Left            =   3900
         TabIndex        =   1
         Top             =   2760
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   847
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigatorString=   "������:|��"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ScrollToolTipColumn=   ""
         TabKeyBehavior  =   1
         HideSelection   =   2
         SelectionStyle  =   1
         GroupByBoxInfoText=   "�����������"
         AllowDelete     =   -1  'True
         AllowEdit       =   0   'False
         ItemCount       =   0
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "vpnITTD_SRV_admi.ctx":0000
         Column(2)       =   "vpnITTD_SRV_admi.ctx":00C8
         FormatStylesCount=   5
         FormatStyle(1)  =   "vpnITTD_SRV_admi.ctx":016C
         FormatStyle(2)  =   "vpnITTD_SRV_admi.ctx":02C8
         FormatStyle(3)  =   "vpnITTD_SRV_admi.ctx":0378
         FormatStyle(4)  =   "vpnITTD_SRV_admi.ctx":042C
         FormatStyle(5)  =   "vpnITTD_SRV_admi.ctx":0504
         ImageCount      =   0
         PrinterProperties=   "vpnITTD_SRV_admi.ctx":05BC
      End
      Begin MTZ_PANEL.DropButton cmdITTD_SRVRun 
         Height          =   330
         Left            =   2700
         TabIndex        =   9
         Tag             =   "run.ico"
         ToolTipText     =   "�������"
         Top             =   30
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         Caption         =   ""
      End
   End
   Begin VB.Menu mnuPopupITTD_SRV 
      Caption         =   "���� ��� ������"
      Visible         =   0   'False
      Begin VB.Menu mnuITTD_SRVAdd 
         Caption         =   "�������"
      End
      Begin VB.Menu mnuITTD_SRVEdit 
         Caption         =   "��������"
      End
      Begin VB.Menu mnuITTD_SRVDelete 
         Caption         =   "�������"
      End
      Begin VB.Menu mnuITTD_SRVOpen 
         Caption         =   "�������"
      End
      Begin VB.Menu mnuITTD_SRVRef 
         Caption         =   "��������"
      End
      Begin VB.Menu mnuITTD_SRVS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuITTD_SRVBuffer 
         Caption         =   "�����"
         Begin VB.Menu mnuITTD_SRVCOPY 
            Caption         =   "����������"
         End
         Begin VB.Menu mnuITTD_SRVCUT 
            Caption         =   "��������"
         End
         Begin VB.Menu mnuITTD_SRVPASTE 
            Caption         =   "��������"
         End
         Begin VB.Menu mnuITTD_SRVCHANGE 
            Caption         =   "��������"
         End
         Begin VB.Menu mnuITTD_SRVCLEAN 
            Caption         =   "��������"
         End
      End
      Begin VB.Menu mnuITTD_SRVPrn 
         Caption         =   "������"
      End
      Begin VB.Menu mnuITTD_SRVCfg 
         Caption         =   "���������"
      End
      Begin VB.Menu mnuITTD_SRVFind 
         Caption         =   "�����"
      End
      Begin VB.Menu mnuITTD_SRVAcc 
         Caption         =   "����� ������"
      End
   End
End
Attribute VB_Name = "vpnITTD_SRV_admi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'������ ��� ������ �� �������� �������
Public item As Object
Public ParentForm As Object
Public ModalMode As Boolean
Private fndITTD_SRV As Object
Private menuActionITTD_SRV As String





Private Sub cmdITTD_SRVCfg_Click()
  On Error Resume Next
Dim jset As Object 'MTZ_JSetup.Customizer
Set jset = CreateObject("MTZ_JSetup.Customizer")
jset.ShowSumary gridITTD_SRV, "gridITTD_SRV"
Set jset = Nothing
End Sub

Private Sub cmdITTD_SRVPrn_Click()
  On Error Resume Next
Dim jset As Object 'MTZ_JSetup.Customizer
Set jset = CreateObject("MTZ_JSetup.Customizer")
jset.Preview gridITTD_SRV, "������"
Set jset = Nothing
End Sub

Private Sub cmdITTD_SRVRef_Click()
  On Error Resume Next
  item.ITTD_SRV.Refresh
  gridITTD_SRV.ItemCount = item.ITTD_SRV.Count
  gridITTD_SRV.Refetch
  gridITTD_SRV.Refresh
End Sub
Private Sub cmdITTD_SRVAcc_Click()
On Error Resume Next
If gridITTD_SRV.ItemCount = 0 Then Exit Sub
If gridITTD_SRV.Row > 0 Then
 If gridITTD_SRV.RowIndex(gridITTD_SRV.Row) > 0 Then
    Dim u As Object
    Dim bm
    bm = gridITTD_SRV.RowBookmark(gridITTD_SRV.RowIndex(gridITTD_SRV.Row))
    Set u = item.FindRowObject(Right(bm, Len(bm) - 38), Left(bm, 38))
    item.Application.Manager.ShowSecurityDialog u
 End If
End If
End Sub
Private Sub cmdITTD_SRVEdit_Click()
  On Error Resume Next
If gridITTD_SRV.ItemCount = 0 Then Exit Sub
If gridITTD_SRV.Row > 0 Then
 If gridITTD_SRV.RowIndex(gridITTD_SRV.Row) > 0 Then
    Dim u As Object
    Dim bm
    bm = gridITTD_SRV.RowBookmark(gridITTD_SRV.RowIndex(gridITTD_SRV.Row))
    Set u = item.FindRowObject(Right(bm, Len(bm) - 38), Left(bm, 38))
    Set frmITTD_SRV_admi.item = u
again:     frmITTD_SRV_admi.NotFirstTime = False
    frmITTD_SRV_admi.OnInit
    frmITTD_SRV_admi.Show vbModal
    If frmITTD_SRV_admi.OK Then
    On Error Resume Next
    err.Clear
    u.Save
    If err.Number <> 0 Then
      MsgBox err.Description, vbOKOnly + vbExclamation, "���������"
      GoTo again
    End If
    gridITTD_SRV.RefreshRowBookmark bm
    Else
     u.Refresh
    End If
 End If
End If
End Sub
Private Sub cmdITTD_SRVDel_Click()
  On Error Resume Next
  gridITTD_SRV.Delete
End Sub
Private Sub cmdITTD_SRVAdd_Click()
  On Error Resume Next
      Dim u As Object
      Set u = item.ITTD_SRV.Add()

      Set frmITTD_SRV_admi.item = u
again:       frmITTD_SRV_admi.NotFirstTime = False
      frmITTD_SRV_admi.OnInit
      frmITTD_SRV_admi.Show vbModal
      If frmITTD_SRV_admi.OK Then
        err.Clear
        u.Save
        If err.Number <> 0 Then
          MsgBox err.Description, vbOKOnly + vbExclamation, "��������"
          GoTo again
        End If
        gridITTD_SRV.ItemCount = item.ITTD_SRV.Count
        gridITTD_SRV.Refresh
        gridITTD_SRV.RefreshGroups
        gridITTD_SRV.RefreshSort
        gridITTD_SRV.MoveToBookmark u.ID & "ITTD_SRV"
      Else
        item.ITTD_SRV.Delete u.ID
        item.ITTD_SRV.Remove u.ID
      End If
      Exit Sub
bye:
   MsgBox err.Description, vbOKOnly + vbExclamation, "��������"
End Sub
Private Sub cmdITTD_SRVFnd_Click()
  On Error Resume Next
  fndITTD_SRV.ShowForm
End Sub
Private Sub cmdITTD_SRVRun_Click()
  On Error Resume Next
  gridITTD_SRV_DblClick
End Sub
Public Function IsOK() As Boolean
        IsOK = True
End Function
Private Sub gridITTD_SRV_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error Resume Next
item.ITTD_SRV.LoadRow gridITTD_SRV, RowIndex, Bookmark, Values
End Sub
Private Sub gridITTD_SRV_BeforeDeleteEX(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Cancel As GridEX20.JSRetBoolean)
  If gridITTD_SRV.ItemCount = 0 Then Exit Sub
  Dim u As Object
  Set u = item.FindRowObject("ITTD_SRV", Left(Bookmark, 38))
  If Not u Is Nothing Then
    On Error Resume Next
    If MsgBox("������� �������" & vbCrLf & u.brief & vbCrLf & "?", vbYesNo + vbQuestion, "��������") = vbYes Then
      On Error GoTo bye
      u.Parent.Delete u.ID
    Else
      Cancel = True
    End If
  End If
Exit Sub
bye:
   Cancel = True
   MsgBox err.Description, vbOKOnly + vbExclamation, "��������"
End Sub

Private Sub gridITTD_SRV_DblClick()
  cmdITTD_SRVEdit_Click
End Sub
Private Sub gridITTD_SRV_KeyPress(KeyAscii As Integer)
  If gridITTD_SRV.ItemCount = 0 Then Exit Sub
  On Error Resume Next
  fndITTD_SRV.FindOnPress KeyAscii
End Sub

Private Sub gridITTD_SRV_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
 On Error Resume Next
If LastRow <> gridITTD_SRV.Row And LastRow > 0 Then
  gridITTD_SRV.GetRowData(LastRow).RowStyle = "Default"
  gridITTD_SRV.GetRowData(gridITTD_SRV.Row).RowStyle = "SelectedRow"
End If
End Sub
Private Sub mnuITTD_SRVPASTE_Click()
On Error Resume Next
    Dim u As Object
    Set u = item.ITTD_SRV.Add()
    If GetFromBuffer(u) Then
      gridITTD_SRV.ItemCount = u.Parent.Count
    End If
End Sub
'Hidden GRID Setup...
Private Sub gridITTD_SRV_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown, AltDown, CtrlDown
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyF4 Then
        If ShiftDown And CtrlDown And AltDown Then
            cmdITTD_SRVCfg_Click
        End If
    End If
End Sub

Private Sub gridITTD_SRV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 And Shift = 0 Then
  On Error Resume Next
  PopupMenu mnuPopupITTD_SRV
End If
End Sub
Private Sub mnuITTD_SRVAdd_click()
   menuActionITTD_SRV = "ADD"
   menuTimerITTD_SRV.Enabled = True
End Sub
Private Sub mnuITTD_SRVEdit_click()
   menuActionITTD_SRV = "EDIT"
   menuTimerITTD_SRV.Enabled = True
End Sub
Private Sub mnuITTD_SRVDelete_click()
   menuActionITTD_SRV = "DEL"
   menuTimerITTD_SRV.Enabled = True
End Sub
Private Sub mnuITTD_SRVOpen_click()
   menuActionITTD_SRV = "RUN"
   menuTimerITTD_SRV.Enabled = True
End Sub
Private Sub mnuITTD_SRVRef_click()
   menuActionITTD_SRV = "REF"
   menuTimerITTD_SRV.Enabled = True
End Sub
  Private Sub mnuITTD_SRVCOPY_Click()
    On Error Resume Next
    If gridITTD_SRV.ItemCount = 0 Then Exit Sub
    If gridITTD_SRV.Row > 0 Then
     If gridITTD_SRV.RowIndex(gridITTD_SRV.Row) > 0 Then
        Dim u As Object
        Dim bm
        bm = gridITTD_SRV.RowBookmark(gridITTD_SRV.RowIndex(gridITTD_SRV.Row))
        Set u = item.FindRowObject(Right(bm, Len(bm) - 38), Left(bm, 38))
        SaveToBuffer u
      End If
    End If
  End Sub
  Private Sub mnuITTD_SRVCUT_Click()
    On Error Resume Next
    If gridITTD_SRV.ItemCount = 0 Then Exit Sub
    If gridITTD_SRV.Row > 0 Then
     If gridITTD_SRV.RowIndex(gridITTD_SRV.Row) > 0 Then
        Dim u As Object
        Dim bm
        bm = gridITTD_SRV.RowBookmark(gridITTD_SRV.RowIndex(gridITTD_SRV.Row))
        Set u = item.FindRowObject(Right(bm, Len(bm) - 38), Left(bm, 38))
        SaveToBuffer u
        Set bm = u.Parent
        u.Parent.Delete u.ID
        gridITTD_SRV.ItemCount = bm.Count
      End If
    End If
  End Sub
Private Sub mnuITTD_SRVCHANGE_Click()
  On Error Resume Next
  If gridITTD_SRV.ItemCount = 0 Then Exit Sub
  If gridITTD_SRV.Row > 0 Then
   If gridITTD_SRV.RowIndex(gridITTD_SRV.Row) > 0 Then
    Dim u As Object
    Dim bm
    bm = gridITTD_SRV.RowBookmark(gridITTD_SRV.RowIndex(gridITTD_SRV.Row))
    Set u = item.FindRowObject(Right(bm, Len(bm) - 38), Left(bm, 38))
    If GetFromBuffer(u) Then
      gridITTD_SRV.ItemCount = u.Parent.Count
      gridITTD_SRV.RefreshRowIndex gridITTD_SRV.RowIndex(gridITTD_SRV.Row)
    End If
   End If
  End If
End Sub
  Private Sub mnuITTD_SRVCLEAN_Click()
    item.Manager.ClearBuffer " & p.name & "
  End Sub
Private Sub mnuITTD_SRVPrn_click()
   menuActionITTD_SRV = "PRN"
   menuTimerITTD_SRV.Enabled = True
End Sub
Private Sub mnuITTD_SRVCfg_click()
   menuActionITTD_SRV = "CFG"
   menuTimerITTD_SRV.Enabled = True
End Sub
Private Sub mnuITTD_SRVFind_click()
   menuActionITTD_SRV = "FND"
   menuTimerITTD_SRV.Enabled = True
End Sub
Private Sub mnuITTD_SRVAcc_click()
   menuActionITTD_SRV = "ACC"
   menuTimerITTD_SRV.Enabled = True
End Sub
Private Sub menuTimerITTD_SRV_timer()
   menuTimerITTD_SRV.Enabled = False
   If menuActionITTD_SRV = "ADD" Then cmdITTD_SRVAdd_Click
   If menuActionITTD_SRV = "EDIT" Then cmdITTD_SRVEdit_Click
   If menuActionITTD_SRV = "DEL" Then cmdITTD_SRVDel_Click
   If menuActionITTD_SRV = "RUN" Then cmdITTD_SRVRun_Click
   If menuActionITTD_SRV = "REF" Then cmdITTD_SRVRef_Click
   If menuActionITTD_SRV = "PRN" Then cmdITTD_SRVPrn_Click
   If menuActionITTD_SRV = "CFG" Then cmdITTD_SRVCfg_Click
   If menuActionITTD_SRV = "FND" Then cmdITTD_SRVFnd_Click
   If menuActionITTD_SRV = "ACC" Then cmdITTD_SRVAcc_Click
   menuActionITTD_SRV = ""
End Sub

'������� ������������� ������ �������
'Parameters:
'[IN][OUT]  aItem , ��� ���������: object,
'[IN][OUT]   optional pForm , ��� ���������: object = nothing  - ...
'See Also:
'Example:
'  call me.OnInit({���������})
Public Sub OnInit(aItem As Object, Optional pForm As Object = Nothing)
 On Error Resume Next
 Set item = aItem
 Set ParentForm = pForm

  LoadBtnPictures cmdITTD_SRVAdd, cmdITTD_SRVAdd.Tag
  LoadBtnPictures cmdITTD_SRVEdit, cmdITTD_SRVEdit.Tag
  LoadBtnPictures cmdITTD_SRVDel, cmdITTD_SRVDel.Tag
  LoadBtnPictures cmdITTD_SRVRef, cmdITTD_SRVRef.Tag
  LoadBtnPictures cmdITTD_SRVPrn, cmdITTD_SRVPrn.Tag
  LoadBtnPictures cmdITTD_SRVFnd, cmdITTD_SRVFnd.Tag
  LoadBtnPictures cmdITTD_SRVRun, cmdITTD_SRVRun.Tag
  LoadBtnPictures cmdITTD_SRVAcc, cmdITTD_SRVAcc.Tag
  LoadBtnPictures cmdITTD_SRVCfg, cmdITTD_SRVCfg.Tag
  item.ITTD_SRV.PrepareGrid gridITTD_SRV
  LoadGridLayout gridITTD_SRV
  Set fndITTD_SRV = Nothing
  On Error Resume Next
  Set fndITTD_SRV = CreateObject("MTZ_JSetup.GridFinder")
  fndITTD_SRV.Init gridITTD_SRV
End Sub

'�������� ��� ����������� ����
'Parameters:
' ���������� ���
'See Also:
'Example:
'  call me.OnTabClick()
Private Sub OnTabClick()

      gridITTD_SRV.Top = 40 * Screen.TwipsPerPixelX
      gridITTD_SRV.Left = 5 * Screen.TwipsPerPixelX
      gridITTD_SRV.Width = UserControl.Width - 10 * Screen.TwipsPerPixelX
      gridITTD_SRV.Height = UserControl.Height - 45 * Screen.TwipsPerPixelY
      gridITTD_SRV.Top = 40 * Screen.TwipsPerPixelX
      gridITTD_SRV.Left = 5 * Screen.TwipsPerPixelX
      gridITTD_SRV.Width = UserControl.Width - 10 * Screen.TwipsPerPixelX
      gridITTD_SRV.Height = UserControl.Height - 45 * Screen.TwipsPerPixelY
End Sub

'�������� ��� ����������� ������
'Parameters:
'[IN][OUT]  aItem , ��� ���������: object,
'[IN][OUT]   optional pForm , ��� ���������: object = nothing  - ...
'See Also:
'Example:
'  call me.OnClick({���������})
Public Sub OnClick(aItem As Object, Optional pForm As Object = Nothing)
  On Error Resume Next
  Set item = aItem
  Set ParentForm = pForm
  OnTabClick
End Sub

'�������� ��� ����������
'Parameters:
' ���������� ���
'See Also:
'Example:
'  call me.OnSave()
Public Sub OnSave()
  
End Sub

'������� ���������
'Parameters:
' ���������� ���
'Returns:
' Boolean, ��������� ����������:
'   true  -
'   false -
'See Also:
'Example:
' dim variable as boolean
'  variable = me.IsChanged()
Public Function IsChanged() As Boolean
  Dim m_IsChanged As Boolean
  On Error Resume Next
  
  IsChanged = m_IsChanged
End Function
Private Sub UserControl_Resize()
  On Error Resume Next
  pnlITTD_SRV.Move 0, 0, UserControl.Width, UserControl.Height
  OnTabClick
End Sub

'������������ ��������
'Parameters:
' ���������� ���
'See Also:
'Example:
'  call me.CloseClass()
Public Sub CloseClass()
  On Error Resume Next
  Set ParentForm = Nothing
  Set item = Nothing
  Set fndITTD_SRV = Nothing
End Sub
Private Sub UserControl_Terminate()
  On Error Resume Next
  CloseClass
End Sub

'����������� ������� � ��������
'{5CB1388C-1623-4B36-A775-00B70BEE27AF}
Private Sub Run_VBMoveVRC(VRCATFolder As Variant, Optional RowItem As Object)
On Error Resume Next

'do nothing
End Sub



'��� �� ��������� Square >0
'{53371FFA-B514-447A-A1F9-26EE4FD409C9}
Private Sub Run_VBUpdateObjNamePEO(Name As Variant, Optional RowItem As Object)
On Error Resume Next

On Error Resume Next
RowItem.Application.Name = Name
RowItem.Application.Save

End Sub



'�������� ����� ������� ���������� ���� �� ������� ������
'{79DED4FD-045C-45F8-AC79-2E5A3D956D97}
Private Sub Run_VBMigrateRight(Optional RowItem As Object)
On Error Resume Next

'Migrate security
On Error GoTo bye
  If RowItem.Person Is Nothing Then Exit Sub

  RowItem.Person.Secure item.SecureStyleID
  RowItem.Person.Propagate
bye:
  Exit Sub
  MsgBox err.Description, vbOKOnly + vbCritical, "�������� ����� ������"
End Sub



'�������� ������� ��������� �������������
'{79E6BDEB-91D5-4B2E-81F7-3E091FB65E3A}
Private Sub Run_VBCheckDescrs(DesPartName As Variant, Optional RowItem As Object)
On Error Resume Next

    On Error Resume Next

    Dim part_col As Variant
    Dim part_item As Object
    Set part_col = CallByName(RowItem.Parent.Parent, DesPartName, VbGet, False)
    Set part_item = part_col.item(1)
    If RowItem.Parent.Count = 1 Then
        If part_item.HasDescrs = -1 Then
            part_item.HasDescrs = 0
        Else
            part_item.HasDescrs = -1
        End If
    End If
    part_item.Save

End Sub



'��������� ����� ������ ��� ������� ������������ ��� �������� ��������
'{31EC6CF7-8DBD-4EFE-BF12-4D168F653D34}
Private Sub Run_VBApplySecurity(Optional RowItem As Object)
On Error Resume Next

'Apply security
On Error GoTo bye
  If RowItem.Client Is Nothing Then Exit Sub
  If RowItem.Parent.Parent.AccessLevel Is Nothing Then Exit Sub
  RowItem.Client.Secure RowItem.Parent.Parent.AccessLevel.ID
  RowItem.Client.Propagate
  Exit Sub
bye:
  MsgBox err.Description, vbOKOnly + vbCritical, "�������� ����� ������"
End Sub



'�������� ������ �� ������ ��������
'{5B8FB7B9-D8B1-4CA0-90AF-55F83D1A6E5D}
Private Sub Run_VBMakeReport(ReportType As Variant, Optional RowItem As Object)
On Error Resume Next

On Error GoTo bye
Dim ID As String
 Dim Obj As Object
 'ID = CreateGUID2
 'Call RowItem.Application.Manager.NewInstance(ID, "VRRPT", "����� " & Date)
 'Set RowItem.Report = RowItem.Application.Manager.GetInstanceObject(ID)
 If RowItem.Report.VRRPT_MAIN.Count = 0 Then
  Set Obj = RowItem.Report.VRRPT_MAIN.Add
 Else
  Set Obj = RowItem.Report.VRRPT_MAIN.item(1)
 End If
 
 Set Obj.Author = RowItem.Application.FindRowObject("Users", item.Application.MTZSession.GetSessionUserID())
 Obj.TheDate = Date
 
 If ReportType = "CLNT" Then
   Set Obj.Client = RowItem.Application
 End If
 If ReportType = "PRJ" Then
   Set Obj.Project = RowItem.Application
 End If
 If ReportType = "CONT" Then
   Set Obj.Contract = RowItem.Application
 End If
 If ReportType = "PERS" Then
   Set Obj.Person = RowItem.Application
 End If
 
 Obj.Save
 RowItem.Save

 Exit Sub
bye:
  MsgBox err.Description, vbOKOnly + vbCritical, "�������� ������"
End Sub



'
'{AA4085E6-745B-4A37-8EC4-65D99A653966}
Private Sub Run_VBRemoveSymmetricObjRef(ForwardFieldName As Variant, ObjTypeName As Variant, SymmetricPartName As Variant, SymmetricFieldName As Variant, Optional RowItem As Object)
On Error Resume Next

    Dim OK As Boolean
    Dim ID As String
    Dim brief As String

    On Error Resume Next
'     On Error GoTo bye

    Dim Obj As Object
    Dim part_col As Variant
    Dim part_item As Object
    Set Obj = CallByName(RowItem, ForwardFieldName, VbGet)
    Set part_col = CallByName(Obj, SymmetricPartName, VbGet, False)
    part_col.Filter = SymmetricFieldName + "='" + RowItem.Application.ID + "'"
'    RowItem.Parent.Remove RowItem.ID
    Set part_item = part_col.item(1)
    part_col.Delete part_item.ID
    Exit Sub
bye:
Resume

End Sub



'������� ����� ������ �� �������
'{42A1A436-8AA2-4F1F-999B-6680DFF514DE}
Private Sub Run_VBNewPayIn(Optional RowItem As Object)
On Error Resume Next

On Error GoTo bye
Dim ID As String
 Dim Obj As Object
 ID = CreateGUID2
 Call RowItem.Application.Manager.NewInstance(ID, "PEKP", "������ " & Date)
 Set RowItem.TheDocument = RowItem.Application.Manager.GetInstanceObject(ID)

 If RowItem.TheDocument.PEKP_DEF.Count = 0 Then
  Set Obj = RowItem.TheDocument.PEKP_DEF.Add
 Else
  Set Obj = RowItem.TheDocument.PEKP_DEF.item(1)
 End If
 
 Set Obj.FromClient = RowItem.Application
 Obj.PLPDate = Date
 Obj.Save
 RowItem.Save

 Exit Sub
bye:
  MsgBox err.Description, vbOKOnly + vbCritical, "�������� �������"
End Sub



'��� �������, � ������� ���� ���� - ������, �������/�������� ������ ��� �������� ���� �������������� ������
'{A2EEE876-54D8-4AED-B124-775F5DA2D911}
Private Sub Run_VBAddObjByRef(ObjTypeName As Variant, ForwardFieldName As Variant, SymmetricPartName As Variant, SymmetricFieldName As Variant, Optional RowItem As Object)
On Error Resume Next

    Dim OK As Boolean
    Dim ID As String
    Dim brief As String
    Dim Mode As String
    Dim ResObject As Object
    On Error Resume Next
    Mode = Mid(TypeName(Me), InStr(TypeName(Me), "_") + 1)
    
'     On Error GoTo bye
    If Len(Mode) = 0 Then
        OK = item.Application.Manager.GetObjectListDialogEx(ID, brief, "", ObjTypeName)
    Else
        ID = CreateGUID2
        If Len(ObjTypeName) = 0 Then
            Dim newObj As Object
            Set newObj = item.Application.Manager.GetNewObject
            If Not (newObj Is Nothing) Then
                OK = True
                ID = newObj.ID
            End If
        Else
            OK = item.Application.Manager.NewInstance(ID, ObjTypeName, "")
        End If
        Dim ref As Object, objGui As Object
        Set ref = item.Application.Manager.GetInstanceObject(ID)
        If Not ref Is Nothing Then
          Set objGui = item.Application.Manager.GetInstanceGUI(ID)
          If objGui Is Nothing Then Set ref = Nothing: Exit Sub
          objGui.Show "", ref, False
          Set objGui = Nothing
        Else
          OK = False
        End If
    End If
    Dim Obj As Object
    Set Obj = item.Application.Manager.GetInstanceObject(ID)
    If Obj Is Nothing Then
        OK = False
    End If
    If OK Then
  Dim Coll As New Collection
        Dim part_col As Variant
        Dim part_item As Object
        CallByName RowItem, ForwardFieldName, VbSet, Obj
        Coll.Add TypeName(RowItem) + ":" + RowItem.ID
        RowItem.Save
        If Len(SymmetricPartName) > 0 And Len(SymmetricFieldName) > 0 Then
          Set part_col = CallByName(Obj, SymmetricPartName, VbGet, True)
          Set part_item = part_col.Add
          CallByName part_item, SymmetricFieldName, VbSet, RowItem.Application
          part_item.Save
          Coll.Add SymmetricPartName + ":" + part_item.ID
        End If
        Call item.Application.Manager.AddCustomObjects(Coll, Obj.ID)
    Else
        RowItem.Parent.Remove RowItem.ID
    End If
    Exit Sub
bye:
Resume
End Sub



'
'{5B376AF5-339B-4365-BA80-785E28BCF4DA}
Private Sub Run_VBUpdateSymmetricObjRef(SymmetricFieldName As Variant, ForwardFieldName As Variant, SymmetricPartName As Variant, ObjTypeName As Variant, Optional RowItem As Object)
On Error Resume Next

 
End Sub



'������� ����� ������ �� �������
'{2BB30818-90ED-4627-8ABB-85B3FBA46750}
Private Sub Run_VBNewPayOut(Optional RowItem As Object)
On Error Resume Next

On Error GoTo bye
Dim ID As String
 Dim Obj As Object
 ID = CreateGUID2
 Call RowItem.Application.Manager.NewInstance(ID, "PEKO", "������ " & Date)
 Set RowItem.TheDocument = RowItem.Application.Manager.GetInstanceObject(ID)

 If RowItem.TheDocument.PEKO_DEF.Count = 0 Then
  Set Obj = RowItem.TheDocument.PEKO_DEF.Add
 Else
  Set Obj = RowItem.TheDocument.PEKO_DEF.item(1)
 End If
 
 Set Obj.ToClient = RowItem.Application
 Obj.PLPDate = Date
 Obj.Save
 RowItem.Save

 Exit Sub
bye:
  MsgBox err.Description, vbOKOnly + vbCritical, "�������� �������"
End Sub



'�������� �������� �� ��������������
'{94E8F6DB-106A-44DC-9483-86C801798FF0}
Private Sub Run_VBOpenRef(StartMode As Variant, ID As Variant, Optional RowItem As Object)
On Error Resume Next

On Error Resume Next
If ID <> "" Then
    Dim Obj As Object
    Set Obj = item.Manager.GetInstanceObject(ID)
    If Not Obj Is Nothing Then
      Dim objGui As Object
      Set objGui = item.Manager.GetInstanceGUI(Obj.ID)
      If objGui Is Nothing Then Exit Sub

       If StartMode = "AUTO" Then
        StartMode = ""
        Dim i As Long
        For i = 100 To 0 Step -10
          If Obj.MTZSession.CheckRight(Obj.SecureStyleID, Obj.TypeName & ":" & "M" & i) Then
            StartMode = "M" & i
            Exit For
          End If
        Next
       End If
      
      objGui.Show StartMode & "", Obj
      Set objGui = Nothing
    End If
  End If
  
End Sub



'
'{069956DC-3305-45EF-9331-91CE323B5942}
Private Sub Run_WFDefName(Optional RowItem As Object)
On Error Resume Next

On Error Resume Next
item.Name = RowItem.Description
ParentForm.Caption = item.Name
item.Save
End Sub



'
'{D8914FB4-6B5D-491A-A72F-985617727583}
Private Sub Run_WFFuncName(Optional RowItem As Object)
On Error Resume Next

On Error Resume Next
item.Name = RowItem.Name
ParentForm.Caption = item.Name
item.Save
End Sub



'���������� ����� �������
'{61393545-ABF7-46F7-82F3-9B7E610DD9C0}
Private Sub Run_VBUpdateObjName(Name As Variant, Optional RowItem As Object)
On Error Resume Next

On Error Resume Next
RowItem.Application.Name = Name
RowItem.Application.Save
End Sub



'����� ������� ����� ������� � �������� ������������� ���������� (���������� RealEstate) ��� �������� ������ �������� � ���� ��������� ������ ��, ����� �������� � ���������� �� ��������������
'{477B8D25-4FF7-491A-A0B0-D3437EC16957}
Private Sub Run_MakeNewFolderEC(FolderID As Variant, Optional RowItem As Object)
On Error Resume Next

On Error GoTo bye
 Dim ID As String
 Dim Obj As Object ' EstComplex.Application ' Object
 Dim GObj As Object
 Dim fold As Object 'EstCatalog.Application ' Object
 
 ID = CreateGUID2
 Call RowItem.Application.Manager.NewInstance(ID, "EstComplex", RowItem.TheName & " " & Date)
 Set Obj = RowItem.Application.Manager.GetInstanceObject(ID)
 If Obj.EC_Def.Count = 0 Then
    With Obj.EC_Def.Add
        .TheName = RowItem.TheName
    End With
 Else
    Obj.EC_Def.item(1).TheName = RowItem.TheName
 End If
 Obj.Save
 Set RowItem.LinkedEC = Obj
 RowItem.Save
 Set GObj = RowItem.Application.Manager.GetInstanceGUI(Obj.ID)
 GObj.Show "", Obj, True 'False
 Set Obj = RowItem.Application.Manager.GetInstanceObject(ID)
 RowItem.TheName = Obj.EC_Def.item(1).TheName
 RowItem.Save
 Exit Sub
bye:
  MsgBox err.Description, vbOKOnly + vbCritical, "�������� ��"

End Sub



'�������� �������� �������
'{4FB59D1A-0123-47D3-9F4F-E12085C5D074}
Private Sub Run_VBUpdateItemName(Name As Variant, Optional RowItem As Object)
On Error Resume Next

On Error Resume Next
item.Name = Name
' ����� �� ���� �����, ���� ���������� ������ ActiveX
ParentForm.Caption = item.Name
item.Save
End Sub



'������� ������ �� �������
'{49EA5CBF-93CF-41A8-B1F1-E37FE4D59EA5}
Private Sub Run_VBNewZayavka(Optional RowItem As Object)
On Error Resume Next

On Error GoTo bye
Dim ID As String
 Dim Obj As Object
 ID = CreateGUID2
 Call RowItem.Application.Manager.NewInstance(ID, "PEKZ", "������ " & Date)
 Set RowItem.TheDocument = RowItem.Application.Manager.GetInstanceObject(ID)

 If RowItem.TheDocument.PEKO_DEF.Count = 0 Then
  Set Obj = RowItem.TheDocument.PEKZ_DEF.Add
 Else
  Set Obj = RowItem.TheDocument.PEKZ_DEF.item(1)
 End If
 
 Set Obj.ClientFrom = RowItem.Application
 Obj.QueryDate = Date
 Obj.Save
 RowItem.Save

 Exit Sub
bye:
  MsgBox err.Description, vbOKOnly + vbCritical, "�������� ������"
End Sub



'�������� ������������ �� �������
'{B91ABF3A-31F8-4A82-8D41-EF463DBA32D0}
Private Sub Run_SSCreateNomen(Name As Variant, Optional RowItem As Object)
On Error Resume Next

'pointCreateLine
End Sub





