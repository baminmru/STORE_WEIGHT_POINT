VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ctlmain_main 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Tag             =   "Card.ICO"
   Begin MSComctlLib.TabStrip ts 
      Height          =   1500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ITTDGUI.vpnITTD_GTYPE_main pnlITTD_GTYPE 
      Height          =   1500
      Left            =   1500
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_ZTYPE_main pnlITTD_ZTYPE 
      Height          =   1500
      Left            =   3000
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_PLTYPE_main pnlITTD_PLTYPE 
      Height          =   1500
      Left            =   4500
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_QTYPE_main pnlITTD_QTYPE 
      Height          =   1500
      Left            =   6000
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_ATYPE_main pnlITTD_ATYPE 
      Height          =   1500
      Left            =   0
      TabIndex        =   5
      Top             =   1500
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_SRV_main pnlITTD_SRV 
      Height          =   1500
      Left            =   1500
      TabIndex        =   6
      Top             =   1500
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_PART_main pnlITTD_PART 
      Height          =   1500
      Left            =   3000
      TabIndex        =   7
      Top             =   1500
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_FACTORY_main pnlITTD_FACTORY 
      Height          =   1500
      Left            =   4500
      TabIndex        =   8
      Top             =   1500
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_KILLPLACE_main pnlITTD_KILLPLACE 
      Height          =   1500
      Left            =   6000
      TabIndex        =   9
      Top             =   1500
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_COUNTRY_main pnlITTD_COUNTRY 
      Height          =   1500
      Left            =   0
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_CAMERA_main pnlITTD_CAMERA 
      Height          =   1500
      Left            =   1500
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_RULE_main pnlITTD_RULE 
      Height          =   1500
      Left            =   3000
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_OPTTYPE_main pnlITTD_OPTTYPE 
      Height          =   1500
      Left            =   4500
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_MOROZ_main pnlITTD_MOROZ 
      Height          =   1500
      Left            =   6000
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin ITTDGUI.vpnITTD_EMAIL_main pnlITTD_EMAIL 
      Height          =   1500
      Left            =   0
      TabIndex        =   15
      Top             =   4500
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
   End
   Begin VB.Menu mnuCtl 
      Caption         =   "mnuCtl"
      Visible         =   0   'False
      Begin VB.Menu mnuSetup 
         Caption         =   "���������"
      End
   End
End
Attribute VB_Name = "ctlmain_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' ������� ����������� ������� ��� �������������� ��������� ����� main
Public item As Object
Public Host As GUI
Public ModalMode As Boolean
Public ParentForm As Object
Private TSCustom As MTZ_CUSTOMTAB.TabStripCustomizer






'�������� ������ ������ ��������
'Parameters:
' ���������� ���
'Returns:
'  �������� ���� Long
'See Also:
'Example:
' dim variable as Long
' variable = me. PrefferedWidth
Public Property Get PrefferedWidth() As Long
    PrefferedWidth = 0
End Property


'������ ������ ������ ��������
'Parameters:
' ���������� ���
'Returns:
'  �������� ���� Long
'See Also:
'Example:
' dim variable as Long
' variable = me. PrefferedHeight
Public Property Get PrefferedHeight() As Long
    PrefferedHeight = 0
End Property

Private Sub mnuSetup_Click()
TSCustom.Setup ts
End Sub
Private Sub ts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 And Shift = 0 Then
    PopupMenu mnuCtl
  End If
End Sub

'��� ������ ��������� ���������
'Parameters:
' ���������� ���
'Returns:
' Boolean, ��������� ����������:
'   true  -
'   false -
'See Also:
'Example:
' dim variable as BooLEAN
'  variable = me.IsOK()
Public Function IsOK() As Boolean
Dim IsItOk As Boolean
IsItOk = True
On Error Resume Next
If IsItOk Then IsItOk = pnlITTD_GTYPE.IsOK
If IsItOk Then IsItOk = pnlITTD_ZTYPE.IsOK
If IsItOk Then IsItOk = pnlITTD_PLTYPE.IsOK
If IsItOk Then IsItOk = pnlITTD_QTYPE.IsOK
If IsItOk Then IsItOk = pnlITTD_ATYPE.IsOK
If IsItOk Then IsItOk = pnlITTD_SRV.IsOK
If IsItOk Then IsItOk = pnlITTD_PART.IsOK
If IsItOk Then IsItOk = pnlITTD_FACTORY.IsOK
If IsItOk Then IsItOk = pnlITTD_KILLPLACE.IsOK
If IsItOk Then IsItOk = pnlITTD_COUNTRY.IsOK
If IsItOk Then IsItOk = pnlITTD_CAMERA.IsOK
If IsItOk Then IsItOk = pnlITTD_RULE.IsOK
If IsItOk Then IsItOk = pnlITTD_OPTTYPE.IsOK
If IsItOk Then IsItOk = pnlITTD_MOROZ.IsOK
If IsItOk Then IsItOk = pnlITTD_EMAIL.IsOK
IsOK = IsItOk
End Function
Private Sub ts_click()
  On Error Resume Next
  pnlITTD_EMAIL.Visible = False
  pnlITTD_MOROZ.Visible = False
  pnlITTD_OPTTYPE.Visible = False
  pnlITTD_RULE.Visible = False
  pnlITTD_CAMERA.Visible = False
  pnlITTD_COUNTRY.Visible = False
  pnlITTD_KILLPLACE.Visible = False
  pnlITTD_FACTORY.Visible = False
  pnlITTD_PART.Visible = False
  pnlITTD_SRV.Visible = False
  pnlITTD_ATYPE.Visible = False
  pnlITTD_QTYPE.Visible = False
  pnlITTD_PLTYPE.Visible = False
  pnlITTD_ZTYPE.Visible = False
  pnlITTD_GTYPE.Visible = False

   Select Case ts.SelectedItem.Key
   Case "ITTD_GTYPE"
     With pnlITTD_GTYPE
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_GTYPE.OnClick item, ParentForm
     End With
   Case "ITTD_ZTYPE"
     With pnlITTD_ZTYPE
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_ZTYPE.OnClick item, ParentForm
     End With
   Case "ITTD_PLTYPE"
     With pnlITTD_PLTYPE
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_PLTYPE.OnClick item, ParentForm
     End With
   Case "ITTD_QTYPE"
     With pnlITTD_QTYPE
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_QTYPE.OnClick item, ParentForm
     End With
   Case "ITTD_ATYPE"
     With pnlITTD_ATYPE
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_ATYPE.OnClick item, ParentForm
     End With
   Case "ITTD_SRV"
     With pnlITTD_SRV
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_SRV.OnClick item, ParentForm
     End With
   Case "ITTD_PART"
     With pnlITTD_PART
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_PART.OnClick item, ParentForm
     End With
   Case "ITTD_FACTORY"
     With pnlITTD_FACTORY
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_FACTORY.OnClick item, ParentForm
     End With
   Case "ITTD_KILLPLACE"
     With pnlITTD_KILLPLACE
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_KILLPLACE.OnClick item, ParentForm
     End With
   Case "ITTD_COUNTRY"
     With pnlITTD_COUNTRY
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_COUNTRY.OnClick item, ParentForm
     End With
   Case "ITTD_CAMERA"
     With pnlITTD_CAMERA
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_CAMERA.OnClick item, ParentForm
     End With
   Case "ITTD_RULE"
     With pnlITTD_RULE
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_RULE.OnClick item, ParentForm
     End With
   Case "ITTD_OPTTYPE"
     With pnlITTD_OPTTYPE
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_OPTTYPE.OnClick item, ParentForm
     End With
   Case "ITTD_MOROZ"
     With pnlITTD_MOROZ
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_MOROZ.OnClick item, ParentForm
     End With
   Case "ITTD_EMAIL"
     With pnlITTD_EMAIL
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     pnlITTD_EMAIL.OnClick item, ParentForm
     End With
     End Select
End Sub

'������������� ��������
'Parameters:
'[IN][OUT]   ObjItem , ��� ���������: object,
'[IN][OUT]   MyHost , ��� ���������: GUI,
'[IN][OUT]  aModalMode , ��� ���������: boolean ,
'[IN][OUT]   aParentForm , ��� ���������: object  - ...
'See Also:
'Example:
'  call me.Init({���������})
Public Sub Init(ObjItem As Object, MyHost As GUI, aModalMode As Boolean, aParentForm As Object)
  On Error Resume Next
 Set item = ObjItem
 Set Host = MyHost
 Set ParentForm = aParentForm
 ModalMode = aModalMode
  Dim ff As Long, buf As String

ts.Tabs.item(1).Caption = "��� ������"
ts.Tabs.item(1).Key = "ITTD_GTYPE"
pnlITTD_GTYPE.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_ZTYPE", "��� ����")
pnlITTD_ZTYPE.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_PLTYPE", "��� ������")
pnlITTD_PLTYPE.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_QTYPE", "��� ������")
pnlITTD_QTYPE.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_ATYPE", "��� ��������")
pnlITTD_ATYPE.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_SRV", "������")
pnlITTD_SRV.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_PART", "������ ������")
pnlITTD_PART.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_FACTORY", "�����")
pnlITTD_FACTORY.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_KILLPLACE", "�����")
pnlITTD_KILLPLACE.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_COUNTRY", "������")
pnlITTD_COUNTRY.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_CAMERA", "����������� ������")
pnlITTD_CAMERA.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_RULE", "������� ������������ ������")
pnlITTD_RULE.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_OPTTYPE", "��� �����������")
pnlITTD_OPTTYPE.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_MOROZ", "��������� ���������")
pnlITTD_MOROZ.OnInit item, ParentForm
Call ts.Tabs.Add(, "ITTD_EMAIL", "�������� ������")
pnlITTD_EMAIL.OnInit item, ParentForm
  Set TSCustom = New MTZ_CUSTOMTAB.TabStripCustomizer
  TSCustom.Init ts, "ITTD", "ctlmain_main"
  TSCustom.SetupFromRegistry ts
 ts_click
End Sub
Private Sub UserControl_Terminate()
  On Error Resume Next
  Set item = Nothing
  Set Host = Nothing
  Set ParentForm = Nothing
  Set TSCustom = Nothing
 pnlITTD_GTYPE.CloseClass
 pnlITTD_ZTYPE.CloseClass
 pnlITTD_PLTYPE.CloseClass
 pnlITTD_QTYPE.CloseClass
 pnlITTD_ATYPE.CloseClass
 pnlITTD_SRV.CloseClass
 pnlITTD_PART.CloseClass
 pnlITTD_FACTORY.CloseClass
 pnlITTD_KILLPLACE.CloseClass
 pnlITTD_COUNTRY.CloseClass
 pnlITTD_CAMERA.CloseClass
 pnlITTD_RULE.CloseClass
 pnlITTD_OPTTYPE.CloseClass
 pnlITTD_MOROZ.CloseClass
 pnlITTD_EMAIL.CloseClass
End Sub

'�������� ��� ����������
'Parameters:
' ���������� ���
'See Also:
'Example:
'  call me.OnSave()
Public Sub OnSave()
 pnlITTD_GTYPE.OnSave
 pnlITTD_ZTYPE.OnSave
 pnlITTD_PLTYPE.OnSave
 pnlITTD_QTYPE.OnSave
 pnlITTD_ATYPE.OnSave
 pnlITTD_SRV.OnSave
 pnlITTD_PART.OnSave
 pnlITTD_FACTORY.OnSave
 pnlITTD_KILLPLACE.OnSave
 pnlITTD_COUNTRY.OnSave
 pnlITTD_CAMERA.OnSave
 pnlITTD_RULE.OnSave
 pnlITTD_OPTTYPE.OnSave
 pnlITTD_MOROZ.OnSave
 pnlITTD_EMAIL.OnSave
End Sub

'��� �� ������� ����� ��������
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
m_IsChanged = m_IsChanged Or pnlITTD_GTYPE.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_ZTYPE.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_PLTYPE.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_QTYPE.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_ATYPE.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_SRV.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_PART.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_FACTORY.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_KILLPLACE.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_COUNTRY.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_CAMERA.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_RULE.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_OPTTYPE.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_MOROZ.IsChanged
m_IsChanged = m_IsChanged Or pnlITTD_EMAIL.IsChanged
  IsChanged = m_IsChanged
End Function
Private Sub UserControl_Resize()
 On Error Resume Next
ts.Top = 0
ts.Left = 0
ts.Width = UserControl.Width
ts.Height = UserControl.Height
ts_click
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



'���������� ������� ���� ��������� (����� � ����������)
'Parameters:
' ���������� ���
'Returns:
'  �������� ���� Integer
'See Also:
'Example:
' dim variable as Integer
'  variable = me.StatusMenuCount()
Public Function StatusMenuCount() As Integer
  StatusMenuCount = 0
End Function

'��������� ���� ���������
'Parameters:
' ���������� ���
'Returns:
'  ������ ������ Object)
'  ,��� Nothing
'See Also:
'Example:
' dim variable as Object)
' Set variable = me.SetupStatusMenu()
Public Function SetupStatusMenu(m() As Object)
    Dim i As Long
    On Error Resume Next
    i = 0
End Function

'�������� ������� ��������� � ��������� ���������� ���������
'Parameters:
' ���������� ���
'See Also:
'Example:
'  call me.CheckStatusMenu()
Public Sub CheckStatusMenu(m() As Object)
    Dim i As Long
    On Error Resume Next
    For i = 1 To StatusMenuCount
        m(i).Checked = False
        m(i).Enabled = False
        If Not item Is Nothing Then
          If m(i).Tag = item.Statusid Then
            m(i).Checked = True
          End If
        End If
    Next
    If Not item Is Nothing Then
   End If
End Sub



