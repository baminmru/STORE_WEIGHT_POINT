VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1801C003-859D-471D-BF31-D4428050324B}#2.1#0"; "MTZ_PANEL.ocx"
Begin VB.UserControl ITTOPT_DEFadmi 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MTZ_PANEL.ScrolledWindow Panel 
      Height          =   1000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1000
      _ExtentX        =   1773
      _ExtentY        =   1773
      Begin VB.TextBox txtDocNum 
         Height          =   300
         Left            =   3450
         MaxLength       =   15
         TabIndex        =   34
         ToolTipText     =   "�����"
         Top             =   4635
         Width           =   1800
      End
      Begin VB.TextBox txtPalType 
         Height          =   300
         Left            =   3450
         MaxLength       =   1
         TabIndex        =   32
         ToolTipText     =   "��� �������"
         Top             =   3930
         Width           =   3000
      End
      Begin VB.TextBox txtIsBrak 
         Height          =   300
         Left            =   3450
         MaxLength       =   20
         TabIndex        =   30
         ToolTipText     =   "����"
         Top             =   3225
         Width           =   3000
      End
      Begin VB.TextBox txtVetSved 
         Height          =   300
         Left            =   3450
         MaxLength       =   30
         TabIndex        =   28
         ToolTipText     =   "����������"
         Top             =   2520
         Width           =   3000
      End
      Begin VB.TextBox txtPartia 
         Height          =   300
         Left            =   3450
         MaxLength       =   255
         TabIndex        =   26
         ToolTipText     =   "������"
         Top             =   1815
         Width           =   3000
      End
      Begin VB.TextBox txtKILL_NUMBER 
         Height          =   300
         Left            =   3450
         MaxLength       =   255
         TabIndex        =   24
         ToolTipText     =   "� �����"
         Top             =   1110
         Width           =   3000
      End
      Begin VB.TextBox txtFactory 
         Height          =   300
         Left            =   3450
         MaxLength       =   255
         TabIndex        =   22
         ToolTipText     =   "�����"
         Top             =   405
         Width           =   3000
      End
      Begin VB.TextBox txtmade_country 
         Height          =   300
         Left            =   300
         MaxLength       =   255
         TabIndex        =   20
         ToolTipText     =   "������ �������������"
         Top             =   6045
         Width           =   3000
      End
      Begin VB.TextBox txtTheClient 
         Height          =   300
         Left            =   300
         MaxLength       =   255
         TabIndex        =   18
         ToolTipText     =   "������"
         Top             =   5340
         Width           =   3000
      End
      Begin VB.TextBox txtgood 
         Height          =   300
         Left            =   300
         MaxLength       =   255
         TabIndex        =   16
         ToolTipText     =   "�����"
         Top             =   4635
         Width           =   3000
      End
      Begin VB.TextBox txtTheKamera 
         Height          =   300
         Left            =   300
         MaxLength       =   15
         TabIndex        =   14
         ToolTipText     =   "������"
         Top             =   3930
         Width           =   1800
      End
      Begin MSComCtl2.DTPicker dtpDateToOptimize 
         Height          =   300
         Left            =   300
         TabIndex        =   12
         ToolTipText     =   "�������� ���� �����������"
         Top             =   3225
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd.MM.yyyy"
         Format          =   52035587
         CurrentDate     =   39861
      End
      Begin MTZ_PANEL.DropButton cmdTheRule 
         Height          =   300
         Left            =   2850
         TabIndex        =   10
         Tag             =   "refopen.ico"
         ToolTipText     =   "������� ������������ ������"
         Top             =   2520
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtTheRule 
         Height          =   300
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "������� ������������ ������"
         Top             =   2520
         Width           =   2550
      End
      Begin MSComCtl2.DTPicker dtpOPtDate 
         Height          =   300
         Left            =   300
         TabIndex        =   7
         ToolTipText     =   "���� �������� ������"
         Top             =   1815
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd.MM.yyyy"
         Format          =   52035587
         CurrentDate     =   39861
      End
      Begin MTZ_PANEL.DropButton cmdOptType 
         Height          =   300
         Left            =   2850
         TabIndex        =   5
         Tag             =   "refopen.ico"
         ToolTipText     =   "��� �����������"
         Top             =   1110
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtOptType 
         Height          =   300
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "��� �����������"
         Top             =   1110
         Width           =   2550
      End
      Begin VB.TextBox txtFormattedDocNum 
         Height          =   300
         Left            =   300
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "����� ���������"
         Top             =   405
         Width           =   3000
      End
      Begin VB.Label lblDocNum 
         BackStyle       =   0  'Transparent
         Caption         =   "�����:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3450
         TabIndex        =   33
         Top             =   4305
         Width           =   3000
      End
      Begin VB.Label lblPalType 
         BackStyle       =   0  'Transparent
         Caption         =   "��� �������:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3450
         TabIndex        =   31
         Top             =   3600
         Width           =   3000
      End
      Begin VB.Label lblIsBrak 
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3450
         TabIndex        =   29
         Top             =   2895
         Width           =   3000
      End
      Begin VB.Label lblVetSved 
         BackStyle       =   0  'Transparent
         Caption         =   "����������:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3450
         TabIndex        =   27
         Top             =   2190
         Width           =   3000
      End
      Begin VB.Label lblPartia 
         BackStyle       =   0  'Transparent
         Caption         =   "������:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3450
         TabIndex        =   25
         Top             =   1485
         Width           =   3000
      End
      Begin VB.Label lblKILL_NUMBER 
         BackStyle       =   0  'Transparent
         Caption         =   "� �����:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3450
         TabIndex        =   23
         Top             =   780
         Width           =   3000
      End
      Begin VB.Label lblFactory 
         BackStyle       =   0  'Transparent
         Caption         =   "�����:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3450
         TabIndex        =   21
         Top             =   75
         Width           =   3000
      End
      Begin VB.Label lblmade_country 
         BackStyle       =   0  'Transparent
         Caption         =   "������ �������������:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   300
         TabIndex        =   19
         Top             =   5715
         Width           =   3000
      End
      Begin VB.Label lblTheClient 
         BackStyle       =   0  'Transparent
         Caption         =   "������:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   300
         TabIndex        =   17
         Top             =   5010
         Width           =   3000
      End
      Begin VB.Label lblgood 
         BackStyle       =   0  'Transparent
         Caption         =   "�����:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   300
         TabIndex        =   15
         Top             =   4305
         Width           =   3000
      End
      Begin VB.Label lblTheKamera 
         BackStyle       =   0  'Transparent
         Caption         =   "������:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   300
         TabIndex        =   13
         Top             =   3600
         Width           =   3000
      End
      Begin VB.Label lblDateToOptimize 
         BackStyle       =   0  'Transparent
         Caption         =   "�������� ���� �����������:"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   300
         TabIndex        =   11
         Top             =   2895
         Width           =   3000
      End
      Begin VB.Label lblTheRule 
         BackStyle       =   0  'Transparent
         Caption         =   "������� ������������ ������:"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   300
         TabIndex        =   8
         Top             =   2190
         Width           =   3000
      End
      Begin VB.Label lblOPtDate 
         BackStyle       =   0  'Transparent
         Caption         =   "���� �������� ������:"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   300
         TabIndex        =   6
         Top             =   1485
         Width           =   3000
      End
      Begin VB.Label lblOptType 
         BackStyle       =   0  'Transparent
         Caption         =   "��� �����������:"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   300
         TabIndex        =   3
         Top             =   780
         Width           =   3000
      End
      Begin VB.Label lblFormattedDocNum 
         BackStyle       =   0  'Transparent
         Caption         =   "����� ���������:"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   300
         TabIndex        =   1
         Top             =   75
         Width           =   3000
      End
   End
End
Attribute VB_Name = "ITTOPT_DEFadmi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'������ �������������� ������� �������� ������� �� �����������
   Public item As Object
   Private OnInit As Boolean
   Public Event Changed()
   Private mIsChanged As Boolean






'������� ������ ������
'Parameters:
'[IN][OUT]  Runner , ��� ���������: mtzmanager.main,
'[IN]   TypeName , ��� ���������: String,
'[IN][OUT]   ID , ��� ���������: string,
'[IN][OUT]   Brief , ��� ���������: string,
'[IN][OUT]   Cancel , ��� ���������: boolean  - ...
'Returns:
' Boolean, ��������� ����������:
'   true  -
'   false -
'See Also:
'Example:
' dim variable as Boolean
' variable = me.findObject(...���������...)
Private Function findObject(Runner As MTZManager.Main, ByVal TypeName As String, ByRef ID As String, ByRef brief As String, ByRef Cancel As Boolean) As Boolean
Dim objFinder As Object
Dim result As Boolean
result = False
On Error GoTo bye
Set objFinder = CreateObject(TypeName & "_FIND.Find")
result = objFinder.Run(Runner, ID, brief, Cancel)
bye:
findObject = result
End Function

'������� ��������� ������
'Parameters:
' ���������� ���
'Returns:
' Boolean, ��������� ����������:
'   true  -
'   false -
'See Also:
'Example:
' dim variable as Boolean
'  variable = me.IsChanged()
Public Function IsChanged() As Boolean
  IsChanged = mIsChanged
End Function
Private Sub UserControl_Resize()
  On Error Resume Next
  Panel.Width = UserControl.Width
  Panel.Height = UserControl.Height
End Sub

Private Sub txtFormattedDocNum_Change()
  Changing

End Sub
Private Sub txtOptType_Change()
  If Not (OnInit) Then
  Changing

  End If
End Sub
Private Sub cmdOptType_CLick()
  On Error Resume Next
        Dim ID As String, brief As String
        If item.Application.Manager.GetReferenceDialogEx2("ITTD_OPTTYPE", ID, brief) Then
          txtOptType.Tag = Left(ID, 38)
          txtOptType = brief
        End If
End Sub
Private Sub cmdOptType_MenuClick(ByVal sCaption As String)
          txtOptType.Tag = ""
          txtOptType = ""
End Sub
Private Sub dtpOPtDate_Change()
  Changing

End Sub
Private Sub txtTheRule_Change()
  If Not (OnInit) Then
  Changing

  End If
End Sub
Private Sub cmdTheRule_CLick()
  On Error Resume Next
        Dim ID As String, brief As String
        If item.Application.Manager.GetReferenceDialogEx2("ITTD_RULE", ID, brief) Then
          txtTheRule.Tag = Left(ID, 38)
          txtTheRule = brief
        End If
End Sub
Private Sub cmdTheRule_MenuClick(ByVal sCaption As String)
          txtTheRule.Tag = ""
          txtTheRule = ""
End Sub
Private Sub dtpDateToOptimize_Change()
  Changing

End Sub
Private Sub txtTheKamera_Validate(Cancel As Boolean)
If txtTheKamera.Text <> "" Then
 On Error Resume Next
  If Not IsNumeric(txtTheKamera.Text) Then
     Cancel = True
     MsgBox "��������� �����", vbOKOnly + vbExclamation, "��������"
     txtTheKamera.SetFocus
  ElseIf val(txtTheKamera.Text) <> CLng(val(txtTheKamera.Text)) Then
     Cancel = True
     MsgBox "��������� ����� �����", vbOKOnly + vbExclamation, "��������"
     txtTheKamera.SetFocus
  End If
End If
End Sub
Private Sub txtTheKamera_KeyPess(KeyAscii As Integer)
Dim s As String
s = "0123456789.,-" & Chr(8)
If InStr(s, Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = 0

End Sub
Private Sub txtTheKamera_Change()
  Changing

End Sub
Private Sub txtgood_Change()
  Changing

End Sub
Private Sub txtTheClient_Change()
  Changing

End Sub
Private Sub txtmade_country_Change()
  Changing

End Sub
Private Sub txtFactory_Change()
  Changing

End Sub
Private Sub txtKILL_NUMBER_Change()
  Changing

End Sub
Private Sub txtPartia_Change()
  Changing

End Sub
Private Sub txtVetSved_Change()
  Changing

End Sub
Private Sub txtIsBrak_Change()
  Changing

End Sub
Private Sub txtPalType_Change()
  Changing

End Sub
Private Sub txtDocNum_Validate(Cancel As Boolean)
If txtDocNum.Text <> "" Then
 On Error Resume Next
  If Not IsNumeric(txtDocNum.Text) Then
     Cancel = True
     MsgBox "��������� �����", vbOKOnly + vbExclamation, "��������"
     txtDocNum.SetFocus
  ElseIf val(txtDocNum.Text) <> CLng(val(txtDocNum.Text)) Then
     Cancel = True
     MsgBox "��������� ����� �����", vbOKOnly + vbExclamation, "��������"
     txtDocNum.SetFocus
  End If
End If
End Sub
Private Sub txtDocNum_KeyPess(KeyAscii As Integer)
Dim s As String
s = "0123456789.,-" & Chr(8)
If InStr(s, Chr(KeyAscii)) > 0 Then Exit Sub
KeyAscii = 0

End Sub
Private Sub txtDocNum_Change()
  Changing

End Sub
Private Sub UserControl_Terminate()
  Set item = Nothing
End Sub

'�������� ������������ ���������� ������ �� ������ ��������������
'Parameters:
' ���������� ���
'Returns:
' Boolean, ��������� ����������:
'   true  -
'   false -
'See Also:
'Example:
' dim variable as boolean
'  variable = me.IsOK()
Public Function IsOK() As Boolean
  On Error Resume Next
  Dim mIsOK As Boolean
  mIsOK = True

'If mIsOK Then mIsOK = IsSet(txtFormattedDocNum.Text)
If mIsOK Then mIsOK = txtOptType.Tag <> ""
If mIsOK Then mIsOK = IsSet(dtpOPtDate.Value)
If mIsOK Then mIsOK = txtTheRule.Tag <> ""
If mIsOK Then mIsOK = IsSet(dtpDateToOptimize.Value)
  IsOK = mIsOK
End Function
Private Function AddSQLRefIds(ByVal strTo As String, ByVal fldName As String, ByVal strFrom As String) As String
  Dim XMLDocFrom As New DOMDocument
  Dim XMLDocTo As New DOMDocument
  AddSQLRefIds = strTo
  On Error GoTo err
  Call XMLDocTo.loadXML(strTo)
  Call XMLDocFrom.loadXML(strFrom)
  Dim Node As MSXML2.IXMLDOMNode
  Dim ID As String
  For Each Node In XMLDocFrom.childNodes.item(0).childNodes
    If (Node.baseName = "ID") Then
      ID = Node.Text
      Dim NodeTO As MSXML2.IXMLDOMNode
      Dim bAdded As Boolean
      bAdded = False
      For Each NodeTO In XMLDocTo.childNodes.item(0).childNodes
       If (NodeTO.baseName = fldName & "ID") Then
         NodeTO.Text = ID
         bAdded = True
         Exit For
       End If
      Next
      If (Not bAdded) Then
       Dim newNode As MSXML2.IXMLDOMNode
       Set newNode = XMLDocTo.createNode(MSXML2.NODE_ELEMENT, fldName & "ID", XMLDocTo.namespaceURI)
        newNode.Text = ID
       Call XMLDocTo.childNodes.item(0).appendChild(newNode)
      End If
      AddSQLRefIds = XMLDocTo.xml
      Exit For
    End If
  Next
err:
End Function

'������������� ��������� ������
'Parameters:
' ���������� ���
'See Also:
'Example:
'  call me.InitPanel()
Public Sub InitPanel()
OnInit = True
Dim iii As Long ' for combo only
If item.CanChange Then
  Panel.Enabled = True
Else
  Panel.Enabled = False
End If

  On Error Resume Next
txtFormattedDocNum = item.FormattedDocNum
If Not item.OptType Is Nothing Then
  txtOptType.Tag = item.OptType.ID
  txtOptType = item.OptType.brief
Else
  txtOptType.Tag = ""
  txtOptType = ""
End If
 LoadBtnPictures cmdOptType, cmdOptType.Tag
  cmdOptType.RemoveAllMenu
  cmdOptType.AddMenu "��������"
dtpOPtDate = Date
If item.OPtDate <> 0 Then
 dtpOPtDate = item.OPtDate
End If
If Not item.TheRule Is Nothing Then
  txtTheRule.Tag = item.TheRule.ID
  txtTheRule = item.TheRule.brief
Else
  txtTheRule.Tag = ""
  txtTheRule = ""
End If
 LoadBtnPictures cmdTheRule, cmdTheRule.Tag
  cmdTheRule.RemoveAllMenu
  cmdTheRule.AddMenu "��������"
dtpDateToOptimize = Date
If item.DateToOptimize <> 0 Then
 dtpDateToOptimize = item.DateToOptimize
End If
txtTheKamera = item.TheKamera
  On Error Resume Next
txtgood = item.good
  On Error Resume Next
txtTheClient = item.TheClient
  On Error Resume Next
txtmade_country = item.made_country
  On Error Resume Next
txtFactory = item.Factory
  On Error Resume Next
txtKILL_NUMBER = item.KILL_NUMBER
  On Error Resume Next
txtPartia = item.Partia
  On Error Resume Next
txtVetSved = item.VetSved
  On Error Resume Next
txtIsBrak = item.IsBrak
  On Error Resume Next
txtPalType = item.PalType
txtDocNum = item.DocNum
' ������� �������� ID �� ���� SQLReference
OnInit = False
End Sub
Private Sub Changing()
If OnInit Then Exit Sub

 mIsChanged = True
 RaiseEvent Changed
End Sub

'����������
'Parameters:
' ���������� ���
'See Also:
'Example:
'  call me.Save({���������})
Public Sub Save()
If OnInit Then Exit Sub

item.FormattedDocNum = txtFormattedDocNum
If txtOptType.Tag <> "" Then
  Set item.OptType = item.Application.FindRowObject("ITTD_OPTTYPE", txtOptType.Tag)
Else
  Set item.OptType = Nothing
End If
  If IsNull(dtpOPtDate) Then
    item.OPtDate = 0
  Else
    item.OPtDate = dtpOPtDate.Value
  End If
If txtTheRule.Tag <> "" Then
  Set item.TheRule = item.Application.FindRowObject("ITTD_RULE", txtTheRule.Tag)
Else
  Set item.TheRule = Nothing
End If
  If IsNull(dtpDateToOptimize) Then
    item.DateToOptimize = 0
  Else
    item.DateToOptimize = dtpDateToOptimize.Value
  End If
item.TheKamera = CDbl(txtTheKamera)
item.good = txtgood
item.TheClient = txtTheClient
item.made_country = txtmade_country
item.Factory = txtFactory
item.KILL_NUMBER = txtKILL_NUMBER
item.Partia = txtPartia
item.VetSved = txtVetSved
item.IsBrak = txtIsBrak
item.PalType = txtPalType
item.DocNum = CDbl(txtDocNum)
 mIsChanged = False
 RaiseEvent Changed
End Sub

'������ ������� ������ �� ������� ���������
'Parameters:
'[IN][OUT]  x , ��� ���������: Single,
'[IN][OUT]   y , ��� ���������: single  - ...
'See Also:
'Example:
'  call me.OptimalSize({���������})
 Public Sub OptimalSize(X As Single, Y As Single)
   Panel.OptimalSize X, Y
   X = X + Panel.Left
   Y = Y + Panel.Top
 End Sub
 
 
 Public Function OptimalY() As Single
   Dim X As Single, Y As Single
   Panel.OptimalSize X, Y
   OptimalY = Y
 End Function

'�������� ���� ��������� ������
'Parameters:
' ���������� ���
'See Also:
'Example:
'  call me.Customize()
 Public Sub Customize()
   Panel.Customize
 End Sub

'������ ��������� ��������� �� ������ ��� ���������� �������
'Parameters:
' ���������� ���
'Returns:
'  �������� ���� string
'See Also:
'Example:
' dim variable as string
' variable = me. PanelCustomisationString
 Public Property Get PanelCustomisationString() As String
   PanelCustomisationString = Panel.PanelCustomisationString
 End Property

'��������������  ������� ��������� �� ������
'Parameters:
'[IN][OUT]  s , ��� ���������: string  - ...
'See Also:
'Example:
' dim value as Variant
' value = ...��������...
' me. PanelCustomisationString = value
 Public Property Let PanelCustomisationString(s As String)
   Panel.PanelCustomisationString = s
 End Property

'���������� ��������� ���������
'Parameters:
' ���������� ���
'Returns:
' Boolean, ��������� ����������:
'   true  -
'   false -
'See Also:
'Example:
' dim variable as boolean
' variable = me. Enabled
 Public Property Get Enabled() As Boolean
   Enabled = Panel.Enabled
 End Property

'������ \ ���������� ��������� ���������
'Parameters:
'[IN]   v , ��� ���������: boolean  - ...
'See Also:
'Example:
' dim value as Variant
' value = ...��������...
' me. Enabled = value
 Public Property Let Enabled(ByVal v As Boolean)
   Panel.Enabled = v
 End Property



