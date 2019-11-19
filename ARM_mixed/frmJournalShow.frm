VERSION 5.00
Object = "{BB95CD0C-5138-4A76-AF3C-30EFB10D1594}#25.0#0"; "MTZJournal.ocx"
Begin VB.Form frmJournalShow 
   ClientHeight    =   4995
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   7905
   Icon            =   "frmJournalShow.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin MTZJournal.JournalView jv 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   8705
   End
End
Attribute VB_Name = "frmJournalShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 455
Option Explicit

'����� ������� �� ���� ����������� �������� ������� ��� ��������� ( �� ������������)


'������� ��� ������ ���������� ������� ������� � ���������� ������������
' ��� ����������� � ������� ���� ����������
Public Event OnAdd(usedefaut As Boolean, Refesh As Boolean)
Attribute OnAdd.VB_HelpID = 460
Public Event OnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
Attribute OnRun.VB_HelpID = 495
Public Event OnEdit(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
Attribute OnEdit.VB_HelpID = 475
Public Event OnFilter(usedefaut As Boolean)
Attribute OnFilter.VB_HelpID = 480
Public Event OnClearFilter()
Attribute OnClearFilter.VB_HelpID = 465
Public Event OnPrint(usedefaut As Boolean)
Attribute OnPrint.VB_HelpID = 490
Public Event OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean)
Attribute OnInit.VB_HelpID = 485
Public Event OnDel(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
Attribute OnDel.VB_HelpID = 470


Private Sub Form_Load()
Dim bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean
bAdd = True
bEdit = True
bRun = True
bDel = True
bFilter = True
RaiseEvent OnInit(bAdd, bEdit, bRun, bDel, bFilter)
jv.AllowAdd = bAdd
jv.AllowEdit = bEdit
jv.AllowRun = bRun
jv.AllowDel = bDel
jv.AllowFilter = bFilter
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormMDIForm Or UnloadMode = 5 Or UnloadMode = vbFormCode Or UnloadMode = vbAppWindows Or UnloadMode = vbAppTaskManager Then
    Cancel = False
  Else
    Cancel = True
    Me.Hide
  End If

End Sub

Private Sub Form_Resize()
  On Error Resume Next
  jv.Top = 0
  jv.Left = 0
  jv.Width = Me.ScaleWidth
  jv.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  jv.Visible = False
  Set jv.journal = Nothing
End Sub


Private Sub jv_JVDblClick(ByVal RowIndex As Long, UseDefault As Boolean)
  Dim r As Boolean
  r = False
  RaiseEvent OnRun(RowIndex, UseDefault, r)
End Sub

Private Sub jv_JVGetDocMode(ByVal Doc As Object, mode As String, IsDenied As Boolean)
  IsDenied = IsDocDenied(Doc)
  mode = GetDocumentMode(Doc)
End Sub

Private Sub jv_JVIsDocDeletable(ByVal Doc As Object, IsDeletable As Boolean)
  IsDeletable = RoleDocAllowDelete(Doc)
End Sub

Private Sub jv_JVOnAdd(usedefaut As Boolean, Refesh As Boolean)
  RaiseEvent OnAdd(usedefaut, Refesh)
End Sub

Private Sub jv_JVOnClearFilter()
  RaiseEvent OnClearFilter
End Sub

Private Sub jv_JVOnDel(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
  RaiseEvent OnDel(RowIndex, usedefaut, Refesh)
End Sub

Private Sub jv_JVOnEdit(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
  RaiseEvent OnEdit(RowIndex, usedefaut, Refesh)
End Sub

Private Sub jv_JVOnFilter(usedefaut As Boolean)
  RaiseEvent OnFilter(usedefaut)
End Sub

Private Sub jv_JVOnPrint(usedefaut As Boolean)
  RaiseEvent OnPrint(usedefaut)
End Sub

Private Sub jv_JVOnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
  RaiseEvent OnRun(RowIndex, usedefaut, Refesh)
End Sub

