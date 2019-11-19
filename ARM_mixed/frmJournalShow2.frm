VERSION 5.00
Object = "{81B9EB63-8321-4309-ABCB-72BFBEE99BC3}#6.3#0"; "MTZJournal2.ocx"
Begin VB.Form frmJournalShow2 
   ClientHeight    =   4995
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   7485
   Icon            =   "frmJournalShow2.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin MTZJournal2.JournalView2 jv 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7011
   End
End
Attribute VB_Name = "frmJournalShow2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 500
Option Explicit
'Форма отображения журнала

'События для прямой трансляции событий журнала в обработчик пользователя
' все обработчики в главном окне приложения
Public Event OnAdd(usedefaut As Boolean, Refesh As Boolean)
Attribute OnAdd.VB_HelpID = 505
Public Event OnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
Attribute OnRun.VB_HelpID = 540
Public Event OnEdit(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
Attribute OnEdit.VB_HelpID = 520
Public Event OnFilter(usedefaut As Boolean)
Attribute OnFilter.VB_HelpID = 525
Public Event OnPrint(usedefaut As Boolean)
Attribute OnPrint.VB_HelpID = 535
Public Event OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean)
Attribute OnInit.VB_HelpID = 530
Public Event OnDel(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
Attribute OnDel.VB_HelpID = 515
Public Event OnClearFilter()
Attribute OnClearFilter.VB_HelpID = 510

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
  If UnloadMode = vbFormMDIForm Or UnloadMode = vbFormCode Or UnloadMode = vbAppWindows Or UnloadMode = vbAppTaskManager Then
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
  RaiseEvent OnRun(RowIndex, UseDefault, r)
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

'получение режима для запуска документа
Private Sub jv_JVGetDocMode(ByVal Doc As Object, mode As String, IsDenied As Boolean)
  IsDenied = IsDocDenied(Doc)
  mode = GetDocumentMode(Doc)
End Sub

' проверка возможности удаления документа из журнала
Private Sub jv_JVIsDocDeletable(ByVal Doc As Object, IsDeletable As Boolean)
  IsDeletable = RoleDocAllowDelete(Doc)
End Sub
