VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSaveDesc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Сохранить описание типа"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6855
   Icon            =   "frmSaveDesc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "Выделить все"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdUnselAll 
      Caption         =   "Отменить все"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   315
      Left            =   6405
      TabIndex        =   3
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   4320
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Сохранить"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.ListBox cmbType 
      Height          =   2535
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   840
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Описание типа"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   6615
   End
   Begin VB.Label Label8 
      Caption         =   "Куда сохранить:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1785
   End
End
Attribute VB_Name = "frmSaveDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TypeName As String
Public site As String

Public OK As Boolean
Private ID  As String
Public NewObject  As Object
Private types As Collection



Private Sub CancelButton_Click()
OK = False
Set types = Nothing
Me.Hide
End Sub




Private Sub cmbType_DblClick()
OKButton_Click
End Sub

Private Sub cmdSelAll_Click()
  Dim i As Long
  For i = 0 To cmbType.ListCount - 1
  cmbType.Selected(i) = True
  Next
End Sub

Private Sub cmdUnselAll_Click()
  Dim i As Long
  For i = 0 To cmbType.ListCount - 1
  cmbType.Selected(i) = False
  Next
End Sub

Private Sub Form_Load()
  Dim rs As ADODB.Recordset, i
  Dim n As String, tn As String
  Set rs = Session.GetRowsEx("OBJECTTYPE", , , , " order by the_comment")
  Dim o As tmpInst
  Set types = New Collection
  i = 0
  While Not rs.EOF
      i = i + 1
      Set o = New tmpInst
      o.Name = rs!the_comment & "(" & rs!Name & ")"
      o.ObjType = rs!Name
      o.IsSingle = rs!IsSingleInstance
      o.ID = rs!objecttypeid
      types.Add o
      cmbType.AddItem o.Name
      cmbType.ItemData(cmbType.NewIndex) = i
      rs.MoveNext
  Wend
  Set rs = Nothing
  
  If cmbType.ListCount > 0 Then
    cmbType.ListIndex = 0
  End If
End Sub

Private Sub OKButton_Click()
  On Error GoTo bye
  'If cmbType.ListIndex = -1 Then Exit Sub
  
  
'  TypeName = types.item(cmbType.ItemData(cmbType.ListIndex)).ObjType
'  ID = types.item(cmbType.ItemData(cmbType.ListIndex)).ID
'  OK = True
'  Set types = Nothing
'  Me.Hide
Dim i As Long
pb.Max = cmbType.ListCount - 1
pb.Min = 0
pb.Value = 0
pb.Visible = True
For i = 0 To cmbType.ListCount - 1
  If cmbType.Selected(i) Then
    SaveTypeXML types.item(cmbType.ItemData(i)).ID
    cmbType.Selected(i) = False
  End If
  pb.Value = i
Next
pb.Visible = False


bye:
End Sub

Private Sub SaveTypeXML(ByVal ID As String)
On Error Resume Next
 Dim item As OBJECTTYPE
 Set item = model.OBJECTTYPE.item(ID)
 If item Is Nothing Then Exit Sub
 
 If item.Application.MTZSession.CheckRight(item.SecureStyleID, "XMLSAVE") Then
 
  On Error GoTo bye
  Dim fn As String
 
   fn = txtPath & item.Name & ".xml"
   item.LockResource True
   
   
   Dim xdom As MSXML2.DOMDocument
   Set xdom = New MSXML2.DOMDocument
   xdom.loadXML "<OBJECTTYPE></OBJECTTYPE>"
   item.XMLSave xdom.lastChild, xdom
   xdom.Save fn
   item.UnLockResource
 End If
bye:
End Sub


Private Sub cmdPath_Click()
  Dim path As String
  path = GetPath("Каталог для сохранения документов")
  
  If (path <> vbNullString) Then
    txtPath.Text = path
  End If
End Sub

Private Function GetPath(Caption As String) As String
    Dim bi As browseinfo
    Dim lngPath As Long
    Dim lngBrowse As Long
    Dim path As String
    Dim inull As Integer
    
    GetPath = path
    
    Call SHGetSpecialFolderLocation(Me.hwnd, 17, lngPath)

    bi.hwndOwner = Me.hwnd
    bi.lpszTitle = Caption
    bi.pszDisplayName = String(MAX_PATH, 0)
    bi.pidlRoot = lngPath
    bi.lpfn = 0
    bi.ulFlags = 1
    bi.lParam = 0
    
    lngBrowse = SHBrowseForFolder(bi)
    
    path = String(MAX_PATH, 0)
    
    Call SHGetPathFromIDList(lngBrowse, path)
    
    inull = InStr(path, vbNullChar)
    
    If inull Then
      path = Left(path, inull - 1)
    End If
    
    If path <> vbNullString Then
      If Right(path, 1) <> "\" Then
        path = path + "\"
      End If
    End If
    
    GetPath = path
End Function


