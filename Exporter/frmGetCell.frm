VERSION 5.00
Begin VB.Form frmGetCell 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Выбор  ячейки"
   ClientHeight    =   4320
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmGetCell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtmask 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Зона хранения"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Буферная зона"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.ListBox lstCells 
      Height          =   3180
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Шаблон:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmGetCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim col As Collection
Dim db As DBuffer

Public itemid As String
Public country As String
Public factory As String
Public killplace As String

Public OK As Boolean
Public OUtID As String
Public OutCode As String
Private conn As ADODB.Connection
Private cells As Long


Private Sub CancelButton_Click()
  OK = False
  Me.Hide
End Sub

Private Sub FillList(ByVal buffer As Boolean)

 ' запрашиваем свободное место в буферной зоне
  Dim bzrs As ADODB.Recordset
  Dim loccode As ADODB.Recordset
  Dim bzid As String
  Set conn = Manager.GetCustomObjects("refref")
  If conn.State <> adStateOpen Then
    conn.Open
  End If
  
  lstCells.Clear
  
  Set col = New Collection
  Dim qry As String
  Dim subquery As String
  subquery = ""
  
  subquery = Manager.GetCustomObjects("camFilter").Name
  If subquery <> "" Then
    subquery = subquery
  End If
  
  
  If country <> "" Then
    subquery = subquery & " and stock.custom_field6='" & country & "' "
  End If
  
  If factory <> "" Then
    subquery = subquery & " and stock.custom_field4='" & factory & "' "
  End If
   
  If killplace <> "" Then
    subquery = subquery & " and stock.custom_field11='" & killplace & "' "
  End If
   
   If buffer Then
     qry = "select id, code from location where id in (" & _
     " select  distinct location_id id from stock join location on location.id = stock.location_id " & _
     " Where location.description like '%;B%' and location.active='Y' " & _
     " group by location_id,location.description " & _
     " having count(*) < loc_qty_pal " & _
     " and id in (" & _
     " select  distinct location_id id from stock join location on location.id = stock.location_id " & _
     " Where location.description like '%;B%' and location.active='Y' and stock.item_iD = " & itemid & subquery & ")"
   Else
     qry = "select id, code from location where id in (" & _
     " select  distinct location_id id from stock join location on location.id = stock.location_id " & _
     " Where location.description not like '%;B%' and location.active='Y' " & _
     " group by location_id,location.description " & _
     " having count(*) < loc_qy_pal " & _
     " and id in (" & _
     " select  distinct location_id id from stock join location on location.id = stock.location_id " & _
     " Where location.description not like '%;B%' and location.active='Y' and stock.item_iD = " & itemid & subquery & " )"
   End If
   
    
  Set bzrs = conn.Execute(qry)
  
  
  While Not bzrs.EOF
   Set db = New DBuffer
   db.id = bzrs!id
   db.Name = bzrs!code
   
   col.Add db
   lstCells.AddItem "+ " & db.Name
   lstCells.ItemData(lstCells.NewIndex) = col.Count
   
   bzrs.MoveNext
  Wend
  
  bzrs.Close
  Dim qq As String
  If txtmask <> "" Then
    qq = " and code like '%" & txtmask & "%'"
  Else
    qq = ""
  End If
  If Manager.GetCustomObjects("camFilter").Name <> "" Then
    qq = qq & Manager.GetCustomObjects("camFilter").Name
  End If
  
  If buffer Then
    Set bzrs = conn.Execute( _
    "select top " & cells & "  id, code from location where description like '%;B%' " & qq & " and  id not in ( " & _
    " select location_id from stock where location_id is not null )")
  Else
    Set bzrs = conn.Execute( _
    "select top " & cells & "  id, code from location where description not like '%;B%' " & qq & " and  id not in ( " & _
    " select location_id from stock where location_id is not null )")
  End If
  
  While Not bzrs.EOF
   Set db = New DBuffer
   db.id = bzrs!id
   db.Name = bzrs!code
   
   col.Add db
   lstCells.AddItem db.Name
   lstCells.ItemData(lstCells.NewIndex) = col.Count
   bzrs.MoveNext
  Wend
  bzrs.Close
  Me.Caption = "Выбор ячейки (" & col.Count & ")"
  If lstCells.ListCount > 0 Then
  lstCells.ListIndex = 0
  End If
End Sub


Private Sub Form_Load()
 cells = GetSetting("RBH", "ITTSETTINGS", "SHOWCELLS", 100)
 FillList True
End Sub

Private Sub lstCells_DblClick()
OKButton_Click
End Sub

Private Sub OKButton_Click()
  Dim i As Long
  If lstCells.ListIndex = -1 Then Exit Sub
  i = lstCells.ItemData(lstCells.ListIndex)
  Set db = col.Item(i)
  OUtID = db.id
  OutCode = db.Name
  OK = True
  Me.Hide
  
End Sub

Private Sub Option1_Click()
  If Option1.Value Then
    FillList True
  Else
    FillList False
  End If

End Sub

Private Sub Option2_Click()
  If Option2.Value Then
    FillList False
  Else
    FillList True
  End If
End Sub

Private Sub txtmask_Change()
  If Option1.Value Then
    FillList True
  Else
    FillList False
  End If
End Sub
