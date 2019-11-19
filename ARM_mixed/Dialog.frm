VERSION 5.00
Begin VB.Form frmGetCell 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Выбор буферной ячейки"
   ClientHeight    =   4140
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCells 
      Height          =   3765
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
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
Public OK As Boolean
Public OUtID As String
Public OutCode As String
Private conn As ADODB.Connection


Private Sub CancelButton_Click()
  OK = False
  Me.Hide
End Sub

Private Sub Form_Load()

 ' запрашиваем свободное место в буферной зоне
  Dim bzrs As ADODB.Recordset
  Dim loccode As ADODB.Recordset
  Dim bzid As String
  Set conn = Manager.GetCustomObjects("refref")
  If conn.State <> adStateOpen Then
    conn.Open
  End If
  
  Set col = New Collection
  Set bzrs = conn.Execute( _
    "select id, code from location where id in (" & _
    " select  distinct location_id id from stock join location on location.id = stock.location_id " & _
    " Where stock.item_iD = " & itemid & _
    " group by location_id,location.description " & _
    " having count(*) < convert(integer, substring(location.description, 0,charindex(';',location.description,0))))" _
  )
  
  
  While Not bzrs.EOF
   Set db = New DBuffer
   db.id = bzrs!id
   db.name = bzrs!Code
   
   col.Add db
   lstCells.AddItem "+ " & db.name
   lstCells.ItemData(lstCells.NewIndex) = col.Count
   
   bzrs.MoveNext
  Wend
  
  bzrs.Close
  
  
  Set bzrs = conn.Execute( _
  "select top 500  id, code from location where description like '%;B%' and  id not in ( " & _
  " select location_id from stock where location_id is not null )")
  
  While Not bzrs.EOF
   Set db = New DBuffer
   db.id = bzrs!id
   db.name = bzrs!Code
   
   col.Add db
   lstCells.AddItem db.name
   lstCells.ItemData(lstCells.NewIndex) = col.Count
   bzrs.MoveNext
  Wend
  bzrs.Close
 
End Sub

Private Sub lstCells_DblClick()
OKButton_Click
End Sub

Private Sub OKButton_Click()
  Dim i As Long
  i = lstCells.ItemData(lstCells.ListIndex)
  Set db = col.Item(i)
  OUtID = db.id
  OutCode = db.name
  OK = True
  Me.Hide
  
End Sub
