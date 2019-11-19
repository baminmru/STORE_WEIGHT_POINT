VERSION 5.00
Begin VB.Form frmGetCell 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "¬ыбор  €чейки"
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
      Caption         =   "«она хранени€"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ѕуферна€ зона"
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
      Caption         =   "Ўаблон:"
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
Attribute VB_HelpID = 165

Option Explicit
'окно выбора €чейки дл€ постановки паллеты


Dim Col As Collection
Dim db As DBuffer

Public itemid As String
Attribute itemid.VB_VarHelpID = 180
Public country As String
Attribute country.VB_VarHelpID = 170
Public factory As String
Attribute factory.VB_VarHelpID = 175
Public killplace As String
Attribute killplace.VB_VarHelpID = 185
'изменено
Public OK As Boolean
Attribute OK.VB_VarHelpID = 190
Public OUtID As String
Attribute OUtID.VB_VarHelpID = 200
Public OutCode As String
Attribute OutCode.VB_VarHelpID = 195
Private conn As ADODB.Connection
Private cells As Long
Public PTYPE As Double
Attribute PTYPE.VB_VarHelpID = 205


Private Sub CancelButton_Click()
  OK = False
  Me.Hide
End Sub

Private Sub FillList(ByVal buffer As Boolean)

 ' запрашиваем свободное место в буферной зоне
  Dim bzrs As ADODB.Recordset
  Dim loccode As ADODB.Recordset
  Dim bzid As String
  Set conn = GetCoreConn
  If conn.State <> adStateOpen Then
    conn.open
  End If
  
  lstCells.Clear
  
  Set Col = New Collection
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
   
   
   ' изменен запрос дл€ поиска свободного места в €чейке
   ' добавлена в запрос таблица pallet  измено условие дл€ определени€ зан€того объема
   
   
   If buffer Then
     qry = "select id, code from location where id in (" & _
     " select  distinct stock.location_id id from stock join location on location.id = stock.location_id " & _
     " join pallet on stock.pallet_id = pallet.id Where location.description like '%;B%' and location.active='Y' " & _
     " group by stock.location_id,location.description,loc_qty_pal  " & _
     " having sum(case when pallet.type='E' then 1  else 1.25  end  ) < loc_qty_pal - " & MyRound2(PTYPE) & " )" & _
     " and id in (" & _
     " select  distinct location_id id from stock join location on location.id = stock.location_id " & _
     " Where location.description like '%;B%' and location.active='Y' and stock.item_iD = " & itemid & subquery & ")"
   Else
     qry = "select id, code from location where id in (" & _
     " select  distinct stock.location_id id from stock join location on location.id = stock.location_id " & _
     " join pallet on stock.pallet_id = pallet.id Where location.description not like '%;B%' and location.active='Y' " & _
     " group by stock.location_id,location.description,loc_qty_pal  " & _
     " having sum(case  when pallet.type='E' then 1  else 1.25  end  )  < loc_qty_pal - " & MyRound2(PTYPE) & " ) " & _
     " and id in (" & _
     " select  distinct location_id id from stock join location on location.id = stock.location_id " & _
     " Where location.description not like '%;B%' and location.active='Y' and stock.item_iD = " & itemid & subquery & " )"
   End If
   
    
  Set bzrs = conn.Execute(qry)
  
  
  While Not bzrs.EOF
   Set db = New DBuffer
   db.id = bzrs!id
   db.Name = bzrs!code
   
   Col.Add db
   lstCells.AddItem "+ " & db.Name
   lstCells.ItemData(lstCells.NewIndex) = Col.Count
   
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
  
  ' добавл€ем пустые €чейки в список
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
   
   Col.Add db
   lstCells.AddItem db.Name
   lstCells.ItemData(lstCells.NewIndex) = Col.Count
   bzrs.MoveNext
  Wend
  bzrs.Close
  Me.Caption = "¬ыбор €чейки (" & Col.Count & ")"
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
  Set db = Col.Item(i)
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
