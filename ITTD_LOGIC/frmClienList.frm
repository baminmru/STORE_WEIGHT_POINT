VERSION 5.00
Begin VB.Form frmClienList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Выбор поклажедателя"
   ClientHeight    =   5280
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstClient 
      Height          =   4545
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
End
Attribute VB_Name = "frmClienList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 10

Option Explicit
Public ClientText As String
Attribute ClientText.VB_VarHelpID = 15
Public manager As MTZManager.Main
Attribute manager.VB_VarHelpID = 20

Private Sub CancelButton_Click()
Me.Hide
End Sub



Private Sub Form_Load()
 Dim conn As ADODB.Connection
 Set conn = manager.GetCustomObjects("refref")
 
 Dim rs As ADODB.Recordset
 Set rs = conn.Execute("select street1 from ADDRESS")
 lstClient.Clear
 While Not rs.EOF
  lstClient.AddItem rs!street1
  rs.MoveNext
 Wend
 rs.Close
 Set rs = Nothing
 
End Sub

Private Sub OKButton_Click()
If lstClient.ListIndex >= 0 Then
  ClientText = lstClient.List(lstClient.ListIndex)
  Me.Hide
Else
  MsgBox "Слудует выбрать поклажедателя", vbOKOnly + vbCritical, "Ошибка выбора"
End If
End Sub
