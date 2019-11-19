VERSION 5.00
Begin VB.Form frmLocSize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Управление размером ячейки"
   ClientHeight    =   4575
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNewSize 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox txtCur 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox txtPlan 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtLoaded 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtLoc 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3015
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
   Begin VB.Label Label4 
      Caption         =   "Новое значение размера"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Текущий размер"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label lll 
      Caption         =   "Максимальный размер"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Загружено поддонов"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Номер ячейки"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmLocSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Public OK As Boolean

Private Sub CancelButton_Click()
OK = False
Me.Hide
End Sub

Private Sub OKButton_Click()
  Dim sz As Integer
  Dim loc As String
  loc = txtLoc.Text
  If Len(loc) = 11 Then
    Set conn = Manager.GetCustomObjects("refref")
    If IsNumeric(txtNewSize) Then
      sz = Val(txtNewSize)
      Set rs = conn.Execute("select * from v_bami_locsize where code='" & loc & "'")
      If rs Is Nothing Then
        Exit Sub
      End If
      If Not rs.EOF Then
        If sz > rs!plan_loc_qty Then
              
              Call MsgBox("Текущий размер не может превышать плановый", vbOKOnly + vbCritical, "Ошибка ввода")
              Exit Sub
        End If
        
        If sz < rs!cur_qty Then
              Call MsgBox("Текущий размер не может быть меньше текущей загрузки", vbOKOnly + vbCritical, "Ошибка ввода")
              Exit Sub
        End If
        conn.Execute "update location set loc_qty_pal = " & sz & " where code = '" & loc & "'"
        
        OK = True
        Me.Hide
      End If
    End If
  End If
  
    
End Sub

Private Sub txtLoc_Change()
Dim loc As String
loc = txtLoc.Text
If Len(loc) = 11 Then
  readSize (loc)
End If
End Sub


Private Sub readSize(ByVal loc As String)
  Set conn = Manager.GetCustomObjects("refref")
  Set rs = conn.Execute("select * from v_bami_locsize where code='" & loc & "'")
  If rs Is Nothing Then
    Exit Sub
  End If
  If Not rs.EOF Then
    txtPlan = rs!plan_loc_qty
    txtLoaded = rs!cur_qty
    txtCur = rs!loc_qty_pal
  End If

End Sub

