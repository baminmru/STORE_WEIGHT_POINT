VERSION 5.00
Begin VB.Form frmStickerRpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Диапазон номеров поддонов"
   ClientHeight    =   1605
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4800
   Icon            =   "frmStickerRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTo 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "1"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtFrom 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "1"
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "По:"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "C:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmStickerRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public OK As Boolean

Private Sub CancelButton_Click()
OK = False
Me.Hide
End Sub

Private Sub OKButton_Click()
  If IsNumeric(txtFrom.Text) And IsNumeric(txtTo.Text) Then
    OK = True
    Me.Hide
  Else
    MsgBox "Ожидались числовые значения"
  End If
End Sub
