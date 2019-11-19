VERSION 5.00
Begin VB.Form frmPrintSticker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Печать стикера на поддон"
   ClientHeight    =   1485
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmPrintSticker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   960
   End
   Begin VB.TextBox txt3Poddon 
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2655
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
   Begin VB.Label Label6 
      Caption         =   "Поддон №"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmPrintSticker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public OK As Boolean
Public poddon As ITTPL.ITTPL_DEF

Private Sub CancelButton_Click()
  Timer1.Enabled = False
  OK = False
  Me.Hide
  Set poddon = Nothing
End Sub

Private Sub OKButton_Click()
     Set poddon = Nothing
      Set poddon = FindPoddon(txt3Poddon)
      If Not poddon Is Nothing Then
        OK = True
        Me.Hide
      Else
        MsgBox "Номер паддона: " & txt3Poddon & "  не зарегистрирован"
      End If
      
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If txt3Poddon = "" Then
      txt3Poddon.SetFocus
    End If
End Sub
