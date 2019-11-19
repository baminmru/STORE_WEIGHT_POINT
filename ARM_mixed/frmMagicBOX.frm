VERSION 5.00
Begin VB.Form frmMagicBOX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Сообщение об ошибке"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отмена"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8415
      Begin VB.Label txtMessage 
         Alignment       =   2  'Center
         ForeColor       =   &H000000FF&
         Height          =   3735
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   7080
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtMagic 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   4560
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Кодовое слово для отмены повтора"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   4320
      Width           =   3255
   End
End
Attribute VB_Name = "frmMagicBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OK As Boolean

Private Sub cmdCancel_Click()
OK = False
Me.Hide
End Sub

Private Sub cmdOK_Click()
OK = True
Me.Hide
End Sub


Private Sub Form_Load()
If LastMagic <> "" Then
  txtMagic = LastMagic
  log.message "Введено сохранненое магическое слово"
End If
End Sub

Private Sub txtMagic_Change()
 If Len(txtMagic) = 3 Then
  If LCase(txtMagic) = GetMagicWord(Date) Then
    cmdCancel.Visible = True
    LastMagic = LCase(txtMagic)
    log.message "Введено магическое слово"
  End If
 End If
End Sub
