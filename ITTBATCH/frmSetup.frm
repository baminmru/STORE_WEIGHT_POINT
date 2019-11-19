VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Настройка синхронизатора"
   ClientHeight    =   5520
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkLog 
      Caption         =   "Включить логирование"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   5040
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   5040
      Picture         =   "frmSetup.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtCLI 
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Text            =   "Unilever"
      Top             =   4440
      Width           =   3855
   End
   Begin VB.TextBox txtSyncDB 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox txtUSER 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox txtDB 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox txtServer 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   3855
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
      Caption         =   "Код клиента"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "База данных для синхронизации"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Пароль"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Пользователь базы данных"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "База данных (CORE IMS)"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Сервер (CORE IMS)"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
  Me.Hide
End Sub



Private Sub Form_Load()
  txtDB.Text = GetSetting("ITT", "ITTBATCH", "DB", "")
  txtServer.Text = GetSetting("ITT", "ITTBATCH", "SRV", "")
  txtUSER.Text = GetSetting("ITT", "ITTBATCH", "USER", "")
  txtPass.Text = GetSetting("ITT", "ITTBATCH", "PASS", "")
  txtSyncDB.Text = GetSetting("ITT", "ITTBATCH", "SYNCDB", "")
  txtCLI.Text = GetSetting("ITT", "ITTBATCH", "CLI", "")
  If GetSetting("ITT", "ITTBATCH", "LOG", "1") = "1" Then
    chkLog.Value = vbChecked
  Else
    chkLog.Value = vbUnchecked
  End If

End Sub

Private Sub OKButton_Click()
  Call SaveSetting("ITT", "ITTBATCH", "DB", txtDB.Text)
  Call SaveSetting("ITT", "ITTBATCH", "SRV", txtServer.Text)
  Call SaveSetting("ITT", "ITTBATCH", "USER", txtUSER.Text)
  Call SaveSetting("ITT", "ITTBATCH", "PASS", txtPass.Text)
  Call SaveSetting("ITT", "ITTBATCH", "SYNCDB", txtSyncDB.Text)
  Call SaveSetting("ITT", "ITTBATCH", "CLI", txtCLI.Text)
  Call SaveSetting("ITT", "ITTBATCH", "CS", "Driver={SQL Server};Server=" & txtServer.Text & ";Database=" & txtSyncDB.Text & ";Uid=" & txtUSER.Text & ";Pwd=" & txtPass.Text & ";DataSource=" & txtServer.Text)
    If chkLog.Value = vbChecked Then
    Call SaveSetting("ITT", "ITTBATCH", "LOG", "1")
  Else
    Call SaveSetting("ITT", "ITTBATCH", "LOG", "0")
  End If
  Me.Hide
End Sub
