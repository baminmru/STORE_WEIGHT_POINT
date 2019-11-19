VERSION 5.00
Begin VB.Form frmWSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Настройка весов"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   Icon            =   "frmWSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMoroz 
      Caption         =   "Не проверять выморозку"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   3975
   End
   Begin VB.CheckBox chkRestore 
      Caption         =   "Восстанавливать вес при отгрузке"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   3975
   End
   Begin VB.CheckBox chkSound 
      Caption         =   "Включить звук"
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   3915
   End
   Begin VB.CheckBox chkUseEmu 
      Caption         =   "Эмулятор весов"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1635
      Width           =   3870
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отменить"
      Height          =   270
      Left            =   2640
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   270
      Left            =   1080
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtportSetup 
      Height          =   285
      Left            =   135
      TabIndex        =   3
      Text            =   "4800,e,8,1"
      ToolTipText     =   "для BT-150 =4800,e,8,1"
      Top             =   1185
      Width           =   3855
   End
   Begin VB.TextBox txtPort 
      Height          =   300
      Left            =   135
      TabIndex        =   1
      Text            =   "1"
      Top             =   480
      Width           =   3870
   End
   Begin VB.Label Label2 
      Caption         =   "Настройки:"
      Height          =   270
      Left            =   135
      TabIndex        =   2
      Top             =   930
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Номер COM порта (1..4):"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   2010
   End
End
Attribute VB_Name = "frmWSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 740
Option Explicit
'Форма настройки параметров весов


' закрытие без сохранения
Private Sub cmdCancel_Click()
Me.Hide
End Sub


'Закрытие с сохранением
Private Sub cmdOK_Click()
Me.Hide
SaveSetting "RBH", "ITTSETTINGS", "WSETUP", txtportSetup.Text
SaveSetting "RBH", "ITTSETTINGS", "WPORT", txtPort.Text
SaveSetting "RBH", "ITTSETTINGS", "EMULATOR", chkUseEmu.Value = vbChecked
SaveSetting "RBH", "ITTSETTINGS", "SOUND", chkSound.Value = vbChecked
SaveSetting "RBH", "ITTSETTINGS", "RESTORE", chkRestore.Value = vbChecked
SaveSetting "RBH", "ITTSETTINGS", "MOROZ", chkMoroz.Value = vbChecked
End Sub


'Начальная загрузка данных при открытии формы
Private Sub Form_Load()
 txtportSetup.Text = GetSetting("RBH", "ITTSETTINGS", "WSETUP", "4800,e,8,1")
 txtPort.Text = GetSetting("RBH", "ITTSETTINGS", "WPORT", 1)
 If GetSetting("RBH", "ITTSETTINGS", "EMULATOR", "False") = "False" Then
  chkUseEmu.Value = vbUnchecked
 Else
  chkUseEmu.Value = vbChecked
 End If
 If GetSetting("RBH", "ITTSETTINGS", "SOUND", "False") = "False" Then
  chkSound.Value = vbUnchecked
 Else
  chkSound.Value = vbChecked
 End If
 
 If GetSetting("RBH", "ITTSETTINGS", "RESTORE", "False") = "False" Then
  chkRestore.Value = vbUnchecked
 Else
  chkRestore.Value = vbChecked
 End If
 
  If GetSetting("RBH", "ITTSETTINGS", "MOROZ", "False") = "False" Then
  chkMoroz.Value = vbUnchecked
 Else
  chkMoroz.Value = vbChecked
 End If
 
End Sub

