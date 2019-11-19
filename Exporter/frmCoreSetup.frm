VERSION 5.00
Begin VB.Form frmCoreSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Параметры Базы  данных Core IMS"
   ClientHeight    =   2580
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4455
   Icon            =   "frmCoreSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTest 
      Caption         =   "Тест"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox txtUsr 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtDB 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtSrv 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Пароль пользователя"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Пользоваель SQL"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "База данных"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Сервер БД"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmCoreSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
  Me.Hide
End Sub

Private Sub cmdTest_Click()
 On Error Resume Next
 Dim conn As ADODB.Connection
  Set conn = New ADODB.Connection
  conn.Provider = "SQLoledb"
  conn.ConnectionString = "Server=" & txtSrv & ";DataBase=" & txtDB & ";UID=" & txtUsr & ";Pwd=" & txtPass & ";"
  conn.Open
  If conn.State = ADODB.adStateOpen Then
    conn.Close
    MsgBox "Соединение установлено"
  End If
  Set conn = Nothing
End Sub



Private Sub Form_Load()
    txtSrv = GetSetting("RBH", "ITTSETTINGS", "CORESRV", "")
    txtDB = GetSetting("RBH", "ITTSETTINGS", "COREDB", "")
    txtUsr = GetSetting("RBH", "ITTSETTINGS", "COREUSR", "")
    txtPass = GetSetting("RBH", "ITTSETTINGS", "COREPASS", "")
End Sub

Private Sub OKButton_Click()
  Dim conn As ADODB.Connection
  Set conn = New ADODB.Connection
  conn.Provider = "SQLoledb"
  conn.ConnectionString = "Server=" & txtSrv & ";DataBase=" & txtDB & ";UID=" & txtUsr & ";Pwd=" & txtPass & ";"
  conn.Open
  If conn.State = ADODB.adStateOpen Then
    conn.Close
    
    SaveSetting "RBH", "ITTSETTINGS", "CORESRV", txtSrv
    SaveSetting "RBH", "ITTSETTINGS", "COREDB", txtDB
    SaveSetting "RBH", "ITTSETTINGS", "COREUSR", txtUsr
    SaveSetting "RBH", "ITTSETTINGS", "COREPASS", txtPass
    
  Else
    MsgBox "Неверные параметры соединения"
  End If
  Set conn = Nothing
  Me.Hide
End Sub
