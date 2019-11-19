VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Общие настройки"
   ClientHeight    =   2640
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCells 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Text            =   "100"
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CheckBox chkPrintSicker 
      Caption         =   "Печатать стикер"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Обновить описание АРМ а"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CheckBox chkPrintCell 
      Caption         =   "Печатать бланк на ячейку"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3615
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
   Begin VB.Label Label1 
      Caption         =   "Количество записей в окне подбора ячейки"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   3735
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub cmdRefresh_Click()
  frmMain.SynchronizeARMDescription
End Sub

Private Sub Form_Load()
  If GetSetting("RBH", "ITTSETTINGS", "PCELL", 1) = 1 Then
    chkPrintCell.Value = vbChecked
  End If
  
  If GetSetting("RBH", "ITTSETTINGS", "PSTICKER", 1) = 1 Then
    chkPrintSicker.Value = vbChecked
  End If
  
  txtCells = GetSetting("RBH", "ITTSETTINGS", "SHOWCELLS", 100)
End Sub

Private Sub OKButton_Click()
  If chkPrintCell.Value = vbChecked Then
    Call SaveSetting("RBH", "ITTSETTINGS", "PCELL", 1)
  Else
    Call SaveSetting("RBH", "ITTSETTINGS", "PCELL", 0)
  End If
  
  If chkPrintSicker.Value = vbChecked Then
    Call SaveSetting("RBH", "ITTSETTINGS", "PSTICKER", 1)
  Else
    Call SaveSetting("RBH", "ITTSETTINGS", "PSTICKER", 0)
  End If
  
  If Val("0" & txtCells) < 100 Then
    Call SaveSetting("RBH", "ITTSETTINGS", "SHOWCELLS", 100)
  Else
    Call SaveSetting("RBH", "ITTSETTINGS", "SHOWCELLS", Val("0" & txtCells))
  End If
  Unload Me
End Sub
