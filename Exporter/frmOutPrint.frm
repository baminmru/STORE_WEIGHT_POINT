VERSION 5.00
Begin VB.Form frmOutPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Форма отчета"
   ClientHeight    =   1815
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmOutPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optActRas 
      Caption         =   "Акт о весовых расхождениях"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4335
   End
   Begin VB.OptionButton optOTB 
      Caption         =   "Отборочный лист"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Value           =   -1  'True
      Width           =   4335
   End
   Begin VB.OptionButton optSRV 
      Caption         =   "Отчет об оказаных услугах"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4215
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
End
Attribute VB_Name = "frmOutPrint"
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
  OK = True
  Me.Hide
End Sub

