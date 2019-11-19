VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmDates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Диапазон дат"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   Icon            =   "frmDates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   420
      Left            =   1695
      TabIndex        =   5
      Top             =   1680
      Width           =   1440
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   135
      TabIndex        =   4
      Top             =   1650
      Width           =   1455
   End
   Begin VB.CheckBox lbldfrom 
      Caption         =   "С:"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   165
      TabIndex        =   3
      Top             =   105
      Width           =   3000
   End
   Begin VB.CheckBox lbldTo 
      Caption         =   "По:"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   165
      TabIndex        =   1
      Top             =   810
      Width           =   3000
   End
   Begin MSComCtl2.DTPicker dtpdTo 
      Height          =   300
      Left            =   180
      TabIndex        =   0
      ToolTipText     =   "По"
      Top             =   1140
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   49152003
      CurrentDate     =   38311
   End
   Begin MSComCtl2.DTPicker dtpdfrom 
      Height          =   300
      Left            =   165
      TabIndex        =   2
      ToolTipText     =   "С"
      Top             =   435
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   49152003
      CurrentDate     =   38311
   End
End
Attribute VB_Name = "frmDates"
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
dtpdfrom.Value = Date - 30
dtpdTo.Value = Date + 1
End Sub
