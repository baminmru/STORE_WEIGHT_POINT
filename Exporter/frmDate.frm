VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Дата для расчета выморозки"
   ClientHeight    =   1260
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Format          =   56950785
      CurrentDate     =   39198
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
      Caption         =   "Дата"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmDate"
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

Private Sub Form_Load()
  dtpDate.Value = Date
End Sub

Private Sub OKButton_Click()
  If dtpDate.Value < Date Then
    MsgBox "Дата должна быть не раньше текущего дня "
    Exit Sub
  End If
  OK = True
  Me.Hide
End Sub
