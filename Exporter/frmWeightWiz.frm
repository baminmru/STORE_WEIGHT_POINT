VERSION 5.00
Begin VB.Form frmWeightWiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Прием груза"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10305
   Icon            =   "frmWeightWiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   10305
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отмена"
      Height          =   525
      Left            =   5940
      TabIndex        =   4
      Top             =   7305
      Width           =   2100
   End
   Begin VB.Frame pnlDocs 
      Caption         =   "Печать документов"
      Height          =   7035
      Left            =   1965
      TabIndex        =   3
      Top             =   1755
      Visible         =   0   'False
      Width           =   8370
   End
   Begin VB.Frame pnlWeighting 
      Caption         =   "Взвешивание"
      Height          =   7020
      Left            =   2115
      TabIndex        =   2
      Top             =   1125
      Visible         =   0   'False
      Width           =   8355
   End
   Begin VB.Frame pnlChoose 
      Caption         =   "Выбор заявки"
      Height          =   7050
      Left            =   1575
      TabIndex        =   1
      Top             =   150
      Width           =   8640
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">> Далее >>"
      Height          =   555
      Left            =   8145
      TabIndex        =   0
      Top             =   7290
      Width           =   2040
   End
   Begin VB.Image cmdPrint 
      Height          =   1290
      Left            =   120
      Picture         =   "frmWeightWiz.frx":0442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1275
   End
   Begin VB.Image cmdWeighting 
      Height          =   1200
      Left            =   120
      Picture         =   "frmWeightWiz.frx":16A2
      Stretch         =   -1  'True
      Top             =   135
      Width           =   1260
   End
   Begin VB.Image cmdSelZ 
      Height          =   1095
      Left            =   105
      Picture         =   "frmWeightWiz.frx":35C0
      Stretch         =   -1  'True
      Top             =   135
      Width           =   1335
   End
End
Attribute VB_Name = "frmWeightWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WizState As Integer
Private Sub Image1_Click()

End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdNext_Click()
  WizState = WizState + 1
  AcceptState
End Sub


Private Sub AcceptState()
  cmdPrint.Visible = False
  cmdSelZ.Visible = False
  cmdWeighting.Visible = False
  pnlChoose.Visible = False
  pnlDocs.Visible = False
  pnlWeighting.Visible = False
  If WizState = 1 Then
    cmdSelZ.Visible = True
    pnlChoose.Visible = True
  End If
  If WizState = 2 Then
    cmdWeighting.Visible = True
    pnlWeighting.Visible = True
  End If
  If WizState = 3 Then
      cmdPrint.Visible = True
      pnlDocs.Visible = True
      cmdNext.Caption = "Готово"
  End If
  If WizState > 3 Then
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  With pnlChoose
  .Visible = False
  .Top = 5 * Screen.TwipsPerPixelY
  .Left = 1575
  .Width = Me.ScaleWidth - .Left - 5 * Screen.TwipsPerPixelX
  .Height = cmdNext.Top - 10 * Screen.TwipsPerPixelY
  End With
  
  With pnlWeighting
  .Visible = False
  .Top = 5 * Screen.TwipsPerPixelY
  .Left = 1575
  .Width = Me.ScaleWidth - .Left - 5 * Screen.TwipsPerPixelX
  .Height = cmdNext.Top - 10 * Screen.TwipsPerPixelY
  End With
  With pnlDocs
  .Visible = False
  .Top = 5 * Screen.TwipsPerPixelY
  .Left = 1575
  .Width = Me.ScaleWidth - .Left - 5 * Screen.TwipsPerPixelX
  .Height = cmdNext.Top - 10 * Screen.TwipsPerPixelY
  End With
  

WizState = 1
AcceptState
End Sub
