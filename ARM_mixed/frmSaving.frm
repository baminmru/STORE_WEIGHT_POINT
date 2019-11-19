VERSION 5.00
Begin VB.Form frmSaving 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Сохранение данных"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3360
      Top             =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Идет сохранение данных в CORE IMS. Ждите."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmSaving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Red As Boolean

Private Sub Timer1_Timer()
  
  If Red Then
    Label1.ForeColor = RGB(255, 0, 0)
  Else
    Label1.ForeColor = RGB(0, 255, 0)
  End If
  Red = Not Red
  Me.Caption = "Сохранение данных " & Now
End Sub
