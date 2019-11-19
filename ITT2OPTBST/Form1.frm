VERSION 5.00
Object = "{1801C003-859D-471D-BF31-D4428050324B}#2.1#0"; "MTZ_PANEL.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Параметры оптимизации"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPartner 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox txtPartia 
      Height          =   300
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "Партия"
      Top             =   3930
      Width           =   2550
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtGood 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txtmade_country 
      Height          =   300
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Страна производитель"
      Top             =   1770
      Width           =   2550
   End
   Begin VB.TextBox txtFactory 
      Height          =   300
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Завод"
      Top             =   2475
      Width           =   2550
   End
   Begin VB.TextBox txtKILL_NUMBER 
      Height          =   300
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   11
      ToolTipText     =   "№ бойни"
      Top             =   3180
      Width           =   2550
   End
   Begin MTZ_PANEL.DropButton cmdKILL_NUMBER 
      Height          =   300
      Left            =   2790
      TabIndex        =   12
      Tag             =   "refopen.ico"
      ToolTipText     =   "№ бойни"
      Top             =   3180
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Caption         =   ""
   End
   Begin MTZ_PANEL.DropButton cmdFactory 
      Height          =   300
      Left            =   2790
      TabIndex        =   9
      Tag             =   "refopen.ico"
      ToolTipText     =   "Завод"
      Top             =   2475
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Caption         =   ""
   End
   Begin MTZ_PANEL.DropButton cmdmade_country 
      Height          =   300
      Left            =   2790
      TabIndex        =   6
      Tag             =   "refopen.ico"
      ToolTipText     =   "Страна производитель"
      Top             =   1770
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Caption         =   ""
   End
   Begin MTZ_PANEL.DropButton cmdPartia 
      Height          =   300
      Left            =   2790
      TabIndex        =   15
      Tag             =   "refopen.ico"
      ToolTipText     =   "№ бойни"
      Top             =   3930
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Caption         =   ""
   End
   Begin VB.Label Label3 
      Caption         =   "Поклажедатель"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Партия:"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   240
      TabIndex        =   13
      Top             =   3600
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Артикул"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblmade_country 
      BackStyle       =   0  'Transparent
      Caption         =   "Страна производитель:"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   3000
   End
   Begin VB.Label lblFactory 
      BackStyle       =   0  'Transparent
      Caption         =   "Завод:"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   2145
      Width           =   3000
   End
   Begin VB.Label lblKILL_NUMBER 
      BackStyle       =   0  'Transparent
      Caption         =   "№ бойни:"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   240
      TabIndex        =   10
      Top             =   2850
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 165
Option Explicit

Public Manager As MTZManager.Main
Attribute Manager.VB_VarHelpID = 170
Public OK As Boolean
Attribute OK.VB_VarHelpID = 175


Private Sub cmdCancel_Click()
  OK = False
  Me.Hide
End Sub

Private Sub cmdFactory_CLick()
  On Error Resume Next
        Dim ID As String, brief As String
        If Manager.GetReferenceDialogEx2("ITTD_FACTORY", ID, brief) Then
          txtFactory.Tag = Left(ID, 38)
          txtFactory = brief
        End If
End Sub
Private Sub cmdFactory_MenuClick(ByVal sCaption As String)
          txtFactory.Tag = ""
          txtFactory = ""
End Sub

Private Sub cmdKILL_NUMBER_CLick()
  On Error Resume Next
        Dim ID As String, brief As String
        If Manager.GetReferenceDialogEx2("ITTD_KILLPLACE", ID, brief) Then
          txtKILL_NUMBER.Tag = Left(ID, 38)
          txtKILL_NUMBER = brief
        End If
End Sub
Private Sub cmdKILL_NUMBER_MenuClick(ByVal sCaption As String)
          txtKILL_NUMBER.Tag = ""
          txtKILL_NUMBER = ""
End Sub

Private Sub cmdmade_country_CLick()
  On Error Resume Next
        Dim ID As String, brief As String
        If Manager.GetReferenceDialogEx2("ITTD_COUNTRY", ID, brief) Then
          txtmade_country.Tag = Left(ID, 38)
          txtmade_country = brief
        End If
End Sub
Private Sub cmdmade_country_MenuClick(ByVal sCaption As String)
          txtmade_country.Tag = ""
          txtmade_country = ""
End Sub

Private Sub cmdOK_Click()
  OK = True
  Me.Hide
End Sub

Private Sub cmdPartia_MenuClick(ByVal sCaption As String)
 On Error Resume Next
        
          txtPartia.Tag = ""
          txtPartia = ""
        
End Sub
Private Sub cmdPartia_Click()
 On Error Resume Next
        Dim ID As String, brief As String
        If Manager.GetReferenceDialogEx2("ITTD_PART", ID, brief) Then
          txtPartia.Tag = Left(ID, 38)
          txtPartia = brief
        End If
End Sub

Private Sub Form_Load()

  txtmade_country.Tag = ""
  txtmade_country = ""
  LoadBtnPictures cmdmade_country, cmdmade_country.Tag
  cmdmade_country.RemoveAllMenu
  cmdmade_country.AddMenu "Очистить"
  txtFactory.Tag = ""
  txtFactory = ""
  LoadBtnPictures cmdFactory, cmdFactory.Tag
  cmdFactory.RemoveAllMenu
  cmdFactory.AddMenu "Очистить"
  txtKILL_NUMBER.Tag = ""
  txtKILL_NUMBER = ""
  LoadBtnPictures cmdKILL_NUMBER, cmdKILL_NUMBER.Tag
  cmdKILL_NUMBER.RemoveAllMenu
  cmdKILL_NUMBER.AddMenu "Очистить"
  txtPartia.Tag = ""
  txtPartia = ""
  LoadBtnPictures cmdPartia, cmdPartia.Tag
  cmdPartia.RemoveAllMenu
  cmdPartia.AddMenu "Очистить"
End Sub
