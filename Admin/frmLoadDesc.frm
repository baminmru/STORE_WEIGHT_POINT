VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLoadDesc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Загрузить описание типа"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6855
   Icon            =   "frmLoadDesc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   315
      Left            =   6405
      TabIndex        =   2
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4920
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Загрузить"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Путь к файлу:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1785
   End
End
Attribute VB_Name = "frmLoadDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TypeName As String
Public site As String

Public OK As Boolean
Private ID  As String
Public NewObject  As Object
Private types As Collection



Private Sub CancelButton_Click()
OK = False
Set types = Nothing
Me.Hide
End Sub







Private Sub OKButton_Click()
  
  On Error GoTo bye
  Dim item As OBJECTTYPE
  Dim xdom As MSXML2.DOMDocument
  Set xdom = New MSXML2.DOMDocument
  xdom.Load txtPath.Text
  ID = xdom.lastChild.Attributes.getNamedItem("ID").nodeValue
  If model.OBJECTTYPE.item(ID) Is Nothing Then
    model.OBJECTTYPE.Add ID
  End If
  Set item = model.OBJECTTYPE.item(ID)
  item.XMLLoad xdom.lastChild, 1
  item.BatchUpdate
  Set xdom = Nothing

bye:
End Sub

Private Sub cmdPath_Click()
  On Error Resume Next
  
  On Error GoTo bye
  Dim fn As String
  cdlg.CancelError = True
  cdlg.Filter = "Документ XML |*.XML"
  cdlg.DefaultExt = "XML"
  'cdlg.FileName = App.path & "\" & item.ID & ".xml"
  cdlg.Flags = cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNFileMustExist
  cdlg.ShowOpen
  txtPath = cdlg.FileName
  

bye:
End Sub

