VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmXL 
   Caption         =   "Отчет обуслугах"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDForm 
      Height          =   285
      Left            =   4560
      TabIndex        =   3
      Text            =   "YYYY-MM-DD"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdSaveXL 
      Caption         =   "Открыть в  Exсel"
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1665
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgr 
      Height          =   5970
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9825
      _cx             =   17330
      _cy             =   10530
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label1 
      Caption         =   "Формат даты для выгрузки:"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub cmdSaveAs_Click()
'vfgr.SaveGrid App.Path & "\test.csv", flexFileCommaText, True
'
''vfgr.Select 0, 0, vfgr.Rows - 1, vfgr.Cols - 1
'
'End Sub




Private Sub cmdSaveXL_Click()
Dim ex As Object
Dim excel As Object

Set excel = CreateObject("Excel.Application")
With excel.Workbooks.Add
Set ex = .Worksheets.Item(1)
End With

Dim s()
ReDim s(0 To vfgr.Cols - 1)

For i = 0 To vfgr.Rows - 1

  For j = 0 To vfgr.Cols - 1
    If IsNumeric(vfgr.TextMatrix(i, j)) Then
      s(j) = MyRound2(vfgr.TextMatrix(i, j))
    ElseIf IsDate(vfgr.TextMatrix(i, j)) Then
      On Error Resume Next
      s(j) = Format(vfgr.TextMatrix(i, j), txtDForm.Text)
    Else
      s(j) = vfgr.TextMatrix(i, j)
    End If
  Next
  
  ex.Range(ex.cells(i + 2, 1), ex.cells(i + 2, vfgr.Cols)).Value = s
Next


excel.Selection.AutoFormat Format:=12, Number:=True, Font _
        :=True, Alignment:=True, Border:=True, Pattern:=True, Width:=True
ex.Range(ex.cells(1, 1), ex.cells(1, 1)).Value = Me.Caption
excel.ActiveWindow.Visible = True
excel.Visible = True

SaveSetting App.EXEName, "CFG", "DFormat", txtDForm
End Sub

Private Sub Form_Load()
txtDForm = GetSetting(App.EXEName, "CFG", "DFormat", "YYYY-MM-DD")
End Sub

Private Sub Form_Resize()
On Error Resume Next
vfgr.Move 0, vfgr.Top, Me.ScaleWidth, Me.ScaleHeight - vfgr.Top


End Sub


