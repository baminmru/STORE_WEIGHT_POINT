VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMovings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Отработка оптимизации"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отменить"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Сохранить"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   4080
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid srvGrid 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _cx             =   14208
      _cy             =   7011
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   300
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMovings.frx":0000
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
      Editable        =   1
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
End
Attribute VB_Name = "frmMovings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public movetask As ITTOPT.Application
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset


Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdSave_Click()
    Dim i As Long
    Dim OK As Boolean
    If movetask.ITTOPT_MOVE.Count = 0 Then
    Me.Hide
    Exit Sub
    End If
    pb.Min = 0
    pb.Max = movetask.ITTOPT_MOVE.Count
    pb.Value = 0
    pb.Visible = True
    For i = 1 To movetask.ITTOPT_MOVE.Count
     pb.Value = i
      movetask.ITTOPT_MOVE.Item(i).ThePalletteNum = srvGrid.TextMatrix(i, 3)
      movetask.ITTOPT_MOVE.Item(i).save
      OK = OK And RegisterMove(movetask.ITTOPT_MOVE.Item(i))
    Next
    pb.Visible = False

  ' состояния для типа:ITTOPT Задание на перемещения
  ' "{0A7FC795-E787-4D17-9689-96EFFF8F0D9D}" 'Задание исполнено
  ' "{300483B2-1D94-4A33-8ADF-ABF32E72E57B}" 'Оформлено
  ' "{C861FA15-0DF6-42D4-BCE9-2B38C3E6C0CB}" 'Оформляется
  If OK Then
    MsgBox "Есть ошибки при сохранении"
    INIT
  Else
    Me.Hide
  End If
End Sub

Private Sub Form_Load()
  INIT
End Sub

Private Sub srvGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
  
  Dim code As String
  code = Val(srvGrid.TextMatrix(Row, 3))
  If Len(srvGrid.TextMatrix(Row, 3)) <> 6 Then
    srvGrid.TextMatrix(Row, 3) = Right("000000" & srvGrid.TextMatrix(Row, 3), 6)
  End If
  
  
  Set rs = conn.Execute("select pallet.code from stock join pallet on stock.pallet_id=pallet.id join location on location.id=stock.location_id Where pallet_Status Is Null and pallet.code ='" & code & "' and location.code='" & srvGrid.TextMatrix(Row, 1) & "'")
  'Set rs = conn.Execute("select pallet.code from stock join pallet on stock.pallet_id=pallet.id where pallet_Status is null and pallet.code ='" & code & "'")
  If rs.EOF Then
    MsgBox "Поддон " & srvGrid.TextMatrix(Row, 3) & " не зарегистрирован в ячейке " & srvGrid.TextMatrix(Row, 1) & "!", vbExclamation + vbOKOnly, "Внимание!!!"
    srvGrid.TextMatrix(Row, 3) = ""
  End If
  
End Sub

Private Sub srvGrid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  If Col <> 3 Then Cancel = True
  If srvGrid.TextMatrix(Row, 4) = "OK" Then Cancel = True
  
End Sub

Private Sub INIT()
  Dim i As Long
  srvGrid.Cols = 5
  srvGrid.Rows = movetask.ITTOPT_MOVE.Count + 1
  
  movetask.ITTOPT_MOVE.Sort = "sequence"
  For i = 1 To movetask.ITTOPT_MOVE.Count
      srvGrid.TextMatrix(i, 0) = movetask.ITTOPT_MOVE.Item(i).sequence
      srvGrid.TextMatrix(i, 1) = movetask.ITTOPT_MOVE.Item(i).FromLocation
      srvGrid.TextMatrix(i, 2) = movetask.ITTOPT_MOVE.Item(i).ToLocation
      srvGrid.TextMatrix(i, 3) = movetask.ITTOPT_MOVE.Item(i).ThePalletteNum
      If movetask.ITTOPT_MOVE.Item(i).ISMoved Then
        srvGrid.TextMatrix(i, 4) = "OK"
      Else
        srvGrid.TextMatrix(i, 4) = ""
      End If
  Next
  Set conn = Manager.GetCustomObjects("refref")
  
End Sub


Private Function BeforeChangeStatus(Item As Object, NewStatus As String) As Boolean
  Dim logic As Object
  Dim result As Boolean
  result = True
  On Error Resume Next
  Set logic = CreateObject(Item.TypeName & "BST.BEFORESTATUS")
  If Not logic Is Nothing Then
    result = logic.Check(Item, NewStatus, MyUser, Item.TypeName)
    Set logic = Nothing
  End If
  BeforeChangeStatus = result
End Function

Private Sub ChangeSataus()
On Error Resume Next

  
  If RoleDocCanSwitchStatus(movetask) Then
    If BeforeChangeStatus(movetask, "{0A7FC795-E787-4D17-9689-96EFFF8F0D9D}") Then
      movetask.StatusID = "{0A7FC795-E787-4D17-9689-96EFFF8F0D9D}"
    End If
  Else
  MsgBox "Для этой роли не разрешено изменение статуса документа", vbOKOnly + vbInformation, "Изменение состояния"
  End If
End Sub



Private Function RegisterMove(ByRef MD As ITTOPT_MOVE) As Boolean
  Dim rs As ADODB.Recordset
  Dim strs As ADODB.Recordset
  Dim locfrs As ADODB.Recordset
  Dim loctrs As ADODB.Recordset
  Dim palrs As ADODB.Recordset
  Dim itemrs As ADODB.Recordset
  
  Dim cmd As ADODB.Command
  RegisterMove = True
  
  If MD.ISMoved = Boolean_Net And Len(MD.ThePalletteNum) = 6 Then
  
    Set rs = conn.Execute("select * from v_bami_stock where loc_code='" & MD.FromLocation & "' and pallet_code=" & MD.ThePalletteNum)
    If Not rs.EOF Then
      Set strs = conn.Execute("select * from stock where id=" & rs!Stock_ID)
      Set loctrs = conn.Execute("select * from location where code='" & MD.ToLocation & "'")
      Set locfrs = conn.Execute("select * from location where code='" & MD.FromLocation & "'")
      Set itemrs = conn.Execute("select * from item where id=" & strs!item_id)
      Set palrs = conn.Execute("select * from pallet where id=" & strs!pallet_id)
      
      
      ' перемещение с
      Set cmd = New ADODB.Command
      cmd.CommandText = "INSERT INTO HISTORY(" & _
      "stamp,Code , Item, Description, LOT_SN, EXP_DATE, UNIT_COST, QTY_ON_HAND, QTY, UOM, Status," & _
      "LOCATION , REF_NUM, ORD_NUM, USER_NAME, SITE_ID, PALLET, Container," & _
      "CUSTOM_FIELD1 , CUSTOM_FIELD2, CUSTOM_FIELD3, CUSTOM_FIELD4, CUSTOM_FIELD5, CUSTOM_FIELD6, CUSTOM_FIELD7, CUSTOM_FIELD8," & _
      "CUSTOM_FIELD9, CUSTOM_FIELD10, CUSTOM_FIELD11, CUSTOM_FIELD12, CUSTOM_FIELD13, CUSTOM_FIELD14, CUSTOM_FIELD15, CUSTOM_FIELD16 " & _
      ")VALUES( " & _
      "getdate(),3,'" & itemrs!code & "','" & itemrs!Description & "','" & strs!LOT_SN & "'," & MakeMSSQLDate(strs!exp_date) & "," & strs!Unit_COST & "," & MyRound2(strs!QTY_ON_HAND) & ",-" & MyRound2(strs!QTY_ON_HAND) & ",'" & strs!UOM & "',0," & _
      "'" & locfrs!code & "','" & strs!ref_num & "','" & strs!ord_num & "','sa',1,'" & palrs!code & "',''," & _
      "'-" & MyRound2(strs!custom_field1) & "','" & strs!custom_field2 & "','" & strs!custom_field3 & "','" & strs!custom_field4 & "','" & strs!custom_field5 & "','" & strs!custom_field6 & "','" & strs!custom_field7 & "','" & strs!custom_field8 & "'," & _
      "'" & strs!custom_field9 & "','" & strs!custom_field10 & "','" & strs!custom_field11 & "','" & strs!custom_field12 & "','" & strs!custom_field13 & "','" & strs!custom_field14 & "','" & strs!custom_field15 & "','" & strs!custom_field16 & "')"
  
  
      On Error Resume Next
      
      Set cmd.ActiveConnection = conn
      err.Clear
      cmd.Execute
      
      If err.Number <> 0 Then
        MsgBox err.Description
        RegisterMove = False
        Exit Function
      End If
      
      
      ' перемещение на
      
      Set cmd = New ADODB.Command
      cmd.CommandText = "INSERT INTO HISTORY(" & _
      "Stamp, Code , Item, Description, LOT_SN, EXP_DATE, UNIT_COST, QTY_ON_HAND, QTY, UOM, Status," & _
      "LOCATION , REF_NUM, ORD_NUM, USER_NAME, SITE_ID, PALLET, Container," & _
      "CUSTOM_FIELD1 , CUSTOM_FIELD2, CUSTOM_FIELD3, CUSTOM_FIELD4, CUSTOM_FIELD5, CUSTOM_FIELD6, CUSTOM_FIELD7, CUSTOM_FIELD8," & _
      "CUSTOM_FIELD9, CUSTOM_FIELD10, CUSTOM_FIELD11, CUSTOM_FIELD12, CUSTOM_FIELD13, CUSTOM_FIELD14, CUSTOM_FIELD15, CUSTOM_FIELD16 " & _
      ")VALUES( " & _
      "getdate(),4,'" & itemrs!code & "','" & itemrs!Description & "','" & strs!LOT_SN & "'," & MakeMSSQLDate(strs!exp_date) & "," & strs!Unit_COST & "," & 0 & "," & MyRound2(strs!QTY_ON_HAND) & ",'" & strs!UOM & "',0," & _
      "'" & loctrs!code & "','" & strs!ref_num & "','" & strs!ord_num & "','sa',1,'" & palrs!code & "',''," & _
      "'" & strs!custom_field1 & "','" & strs!custom_field2 & "','" & strs!custom_field3 & "','" & strs!custom_field4 & "','" & strs!custom_field5 & "','" & strs!custom_field6 & "','" & strs!custom_field7 & "','" & strs!custom_field8 & "'," & _
      "'" & strs!custom_field9 & "','" & strs!custom_field10 & "','" & strs!custom_field11 & "','" & strs!custom_field12 & "','" & strs!custom_field13 & "','" & strs!custom_field14 & "','" & strs!custom_field15 & "','" & strs!custom_field16 & "')"
  
      Set cmd.ActiveConnection = conn
      err.Clear
      cmd.Execute
      
      If err.Number <> 0 Then
        MsgBox err.Description
        RegisterMove = False
        Exit Function
      End If
    
    
      ' обновление стока
    
      Set cmd = New ADODB.Command
      cmd.CommandText = "Update stock set location_id=" & loctrs!id & " where id = " & strs!id
  
      
      Set cmd.ActiveConnection = conn
      err.Clear
      cmd.Execute
      
      If err.Number <> 0 Then
        MsgBox err.Description
        RegisterMove = False
        Exit Function
      End If
    
      MD.ISMoved = Boolean_Da
      MD.save
    End If
  End If
  
  
    
End Function

