VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMakeOtbor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Провести отбор выморозки"
   ClientHeight    =   3900
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtOut 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3120
      Width           =   4575
   End
   Begin MSComctlLib.ProgressBar pbpl 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar pbCS 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtMaxSize 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "500"
      Top             =   480
      Width           =   3735
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
   Begin VB.Label Label4 
      Caption         =   "Результат:"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   5895
   End
   Begin VB.Label Label3 
      Caption         =   "Поддон в партии"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "Партия"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Минимальный вес для отбора (кг)"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmMakeOtbor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Me.Hide
End Sub

Private Sub OKButton_Click()
  MakeOtbor

  Me.Hide
End Sub


Private Sub MakeOtbor()
  Dim conn As ADODB.Connection
  Dim rs1 As ADODB.Recordset
  Dim rs2 As ADODB.Recordset
  Dim mx As Double
  Dim otbor As Double
  Dim out As Double
  
  mx = Val("0" & txtMaxSize)
  If mx < 0 Then mx = 0
  txtMaxSize = mx
  Set conn = Manager.GetCustomObjects("refref")
  Set rs1 = conn.Execute("select count(*) cnt from v_bami_vimorozka_rpt2 where otbor >=" & CLng(mx))
  pbCS.Min = 0
  pbCS.Max = rs1!cnt
  pbCS.Value = 0
  
  Set rs1 = conn.Execute("select checksum(item_code,Country,Factory,KILL_PLACE,IsBRAK,vetsved) cs,* from v_bami_vimorozka_rpt2 where otbor >=" & CLng(mx))
    
  If rs1.EOF Then Exit Sub
  
  
  
  While Not rs1.EOF
  
    otbor = rs1!otbor
    Set rs2 = conn.Execute("select  count(*) cnt from v_bami_stock where checksum(item_code,Country,Factory,KILL_PLACE,IsBRAK,vetsved) =" & rs1!cs)
    If rs2!cnt > 0 Then
      pbpl.Min = 0
      pbpl.Max = rs2!cnt
      pbpl.Value = 0
    
    
    Set rs2 = conn.Execute("select  * from v_bami_stock where status=0 and  checksum(item_code,Country,Factory,KILL_PLACE,IsBRAK,vetsved) =" & rs1!cs)
    If Not rs2 Is Nothing Then
    
    
      While Not rs2.EOF And otbor > 0
        If rs2!AtStock <= otbor Then
          Call conn.Execute("exec BLOCKPallet '" & rs2!pallet_code & "'")
          otbor = otbor - rs2!AtStock
          out = out + rs2!AtStock
          txtOut = out
          DoEvents
        End If
      
        pbpl.Value = pbpl.Value + 1
        rs2.MoveNext
      Wend
    
    End If
    End If
    DoEvents
    pbCS.Value = pbCS.Value + 1
    rs1.MoveNext
    
  Wend

'BLOCKPallet palcode
End Sub
