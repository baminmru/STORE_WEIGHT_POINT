VERSION 5.00
Begin VB.Form frmAddPalette 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Добавить палеты"
   ClientHeight    =   2295
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFromOwner 
      Caption         =   "Чужая палета"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdPType 
      Caption         =   "..."
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtPltype 
      Enabled         =   0   'False
      Height          =   405
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2175
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
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   15
      Left            =   2400
      TabIndex        =   7
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Тип палеты"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Количество палет"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmAddPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 110

Option Explicit
'форма регистрация паллет

Private Sub CancelButton_Click()
  Unload Me
End Sub

'выбор типа
Private Sub cmdPType_Click()
  On Error Resume Next
  Dim id As String, brief As String
  If Manager.GetReferenceDialogEx2("ITTD_PLTYPE", id, brief) Then
    txtPltype.Tag = Left(id, 38)
    txtPltype = brief
  End If
End Sub

'запуск процесса регистрации
Private Sub OKButton_Click()
  Dim cnt As Long
  Dim pmin As Long, pmax As Long
  cnt = MyRound("0" & txtNum.Text)
  If cnt = 0 Then
    txtNum = "0"
    Exit Sub
  End If
  
  
  ' тип паллеты теперь обязателен
  If txtPltype.Tag = "" Then
   MsgBox "Задайте тип паллеты"
   Exit Sub
  End If
  
  
  Dim Item As ITTPL.Application
  Dim u As ITTPL_DEF
  Dim id As String, i As Long, s As String, sf As String
  Dim fn As ITTFN.Application
  Dim rs As ADODB.Recordset
  Dim conn As ADODB.Connection
  Dim cmd As ADODB.Command
  Set rs = Manager.ListInstances("", "ITTFN")
  Set fn = Manager.GetInstanceObject(rs!InstanceID)
  fn.LockResource False
  If fn.IsLocked <> LockSession Then
    MsgBox "Не удалось заблокировать нумератор, повторите попытку позже"
    Exit Sub
  End If
  
  Dim fnm As ITTFN.ITTFN_MAX
  If fn.ITTFN_MAX.Count = 0 Then
     With fn.ITTFN_MAX.Add
     .PalMaxNum = 0
     .save
     End With
   
  End If
  Set fnm = fn.ITTFN_MAX.Item(1)
  Manager.LockInstanceObject fn.id
  
  Set conn = GetCoreConn
  Set cmd = New ADODB.Command
  pmin = fnm.PalMaxNum + 1
  For i = 1 To cnt
    id = CreateGUID2
    Manager.NewInstance id, "ITTPL", "палета"
    Set Item = Manager.GetInstanceObject(id)
    Set u = Item.ITTPL_DEF.Add
    u.TheNumber = fnm.PalMaxNum + 1
    fnm.PalMaxNum = fnm.PalMaxNum + 1
    s = u.TheNumber
    s = Right("0000000000" & s, 6)
  
'    sf = Left(Right(s, 6), 2) & " " & Left(Right(s, 4), 2) & " " & Right(s, 2)
    
    u.code = s
    u.PalKode = code128(s)
    
    If chkFromOwner.Value = vbChecked Then
      u.PrivatePalet = Boolean_Da
    Else
      u.PrivatePalet = Boolean_Net
    End If
    
    
    Set u.Pltype = Item.FindRowObject("ITTD_PLTYPE", txtPltype.Tag)
    Dim PTYPE As ITTD_PLTYPE
    
    Set PTYPE = u.Pltype
    If PTYPE.TheCode = 0 Then
      cmd.CommandText = "insert into pallet(code,type,site_id) values('" & u.TheNumber & "','E',1)"
    Else
      cmd.CommandText = "insert into pallet(code,type,site_id) values('" & u.TheNumber & "','I',1)"
    End If
    
    ' Записываем код паллеты
    If conn.State <> adStateOpen Then
      conn.open
    End If
    Set cmd.ActiveConnection = conn
    cmd.Execute
    
    ' получаем ID
    Set rs = conn.Execute("select id from pallet where code='" & u.TheNumber & "'")
    u.CorePalette_ID = rs!id
    u.save
    fnm.save
    Item.Name = sf
    Item.save
    
    
    
    ' Записываем тип
    If PTYPE.TheCode = 0 Then
      cmd.CommandText = "insert into pallet_weight(code,TYPE,weight,date_weight) values('" & u.TheNumber & "',2,0,getdate())"
    Else
      cmd.CommandText = "insert into pallet_weight(code,TYPE,weight,date_weight) values('" & u.TheNumber & "',1,0,getdate())"
    End If
    
    If conn.State <> adStateOpen Then
      conn.open
    End If
    Set cmd.ActiveConnection = conn
    cmd.Execute
    
    
    Me.Caption = i
    DoEvents
    If i Mod 100 = 0 Then
       Manager.FreeAllInstanses
    End If
  
  Next
  pmax = fnm.PalMaxNum
  If cnt > 0 Then
    MsgBox "Регистрация паддонов завершена" & vbCrLf & "заняты номера с " & pmin & vbCrLf & "по " & pmax
  End If
  
  fn.UnLockResource
  Manager.UnLockInstanceObject fn.id
  Manager.FreeAllInstanses
  Unload Me
  
End Sub
