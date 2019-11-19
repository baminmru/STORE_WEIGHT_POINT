VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmWeighting 
   Caption         =   "Печать весов заявки №"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPoddon 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   2865
      Left            =   120
      Picture         =   "frmWeighting.frx":0000
      ScaleHeight     =   2805
      ScaleWidth      =   3075
      TabIndex        =   5
      Top             =   3720
      Width           =   3135
      Begin VB.Label lblWeight 
         BackColor       =   &H0000FF00&
         Caption         =   "000,000"
         Height          =   210
         Left            =   2250
         TabIndex        =   6
         Top             =   105
         Width           =   675
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Закончить взвешивание"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3435
      TabIndex        =   4
      Top             =   6090
      Width           =   6750
   End
   Begin VB.ListBox lstWeights 
      Height          =   1425
      Left            =   3405
      TabIndex        =   3
      Top             =   4185
      Width           =   6180
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Начать взвешивание"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3375
      TabIndex        =   1
      Top             =   2895
      Width           =   6705
   End
   Begin VB.CommandButton cmdDelWeight 
      Caption         =   "Х"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   0
      ToolTipText     =   "Удалить текущий вес"
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   390
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5520
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      Handshaking     =   2
   End
   Begin GridEX20.GridEX gr 
      Height          =   2475
      Left            =   30
      TabIndex        =   2
      Top             =   285
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   4366
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ItemCount       =   0
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   1
      Column(1)       =   "frmWeighting.frx":5B36
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmWeighting.frx":5B9A
      FormatStyle(2)  =   "frmWeighting.frx":5C7A
      FormatStyle(3)  =   "frmWeighting.frx":5DD6
      FormatStyle(4)  =   "frmWeighting.frx":5E86
      FormatStyle(5)  =   "frmWeighting.frx":5F3A
      FormatStyle(6)  =   "frmWeighting.frx":6012
      FormatStyle(7)  =   "frmWeighting.frx":60CA
      ImageCount      =   0
      PrinterProperties=   "frmWeighting.frx":60EA
   End
   Begin VB.Label Label6 
      Caption         =   "Поддон"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label lblstate 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ожидание "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4440
      TabIndex        =   12
      Top             =   3465
      Width           =   5670
   End
   Begin VB.Label Label1 
      Caption         =   "Состояние"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3345
      TabIndex        =   11
      Top             =   3510
      Width           =   1020
   End
   Begin VB.Label Label2 
      Caption         =   "Итого"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3420
      TabIndex        =   10
      Top             =   5655
      Width           =   1020
   End
   Begin VB.Label lblItog 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4395
      TabIndex        =   9
      Top             =   5685
      Width           =   2325
   End
   Begin VB.Label Label3 
      Caption         =   "Маршрут"
      Height          =   255
      Left            =   45
      TabIndex        =   8
      Top             =   0
      Width           =   1980
   End
   Begin VB.Label Label5 
      Caption         =   "Список оформленых весов"
      Height          =   315
      Left            =   3405
      TabIndex        =   7
      Top             =   3915
      Width           =   3285
   End
End
Attribute VB_Name = "frmWeighting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private pw As PEKVES.Support
Private port As String
Private psetup As String
Private InTimer As Boolean
Private nw As Single
Private wCnt As Long
Private itog As Double
Public Item As ITTIN.Application


Private curQRow As ITTIN.ITTIN_QLINE
Private StopWeighting As Boolean
Private wave As MTZMCI.WavePlayer


Private emu As Boolean

Public Sub Init()
  Form_Load
  form_Activate
End Sub

Private Sub cmdDelWeight_Click()
  If Not StopWeighting Then
    MsgBox "Сначала остановите процесс взвешивания"
    Exit Sub
  End If
  If lstWeights.ListCount > 0 And lstWeights.ListIndex >= 0 Then
    If MsgBox("Удалить вес?" & vbCrLf & lstWeights.Text, vbYesNo, "Внимание!") = vbYes Then
     
'      ' получаем номер
'      Dim nvs As NamedValues
'      Dim nv As NamedValue
'      Dim vl As Long
'      Dim s As String
'
'      Set nvs = New NamedValues
'      Set nv = nvs.Add("Dept", MyFil.id)
'      nv.ValueType = adGUID
'
'      vl = Year(Now)
'      Set nv = nvs.Add("Year", vl)
'      nv.ValueType = adInteger
'      vl = Month(Now)
'      Set nv = nvs.Add("Month", vl)
'      nv.ValueType = adInteger
'      'MsgBox (curQRow.ITTIN_PALET.Item(lstWeights.ListIndex + 1).ShCode)
'      vl = MyRound(Mid(curQRow.ITTIN_PALET.Item(lstWeights.ListIndex + 1).ShCode, 9, 5))
'      Set nv = nvs.Add("DropNum", vl)
'      nv.ValueType = adInteger
'      nv.ValueDirection = adParamOutput
'
'
'      Session.Exec "DropPPOCode", nvs
'      curQRow.ITTIN_PALET.Item(lstWeights.ListIndex + 1).Delete
'      curQRow.ITTIN_PALET.Refresh
'      InitWeighting
    End If
  End If
End Sub

Private Sub cmdStart_Click()
  If curQRow Is Nothing Then Exit Sub

  On Error Resume Next
  
  
  InitWeighting
  'Timer1.Enabled = True
  StopWeighting = False
  cmdDelWeight.Visible = False

  psetup = GetSetting("RBH", "ITTSETTINGS", "WSETUP", "9600,n,8,1")
  port = GetSetting("RBH", "ITTSETTINGS", "WPORT", 5)
  emu = Not (GetSetting("RBH", "ITTSETTINGS", "EMULATOR", "False") = "False")
  
  Dim ss As Single
  
  While Not StopWeighting
    DoEvents

    
    Dim ws As String
    Dim sss As Double
    If Not emu Then
      nw = GetWeight
    Else
      Dim res As Long
      res = MsgBox("Смоделировать новый вес?", vbOKCancel)
      If res = vbOK Then
        nw = Rnd(Second(Now)) * 30
      End If
      If res = vbCancel Then
        nw = 0
        StopWeighting = True
      End If
      If res = vbNo Then
        nw = 0
        sss = Timer
        While Timer < sss + 0.3
          DoEvents
        Wend
      End If
    End If
    
    'Debug.Print nw
    
    If nw > 0.15 Then
      Picture1.Picture = LoadPicture(App.Path & "\wload.bmp")
      lblstate = "Получен вес"
      If Not wave Is Nothing Then
        wave.OpenFile App.Path & "\new.wav"
        wave.Play
      End If
      
      lblWeight = Round(nw, 3)
      
      ' регистрируем вес
      DoEvents
      wCnt = wCnt + 1
      lstWeights.AddItem wCnt & " : " & lblWeight
      lstWeights.ListIndex = lstWeights.ListCount - 1
      itog = itog + Round(nw + 0.0001, 3)
      lblItog = "(" & wCnt & "): " & Round(itog, 0)
      
      ProcessWeight
      
      
      gr.RefreshRowIndex gr.Row
      
      If StopWeighting Then GoTo stopped
      
      'ждем пока не разгрузятся весы
      DoEvents
      WaitUnload
            
      If StopWeighting Then GoTo stopped
      
      
      
      
      lblstate = "Оформлено отправление"
      
      ' Весы пусты
      Picture1.Picture = LoadPicture(App.Path & "\wempty.bmp")
      lblWeight = "000,000"
      If Not wave Is Nothing Then
        wave.OpenFile App.Path & "\next.wav"
        wave.Play
      End If
      
    Else
       lblstate = "Ожидание веса"
    End If
    
  
  Wend
stopped:
  cmdDelWeight.Visible = True
  ProcessRow
  
End Sub

Private Sub Command1_Click()
StopWeighting = True
End Sub


Private Sub ProcessRow()
  ' состояния для типа:PEKZ Заявка
  ' "{A324A45F-2617-48C5-BC65-A334013A0401}" 'В пути
  ' "{E2A83D8A-BFB7-47D3-9C1D-DF2812BF9383}" 'Доставлена
  ' "{A8B3EF3B-6547-4BEC-A687-750FFC7C4E99}" 'Отвергнута
  ' "{2A121E0F-A8AA-4060-9093-7F3F8135D242}" 'Оформлена
  ' "{C30B597E-E523-4532-887A-9B9DD57EF06C}" 'Оформляется
  ' "{C450C343-4D0E-4010-AEF1-18C4627E6D80}" 'Принято на склад
  ' "{F4C3A104-B225-450F-84C9-4226F9E4F57B}" 'Согласована
  ' "{78A69448-43EF-436B-82ED-9ABC3CBB56E3}" 'Утеряна
  If MsgBox("Весь груз по заявке принят?", vbYesNo, "Завершение взвешивания") = vbYes Then
    If RoleDocCanSwitchStatus(Item) Then
      If BeforeChangeStatus(Item, "{C450C343-4D0E-4010-AEF1-18C4627E6D80}") Then
        Item.StatusID = "{C450C343-4D0E-4010-AEF1-18C4627E6D80}"
        Unload Me
      End If
    End If
  Else
   lblstate = "Взвешивание остановлено. Нажмите кнопку ""Начать взвешивание"""
  End If
  
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
Private Sub form_Activate()
Dim ff As Long
  Dim s As String
  On Error Resume Next
  If Item.IsLocked = ExternalLockSession Or Item.IsLocked = ExternalLockPermanent Then
    MsgBox "Документ заблокирован другим пользователем"
    Unload Me
    Exit Sub
  End If

  
End Sub





Private Sub Form_Load()
  On Error Resume Next
  If GetSetting("RBH", "ITTSETTINGS", "SOUND", "False") <> "False" Then
    Set wave = New MTZMCI.WavePlayer
    wave.OpenDevice
  End If
  
  
  itog = 0
  wCnt = 0
  lstWeights.Clear
  StopWeighting = True
  emu = Not (GetSetting("RBH", "ITTSETTINGS", "EMULATOR", "False") = "False")
  psetup = GetSetting("RBH", "ITTSETTINGS", "WSETUP", "4800,e,8,1")
  port = GetSetting("RBH", "ITTSETTINGS", "WPORT", 1)

  If Not emu Then
    If MSComm1.PortOpen Then
      MSComm1.PortOpen = False
    End If
'    If wtype = 4 Then
      
      MSComm1.Handshaking = comNone
      MSComm1.DTREnable = False
      MSComm1.EOFEnable = False
      
'    Else
'      MSComm1.Handshaking = comRTS
'      MSComm1.DTREnable = True
'      MSComm1.EOFEnable = False
'    End If
    
    
    MSComm1.Settings = psetup
    MSComm1.CommPort = port
    MSComm1.PortOpen = True
    Me.Caption = "Port " & port & " - " & psetup
  End If
  
  On Error Resume Next
  lblItog = "(" & wCnt & "): " & Round(itog, 0)
  Item.LockResource
  gr.ItemCount = 0
  Item.ITTIN_QLINE.PrepareGrid gr
  gr.ItemCount = Item.ITTIN_QLINE.Count

 Exit Sub
bye:
  MsgBox err.Description, , "Взвешивание"
End Sub
Private Sub InitWeighting()
  itog = 0
  lstWeights.Clear
  Dim i As Long
  For i = 1 To curQRow.ITTIN_PALET.Count
    itog = itog + curQRow.ITTIN_PALET.Item(i).PalWeight
    lstWeights.AddItem i & ":" & curQRow.ITTIN_PALET.Item(i).PalWeight
  Next
  
  
  wCnt = curQRow.ITTIN_PALET.Count
  lblItog = "(" & wCnt & "): " & Round(itog + 0.001, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  StopWeighting = True
  On Error Resume Next
  If Not emu Then
    MSComm1.PortOpen = False
  End If
  Item.UnLockResource
End Sub
Private Sub gr_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
On Error Resume Next
If gr.ItemCount = 0 Then
  Exit Sub
End If

If gr.Row > 0 Then
 If gr.RowIndex(gr.Row) > 0 Then
  If LastRow <> gr.Row Then
    
    Dim bm
    bm = gr.RowBookmark(gr.RowIndex(gr.Row))
    
    
    
    If Not curQRow Is Nothing Then
     ' пересчитать вес и количество для строки
     
     StopWeighting = True
     If wCnt <> 0 And itog <> 0 Then
      curQRow.CurValue = Round(itog, 0)
      curQRow.save
     End If
     
     lblstate = "Взвешивание остановлено. Нажмите кнопку ""Начать взвешивание"""
    Else
      lblstate = "Нажмите кнопку ""Начать взвешивание"""
    End If
    
    
    
    Set curQRow = Item.FindRowObject(Right(bm, Len(bm) - 38), Left(bm, 38))
    InitWeighting
    cmdDelWeight.Visible = True
    
  End If
 End If
End If
End Sub

Private Sub gr_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
  On Error Resume Next
  Item.ITTIN_QLINE.LoadRow gr, RowIndex, Bookmark, Values
End Sub



Private Sub WaitUnload()
  On Error Resume Next
  Dim tt As Single
  Picture1.Picture = LoadPicture(App.Path & "\wunload.bmp")
  lblstate = "Снимите отправление с весов"
  If Not wave Is Nothing Then
    wave.OpenFile App.Path & "\unload.wav"
    wave.Play False
  End If
  Dim nw As Double
  Dim start As Double
  If Not emu Then
  'Ждем пока весы не оттдадут нам пустое значение
again:
    nw = 1
    While nw > 0
      DoEvents
      nw = GetWeight
      If StopWeighting Then
        If Not wave Is Nothing Then
          wave.StopPlaying
        End If
        
        Exit Sub
      End If
    Wend
    start = Timer
  
    ' это дело должно сохраняться 1 сек
    While Timer < start + 1
      DoEvents
      nw = GetWeight
      If nw <> 0 Then GoTo again
      If StopWeighting Then
        If Not wave Is Nothing Then
          wave.StopPlaying
        End If
        Exit Sub
      End If
    Wend
    If Not wave Is Nothing Then
      wave.StopPlaying
    End If
  Else
    start = Timer
    While Timer < start + 2
      DoEvents
    Wend
  
  End If
End Sub



Private Sub ProcessWeight()

  On Error Resume Next
  
  If curQRow Is Nothing Then Exit Sub
  ' записываем новое отправление
  
  Dim wl As ITTIN_PALET
    Dim ppo_num As Long
    Set wl = curQRow.ITTIN_PALET.Add
    With wl
        .PalWeight = Round(nw + 0.0001, 1)
        .TheNumber = txtPoddon.Text
        ' Заносим в текущий маршрут
        .save
    End With
  Exit Sub
bye:
  MsgBox err.Description, , "Ошибка печати"
End Sub




Public Function GetWeight() As Double
  Dim gw As Double
  
  gw = GetWeight4
    
  If gw < 0.4 Then
    gw = 0
  End If
  
  GetWeight = gw
  
End Function

'Public Function GetWeight1() As Double
'  On Error Resume Next
'    Dim ws As String
'    Dim ch As String
'    Dim start As Single
'    GetWeight1 = 0
'    MSComm1.output = Chr(5)
'    start = Timer   ' Set start time.
'    Do While Timer < start + 0.2
'    Loop
'    start = Timer   ' Set start time.
'    Do While Timer < start + 0.5
'       If MSComm1.InBufferCount > 17 Then GoTo answer
'    Loop
'    Debug.Print MSComm1.input
'    GetWeight1 = 0
'    Exit Function
'
'answer:
'    ws = MSComm1.input
'    GetWeight1 = MyRound(Mid(ws, 7))
'End Function


'Public Function GetWeight3() As Double
'  On Error Resume Next
'    Dim ws As String
'    Dim ch As String
'    Dim start As Single
'    GetWeight3 = 0
'       If MSComm1.InBufferCount > 0 Then GoTo answer
'    Exit Function
'answer:
'    Dim s1 As String, i1 As Long, i2 As Long
'    Dim s As String
'    start = Timer   ' Set start time.
'    Do While Timer < start + 0.3
'    Loop
'    s = MSComm1.input
'    s1 = ""
'
'    i1 = InStr(1, s, "Weight :", vbTextCompare)
'    If i1 > 0 Then
'      i2 = InStr(i1, s, "kg", vbTextCompare)
'      s1 = Trim(Mid(s, i1 + 8, i2 - i1 - 8))
'    End If
'    GetWeight3 = MyRound(s1)
'End Function


'Public Function GetWeight2() As Double
'    Dim ws As String
'    Dim ch As String
'    Dim start As Single
'again:
'    GetWeight2 = 0
'    MSComm1.output = Chr(5)
'    start = Timer   ' Set start time.
'    Do While Timer < start + 0.2
'      'If MSComm1.InBufferCount > 0 Then GoTo answer
'    Loop
'    start = Timer   ' Set start time.
'    Do While Timer < start + 0.5
'       If MSComm1.InBufferCount > 0 Then GoTo answer
'    Loop
'answer:
'    ch = MSComm1.input
'    If Len(ch) >= 1 Then
'      If Asc(Left(ch, 1)) = 6 Then
'      GoTo found
'      End If
'    End If
'    'MsgBox "Нет весов"
'    Exit Function
'found:
'
'    MSComm1.output = Chr(17)
'    start = Timer
'    Do While Timer < start + 0.2
'      'If MSComm1.InBufferCount > 0 Then GoTo answer2
'    Loop
'    start = Timer
'     Do While Timer < start + 0.5
'       If MSComm1.InBufferCount > 0 Then GoTo answer2
'    Loop
'answer2:
'      ws = MSComm1.input
'  If ws <> "" Then
'    Debug.Print ws
'
'    If UCase(Mid(ws, 3, 1)) = "U" Then
'     GoTo again
'    End If
'    If UCase(Mid(ws, 3, 1)) <> "S" Then
'      GetWeight2 = 0
'      Exit Function
'    End If
'
'    If UCase(Mid(ws, 4, 1)) = "F" Then
'      GetWeight2 = 0
'      Exit Function
'    End If
'    GetWeight2 = MyRound(Replace(Replace(Replace(Replace(Mid(ws, 5), " ", "0"), ",", "."), "kg", "  "), "lb", "  "))
'  Else
'    GetWeight2 = 0
'  End If
'End Function


Public Function GetWeight4() As Double
  On Error Resume Next
    Dim ws As String
    Dim ch As String
    Dim start As Single
    Dim ws1 As String
    Dim ws2 As String
    GetWeight4 = 0
    
    MSComm1.output = Chr(68)
    start = Timer   ' Set start time.
    Do While Timer < start + 0.2
    Loop
    
    If MSComm1.InBufferCount > 0 Then GoTo answer_s1
    start = Timer   ' Set start time
    Do While Timer < start + 0.5
       If MSComm1.InBufferCount > 0 Then GoTo answer_s1
    Loop
    
    GetWeight4 = 0  ' не дождались ответа
    Exit Function
    
answer_s1:
    
    ws = MSComm1.Input
    ' первый раз вес стабилен
    If Asc(Mid(ws, 1, 1)) >= 128 Then
    
      ''''''''''''''''''''''''''''''''''''
      'ЗАДЕРЖКА !!!
      '
      ' ждем чтобы исключить дребезг
      start = Timer   ' Set start time.
      Do While Timer < start + 0.3
      Loop
      
      ' спрашиваем еще раз
      MSComm1.output = Chr(68)
      
      
      start = Timer   ' Set start time.
      Do While Timer < start + 0.2
      Loop
      
      If MSComm1.InBufferCount > 0 Then GoTo answer_s2
      start = Timer   ' Set start time
      Do While Timer < start + 0.5
         If MSComm1.InBufferCount > 0 Then GoTo answer_s2
      Loop
      
    End If
    
    GetWeight4 = 0 ' нет второго ответа
    Exit Function
    
answer_s2:

    ws = MSComm1.Input
    
    ' второй раз вес тоже стабилен
    If Asc(Mid(ws, 1, 1)) >= 128 Then
      MSComm1.output = Chr(69)
      start = Timer   ' Set start time.
      Do While Timer < start + 0.2
      Loop
      If MSComm1.InBufferCount > 0 Then GoTo answer_w1
      start = Timer   ' Set start time
      Do While Timer < start + 0.5
       If MSComm1.InBufferCount > 0 Then GoTo answer_w1
      Loop
    End If
    
    GetWeight4 = 0 ' вес не стабилен, или нет ответа
    Exit Function
    
answer_w1:

    ' прочли показания веса
    ws1 = MSComm1.Input
    
    
    ''''''''''''''''''''''''''''''''''''
    'ЗАДЕРЖКА !!!
    '
    ' ждем чтобы исключить дребезг
    start = Timer   ' Set start time.
    Do While Timer < start + 0.3
    Loop
    
    ' спрашиваем вес еще раз
    MSComm1.output = Chr(69)
    start = Timer   ' Set start time.
    Do While Timer < start + 0.2
    Loop
    
    If MSComm1.InBufferCount > 0 Then GoTo answer_w2
    start = Timer   ' Set start time
    Do While Timer < start + 0.5
       If MSComm1.InBufferCount > 0 Then GoTo answer_w2
    Loop
    
    GetWeight4 = 0 '  нет ответа
    Exit Function
      
answer_w2:
    ws = MSComm1.Input
  
    If ws1 = ws Then
      GetWeight4 = (Asc(Mid(ws, 2, 1)) * 256 + Asc(Mid(ws, 1, 1))) / 100
    Else
      GetWeight4 = 0 ' вес не стабилен, отличаются показания
    End If
  
End Function

'Public Function GetWeight4() As Double
'  On Error Resume Next
'    Dim ws As String
'    Dim ch As String
'    Dim start As Single
'    GetWeight4 = PreGW4
'
'    MSComm1.output = Chr(68)
'    start = Timer   ' Set start time.
'    Do While Timer < start + 0.2
'    Loop
'
'    If MSComm1.InBufferCount > 0 Then GoTo answer1
'    Do While Timer < start + 0.5
'       If MSComm1.InBufferCount > 0 Then GoTo answer1
'    Loop
'    GetWeight4 = 0 'PreGW4
'    Exit Function
'
'answer1:
'    ws = MSComm1.input
'    If Asc(Mid(ws, 1, 1)) >= 128 Then
'
'      MSComm1.output = Chr(69)
'      start = Timer   ' Set start time.
'      Do While Timer < start + 0.2
'      Loop
'
'      If MSComm1.InBufferCount > 0 Then GoTo answer
'
'      Do While Timer < start + 0.5
'       If MSComm1.InBufferCount > 0 Then GoTo answer
'      Loop
'      GetWeight4 = 0 ' PreGW4
'    Else
'      GetWeight4 = 0 'PreGW4
'    End If
'    Exit Function
'
'answer:
'    ws = MSComm1.input
'    PreGW4 = (Asc(Mid(ws, 2, 1)) * 256 + Asc(Mid(ws, 1, 1))) / 100
'    GetWeight4 = PreGW4
'End Function




