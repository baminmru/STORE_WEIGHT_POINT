VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmAssemblyWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Объединить два поддона"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   Icon            =   "frmAssemblyWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Куда переложить"
      Height          =   3735
      Left            =   3240
      TabIndex        =   1
      Top             =   3840
      Width           =   6015
      Begin VB.CommandButton cmd6FindCell 
         Caption         =   "..."
         Height          =   375
         Left            =   5280
         TabIndex        =   40
         ToolTipText     =   "Поиск ячейки"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txt4NewPlace 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1680
         Width           =   5055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "x"
         Height          =   375
         Left            =   5280
         TabIndex        =   22
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txt5PNum 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label17 
         Caption         =   "Место в буферной зоне"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Номер поддона"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Второй поддон"
      Height          =   5055
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txt4MainCell 
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CommandButton cmd4ClearNum 
         Caption         =   "X"
         Height          =   375
         Left            =   2400
         TabIndex        =   42
         ToolTipText     =   "ввести номер еще раз"
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmd4ClearW 
         Caption         =   "X"
         Height          =   375
         Left            =   2400
         TabIndex        =   31
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txt4GoodWeight 
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txt4FullWeight 
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txt4Quantity 
         Height          =   405
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txt4PWeight 
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txt4PNum 
         Height          =   405
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txt4PackageWeight 
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox txt4Netto 
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Ячейка основного хранения"
         Height          =   375
         Left            =   2880
         TabIndex        =   44
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label Label13 
         Caption         =   "Вес товара брутто"
         Height          =   375
         Left            =   2880
         TabIndex        =   38
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label12 
         Caption         =   "Вес груза с поддоном"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label11 
         Caption         =   "Количество коробов"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Вес поддона КГ."
         Height          =   255
         Left            =   2880
         TabIndex        =   35
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Поддон №"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Вес одной упаковки КГ."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Вес товара НЕТТО"
         Height          =   255
         Left            =   2880
         TabIndex        =   32
         Top             =   3000
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Первый подон"
      Height          =   4335
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton cmd3ClearNum 
         Caption         =   "X"
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         ToolTipText     =   "ввести номер еще раз"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txt3PNum 
         Height          =   405
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txt3PWeight 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txt3FullWeight 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txt3GoodWeight 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton cmd3ClearW 
         Caption         =   "X"
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         ToolTipText     =   "Получить вес с  весов"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txt3Quantity 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox txt3PackageWeight 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox txt3MainCell 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3360
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Поддон №"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "Вес груза с поддоном"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Вес груза НЕТТО"
         Height          =   375
         Left            =   2760
         TabIndex        =   18
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label18 
         Caption         =   "Количество коробов"
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label19 
         Caption         =   "Вес поддона"
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label20 
         Caption         =   "Вес одной упаковки"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Ячейка основного хранения"
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   3000
         Width           =   3015
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   720
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отменить"
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Далее"
      Height          =   615
      Left            =   6960
      TabIndex        =   3
      Top             =   6240
      Width           =   1815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      Handshaking     =   2
   End
End
Attribute VB_Name = "frmAssemblyWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As ADODB.Connection
Private StepNo As Integer
Private StopWeighting As Boolean
Private wave As MTZMCI.WavePlayer
Private emu As Boolean
Private port As String
Private psetup As String

Private poddon As ITTPL_DEF
Private item_id As String
Private country As String
Private factory As String
Private killplace As String
Private qry As String
Private edizm As String
Private exp_date As String
Private made_date As String
Private status As String
Private brak As String
Private partia As String

Private Poddon2 As ITTPL_DEF
Private item_id2 As String
Private Country2 As String
Private Factory2 As String
Private killplace2 As String
Private qry2 As String
Private edizm2 As String
Private exp_date2 As String
Private made_date2 As String
Private status2 As String
Private brak2 As String
Private partia2 As String


Private Sub MyBeep(ByVal BeepType As String)
      If Not wave Is Nothing Then
        On Error Resume Next
        wave.OpenFile App.Path & "\" & BeepType & ".wav"
        wave.Play
      End If
End Sub


Private Function CheckPoddon2() As Boolean
On Error Resume Next
  Dim result As Boolean
  result = True
  CheckPoddon2 = False
  If txt4PNum <> "" Then
    If Len(txt4PNum) = 6 Then
      Set Poddon2 = Nothing
      Set Poddon2 = FindPoddon(txt4PNum)
      If Not Poddon2 Is Nothing Then
        If Poddon2.Application.StatusID = "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" Then
          Dim conn As ADODB.Connection
          Set conn = Manager.GetCustomObjects("refref")
          If conn.State <> adStateOpen Then
            conn.Open
          End If
          
          Dim rs As ADODB.Recordset
          'poddon.CurrentGood
          
          Set rs = conn.Execute("select * from stock where PALLET_STATUS is null and pallet_id=" & Poddon2.CorePalette_ID)
          If rs.EOF Then
            MsgBox "Номер паддона: " & txt4PNum & "  не обнаружен в базе CORE IMS"
            result = False
          Else
            MyBeep "Nomer"
            Dim lid  As String
            lid = "" & rs!location_id
            
            item_id2 = rs!item_id
            Country2 = rs!custom_field6
            Factory2 = rs!custom_field4
            killplace2 = rs!custom_field11
            edizm2 = rs!UOM
            qry2 = rs!ORD_NUM
            status2 = rs!status
            brak2 = rs!custom_field12
            exp_date2 = rs!exp_date
            made_date2 = rs!custom_feild5
            partia2 = rs!lot_sn
            
            
            If lid <> "" Then
              Set rs = conn.Execute("select * from location where id=" & lid)
              txt4MainCell = rs!Code
              txt4MainCell.Tag = rs!id
            End If
            
            
            
            
            If item_id2 <> item_id Then
                MsgBox "Товар на поддонах не совпадает"
                Exit Function
            End If
            
            If Country2 <> country Then
                MsgBox "Страна не совпадает "
                Exit Function
            End If
            
            If Factory2 <> factory Then
                MsgBox "Производитель не совпадает "
                Exit Function
            End If
            
            If killplace2 <> killplace Then
                MsgBox "Бойня не совпадает "
                Exit Function
            End If
            
            txt4PWeight = Poddon2.Weight
            
          End If
        Else
          MsgBox "Состояние паддона: " & txt4PNum & "  установлено неверно (" & Poddon2.Application.StatusName & ")"
          result = False
        End If
      Else
        MsgBox "Номер паддона: " & txt4PNum & "  не зарегистрирован"
        result = False
      End If
    End If
  End If
  CheckPoddon2 = result
End Function

Private Sub AdjFrame(f As Frame)
On Error Resume Next
  f.Top = 0
  f.Left = 5 * Screen.TwipsPerPixelX
  f.Width = Me.ScaleWidth - 10 * Screen.TwipsPerPixelX
  f.Height = Me.ScaleHeight - cmdNext.Height - 5 * Screen.TwipsPerPixelY
End Sub

Private Sub cmd3ClearNum_Click()
  txt3PNum = ""
End Sub

Private Sub cmd3ClearW_Click()
  txt3FullWeight = 0
End Sub

Private Sub cmd4ClearW_Click()
txt4FullWeight = "0"
End Sub

Private Sub cmd6FindCell_Click()
  Dim f As frmGetCell
  Set f = New frmGetCell
  f.itemid = item_id
  On Error Resume Next
  f.country = ""
  f.country = country
  f.factory = ""
  f.factory = factory
  f.killplace = ""
  f.killplace = killplace
  err.Clear
  
  f.Show vbModal
  If f.OK Then
    txt4NewPlace = f.OutCode
    txt4NewPlace.Tag = f.OUtID
  End If
  Unload f
  Set f = Nothing
End Sub

Private Sub cmdCancel_Click()
  StepNo = 4
  ProcessStatus
End Sub

Private Sub cmdNext_Click()
 If StepNo = 1 Then
    If Not After1 Then
      Exit Sub
    End If
  End If
  
  If StepNo = 2 Then
    If Not After2 Then
      Exit Sub
    End If
  End If

  If StepNo = 3 Then
    If Not After3 Then
      Exit Sub
    End If
  End If
  
  StepNo = StepNo + 1

  ProcessStatus
End Sub





Private Sub Form_Load()

  On Error Resume Next
    emu = Not (GetSetting("RBH", "ITTSETTINGS", "EMULATOR", "False") = "False")
    psetup = GetSetting("RBH", "ITTSETTINGS", "WSETUP", "4800,e,8,1")
    port = GetSetting("RBH", "ITTSETTINGS", "WPORT", 1)
    
    StepNo = 1
    ProcessStatus
    
    Set conn = Manager.GetCustomObjects("refref")
    If GetSetting("RBH", "ITTSETTINGS", "SOUND", "False") <> "False" Then
      Set wave = New MTZMCI.WavePlayer
      wave.OpenDevice
    End If
    
    If Not emu Then
      If MSComm1.PortOpen Then
        MSComm1.PortOpen = False
      End If
        
      MSComm1.Handshaking = comNone
      MSComm1.DTREnable = False
      MSComm1.EOFEnable = False
        
      MSComm1.Settings = psetup
      MSComm1.CommPort = port
      MSComm1.PortOpen = True
    End If
    
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
  If UnloadMode <> 1 Then
    Cancel = -1
  Else
    wave.StopPlaying
    Set wave = Nothing
    Timer1.Enabled = False
    If MSComm1.PortOpen Then
      MSComm1.PortOpen = False
    End If
  End If
End Sub

  

Private Sub Timer1_Timer()
On Error Resume Next
  Dim w As Double
  If StepNo = 1 Then
    If txt3PNum = "" Then
      txt3PNum.SetFocus
    End If
    If txt3FullWeight = "0" Or Not IsNumeric(txt3FullWeight) Then
      w = GetWeight
      If w > 0 Then
        txt3FullWeight = Round(w + 0.001, 1)
        MyBeep "Gruz"
      End If
    End If
  End If
  
  If StepNo = 2 Then
    If txt4PNum = "" Then
      txt4PNum.SetFocus
    End If
    If txt4FullWeight = "0" Or Not IsNumeric(txt4FullWeight) Then
      w = GetWeight
      If w > 0 Then
        txt4FullWeight = Round(w + 0.001, 1)
        MyBeep "Gruz"
      End If
    End If
  End If
  
  
  If StepNo = 3 Then
    
    If txt5PNum = "" Then
      txt5PNum.SetFocus
    End If
  End If

End Sub

Private Sub txt3PNum_Change()
On Error Resume Next
  
  If CheckPoddon Then
    txt3PWeight = poddon.Weight
    txt3FullWeight = poddon.CurrentWeightBrutto
    txt3Quantity = poddon.CaliberQuantity
    txt3PackageWeight = poddon.PackageWeight / poddon.CaliberQuantity
  End If
End Sub

Private Function CheckPoddon() As Boolean
On Error Resume Next
  Dim result As Boolean
  result = True
  If txt3PNum <> "" Then
    If Len(txt3PNum) = 6 Then
      Set poddon = Nothing
      Set poddon = FindPoddon(txt3PNum)
      If Not poddon Is Nothing Then
        If poddon.Application.StatusID = "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" Then
          Dim conn As ADODB.Connection
          Set conn = Manager.GetCustomObjects("refref")
          If conn.State <> adStateOpen Then
            conn.Open
          End If
          
          Dim rs As ADODB.Recordset
          'poddon.CurrentGood
          
          Set rs = conn.Execute("select * from stock where PALLET_STATUS is null and pallet_id=" & poddon.CorePalette_ID)
          If rs.EOF Then
            MsgBox "Номер паддона: " & txt3PNum & "  не обнаружен в базе CORE IMS"
            result = False
          Else
            MyBeep "Nomer"
            Dim lid  As String
            lid = "" & rs!location_id
            
            item_id = rs!item_id
            country = rs!custom_field6
            factory = rs!custom_field4
            killplace = rs!custom_field11
            edizm = rs!UOM
            qry = rs!ORD_NUM
            status = rs!status
            brak = rs!custom_field12
            exp_date = rs!exp_date
            made_date = rs!custom_feild5
            partia = rs!lot_sn
            
            
            If lid <> "" Then
              Set rs = conn.Execute("select * from location where id=" & lid)
              txt3MainCell = rs!Code
              txt3MainCell.Tag = rs!id
            End If
            txt3PWeight = poddon.Weight
            
            
          End If
        Else
          MsgBox "Состояние паддона: " & txt3PNum & "  установлено неверно (" & poddon.Application.StatusName & ")"
          result = False
        End If
      Else
        MsgBox "Номер паддона: " & txt3PNum & "  не зарегистрирован"
        result = False
      End If
    End If
  End If
  CheckPoddon = result
End Function

Private Sub txt3Poddon_Change()
  CheckPoddon2
End Sub


Private Sub ProcessStatus()
  Frame1.Visible = False
  Frame2.Visible = False
  Frame3.Visible = False
  
  cmdNext.Caption = "Далее"
  cmdCancel.Caption = "Отменить"
  cmdCancel.Visible = True

  Select Case StepNo
  Case 1
  
    'Before1
    Frame1.Visible = True
    AdjFrame Frame1
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
  Case 2
    'Before2
    Frame2.Visible = True
    AdjFrame Frame2
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
  Case 3
    Befor3
    Frame3.Visible = True
    AdjFrame Frame3
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
  
  Case 4
    Unload Me
  End Select
  
  
End Sub


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
      GetWeight4 = (Asc(Mid(ws, 2, 1)) * 256 + Asc(Mid(ws, 1, 1))) / 10
    Else
      GetWeight4 = 0 ' вес не стабилен, отличаются показания
    End If
  
End Function

Private Function GetWeight() As Double
  If emu Then
    If StepNo = 2 Then
      GetWeight = Rnd(Second(Now)) * 20
    Else
      GetWeight = 20 + Rnd(Second(Now)) * 1000
    End If
  Else
    GetWeight = GetWeight4
  End If
End Function

Private Sub SetBtnPos(cmd As CommandButton, ByVal pos As Integer)
  On Error Resume Next
  cmd.Left = (Me.ScaleWidth) / 4 * (pos - 1)
End Sub


Private Sub txt3FullWeight_Change()
On Error Resume Next
txt3GoodWeight = MyRound(txt3FullWeight) - MyRound(txt3PWeight) - (MyRound(txt3PackageWeight) * MyRound(txt3Quantity))
End Sub

Private Sub txt3PackageWeight_Change()
On Error Resume Next
txt3GoodWeight = MyRound(txt3FullWeight) - MyRound(txt3PWeight) - (MyRound(txt3PackageWeight) * MyRound(txt3Quantity))
End Sub

Private Sub txt3Quantity_Change()
On Error Resume Next
txt3GoodWeight = MyRound(txt3FullWeight) - MyRound(txt3PWeight) - (MyRound(txt3PackageWeight) * MyRound(txt3Quantity))
End Sub

'Private Sub txt4FullWeight_Change()
'  On Error Resume Next
'  txt4GoodWeight = MyRound(txt4FullWeight) - (MyRound(Txt4PackageWeight) * MyRound(txt4Quantity))
'
'End Sub


Private Sub Befor3()
  
End Sub


Private Sub txt4FullWeight_Change()
  On Error Resume Next
  txt4GoodWeight = MyRound(txt4FullWeight) - MyRound(txt4PWeight)
End Sub

Private Function After1() As Boolean
  Dim result As Boolean
  After1 = True
  If MyRound(txt3GoodWeight) <= 0 Then
    MsgBox "Не задан вес груза на первом поддоне"
    After1 = False
    Exit Function
  End If
End Function


Private Function After2() As Boolean
Dim result As Boolean
  After2 = True
  If txt3PNum = txt4PNum Then
    MsgBox "Поддоны не могут совпадать"
    After2 = False
    Exit Function
  End If
  If MyRound(txt4Netto) <= 0 Then
    MsgBox "Не задан вес груза на втором поддоне"
    After2 = False
    Exit Function
  End If
  If MyRound(txt3FullWeight) + MyRound(txt4Netto) > 1000 Then
    MsgBox "Суммарный вес превышает 1000 кг."
    After2 = False
    Exit Function
  End If
End Function

Private Function After3() As Boolean
  Dim result As Boolean
'  If MyRound(txt4Netto) <= 0 Then
'    MsgBox "Не задан вес груза на новом поддоне"
'    After3 = False
'    Exit Function
'  End If
'
'  If MyRound(txt4Netto) > MyRound(txt3GoodWeight) Then
'    MsgBox "Вес после переразмещения превышает исходный"
'    After3 = False
'    Exit Function
'  End If
'
'  If MyRound(txt4Quantity) > MyRound(txt3Quantity) Then
'    MsgBox "Количество коробов переразмещения превышает исходный"
'    After3 = False
'    Exit Function
'  End If
  
  If MsgBox("Зарегистрировать объединение поддонов?", vbYesNo) = vbYes Then
  
  
'    If txt3PNum <> txt5PNum Then
        ClosePoddon FindPoddon(txt3PNum)
'    End If
'    If txt4PNum <> txt4PNum Then
        ClosePoddon FindPoddon(txt4PNum)
'    End If
    
    AssemblyPoddon FindPoddon(txt5PNum)
   
    PrintSticker FindPoddon(txt5PNum)
  
    After3 = True
  End If
  

End Function

Private Sub txt4GoodWeight_Change()
  txt4Netto = MyRound(txt4GoodWeight) - (MyRound(txt4PackageWeight) * MyRound(txt4Quantity))
End Sub

Private Sub Txt4PackageWeight_Change()
txt4Netto = MyRound(txt4GoodWeight) - (MyRound(txt4PackageWeight) * MyRound(txt4Quantity))
End Sub

Private Sub txt4PNum_Change()
On Error Resume Next
  
  If CheckPoddon2 Then
    txt4PWeight = Poddon2.Weight
    txt4FullWeight = Poddon2.CurrentWeightBrutto
    txt4Quantity = Poddon2.CaliberQuantity
    txt4PackageWeight = Poddon2.PackageWeight / Poddon2.CaliberQuantity
  End If

End Sub

Private Sub txt4PWeight_Change()
On Error Resume Next
  txt4GoodWeight = MyRound(txt4FullWeight) - MyRound(txt4PWeight)
End Sub

Private Sub txt4Quantity_Change()
  txt4Netto = MyRound(txt4GoodWeight) - (MyRound(txt4PackageWeight) * MyRound(txt4Quantity))
End Sub


Private Sub AssemblyPoddon(poddon As ITTPL_DEF)
' состояния для типа:ITTPL Палетта
' "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}" 'Взвешена
' "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" 'На складе с грузом
' "{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}" 'Отправлена с грузом
' "{E9BFB749-A606-4DEF-A429-07D636F108C6}" 'Пустая
' "{588C5203-1E59-408E-92A1-B3DFED8C19FA}" 'Списана


On Error Resume Next
  Dim conn As ADODB.Connection
  Set conn = Manager.GetCustomObjects("refref")
  Dim cmd As ADODB.Command
  Dim rs As ADODB.Recordset
  Dim rsitem As ADODB.Recordset
  
  Dim rlID As String
  Dim palID As String
  Dim oid As String
 
  palID = poddon.CorePalette_ID
  
  poddon.Application.StatusID = "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}"
  
  ' запрашиваем свободное место в буферной зоне
  Dim bzrs As ADODB.Recordset
  Dim loccode As ADODB.Recordset
  Dim bzid As String
  Set conn = Manager.GetCustomObjects("refref")
  If conn.State <> adStateOpen Then
    conn.Open
  End If
  
  
  Dim s As String
  Dim netto As Double
  
  
  
  
  
  netto = MyRound(txt4Netto) + MyRound(txt3GoodWeight)
  

  poddon.CaliberQuantity = MyRound(txt4Quantity) + MyRound(txt3Quantity)
  poddon.PackageWeight = MyRound(txt4PackageWeight) * (MyRound(txt4Quantity) + MyRound(txt3Quantity))
  poddon.CurrentWeightBrutto = MyRound(txt4FullWeight) + MyRound(txt3FullWeight)
  poddon.CurrentPosition = txt4NewPlace
  
  poddon.save

  
  bzid = txt4NewPlace.Tag
  
  
  
  s = "insert into stock(SITE_ID,ITEM_ID,LOCATION_ID,ORDER_ID,QTY_ON_HAND," & _
  "status,UNIT_COST,UOM,LOT_SN,REF_NUM," & _
  "ORD_NUM,PALLET_ID,custom_field1,custom_field6,custom_field11,custom_field5,exp_date,custom_field3,custom_field4,custom_field12)" & _
  "values(" & _
  "1," & item_id & "," & bzid & ",null," & MyRound2(netto) & _
   "," & status & ",0,'" & edizm & "','" & partia & "','" & qry & "'," & _
  "'" & qry & "'," & palID & "," & (MyRound(txt4Quantity) + MyRound(txt3Quantity)) & ",'" & country & "','" & killplace & "','" & made_date & "'," & MakeMSSQLDate(exp_date) & ",'" & MyRound2(txt4PackageWeight) & "','" & factory & "','" & brak & "') "

  
  Set cmd = New ADODB.Command
  cmd.CommandType = adCmdText
  cmd.CommandText = s
  Set cmd.ActiveConnection = conn
  On Error Resume Next
  cmd.Execute
   If err.Number <> 0 Then
    MsgBox err.Description
  End If
      
  Set loccode = conn.Execute("select code from location where id=" & bzid)
  If Not loccode.EOF Then
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdText
    cmd.CommandText = "update pallet set location_id=" & bzid & " where id=" & palID
    Set cmd.ActiveConnection = conn
    On Error Resume Next
    cmd.Execute
     If err.Number <> 0 Then
      MsgBox err.Description
    End If
  End If
  
  
  
End Sub



Private Sub ClosePoddon(ByVal poddon As ITTPL_DEF)
On Error Resume Next
  Dim conn As ADODB.Connection
  Set conn = Manager.GetCustomObjects("refref")
  Dim cmd As ADODB.Command
  Dim rs As ADODB.Recordset
  Dim oid As String
  Dim rlID As String
  Dim palID As String
  Dim palNum As String
  
  Set conn = Manager.GetCustomObjects("refref")
  If conn.State <> adStateOpen Then
    conn.Open
  End If
  

  palID = poddon.CorePalette_ID
  palNum = poddon.TheNumber
  
  
'''  Dim strs As ADODB.Recordset
'''  Dim LCRS As ADODB.Recordset
'''  Dim rsitem As ADODB.Recordset
'''
'''  Set strs = conn.Execute("select * from STOCK where PALLET_STATUS is null and  PALLET_ID=" & palID)
'''  Set LCRS = conn.Execute("select * from location where id=" & strs!location_id)
'''  Set rsitem = conn.Execute("select * from [item] where [id]=" & strs!item_id)
'''
'''  If strs.EOF Then
'''    MsgBox "Не обнаружены данные о палете"
'''    Exit Sub
'''  End If
'''
'''
'''  Dim w As Double
'''  Dim Q As Long
'''  Dim netto As Double
'''
'''  netto = MyRound(txt4Netto)
'''  w = MyRound(strs!QTY_ON_HAND)
'''  Q = MyRound("0" & strs!custom_field1)
  
'''    If txt4NewPlace <> "" Then
'''       Dim loccode As ADODB.Recordset
'''       Set loccode = conn.Execute("select * from location where code='" & txt4NewPlace & "'")
'''
'''       If Not loccode.EOF Then
'''         Set cmd = New ADODB.Command
'''         cmd.CommandType = adCmdText
'''         cmd.CommandText = "update pallet set location_id=" & loccode!id & " where id=" & palID
'''         Set cmd.ActiveConnection = conn
'''         On Error Resume Next
'''         cmd.Execute
'''          If err.Number <> 0 Then
'''           MsgBox err.Description
'''         End If
'''       End If
'''
'''       Set cmd = New ADODB.Command
'''       cmd.CommandType = adCmdText
'''       cmd.CommandText = "update STOCK set  location_id=" & loccode!id & " where PALLET_STATUS is null and PALLET_ID=" & palID
'''       Set cmd.ActiveConnection = conn
'''       On Error Resume Next
'''       err.Clear
'''       cmd.Execute
'''       If err.Number > 0 Then
'''         MsgBox err.Description
'''       End If
'''
'''    End If
  
''    Set cmd = New ADODB.Command
''    cmd.CommandType = adCmdText
''    cmd.CommandText = "update STOCK set QTY_ON_HAND =isnull(QTY_ON_HAND,0)-(" & MyRound2(netto) & "), CUSTOM_FIELD1='" & MyRound2(Q - MyRound(txt4Quantity)) & "' where PALLET_STATUS is null and PALLET_ID=" & palID
''    Set cmd.ActiveConnection = conn
''    On Error Resume Next
''    cmd.Execute
''    If err.Number > 0 Then
''      MsgBox err.Description
''    End If
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdText
    cmd.CommandText = "delete from STOCK where PALLET_STATUS is null and PALLET_ID=" & palID
    Set cmd.ActiveConnection = conn
    On Error Resume Next
    cmd.Execute
     If err.Number > 0 Then
      MsgBox err.Description
    End If
    
    poddon.Application.StatusID = "{E9BFB749-A606-4DEF-A429-07D636F108C6}"
    poddon.CurrentGood = ""
    poddon.CurrentPosition = ""
    poddon.PackageWeight = 0
    poddon.CaliberQuantity = 0
    poddon.save
End Sub

Private Sub txt5PNum_Change()
    If Len(txt5PNum) = 6 Then
        If Not (txt3PNum = txt5PNum Or txt4PNum = txt5PNum) Then
        
            MsgBox "Номер должен совпадать с одним из поддонов"
        End If
    End If
End Sub
