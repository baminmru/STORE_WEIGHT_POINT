VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmDivideWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������� �� ��� �������"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "���� �� ����� �������"
      Height          =   4575
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txt4NewPlace 
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton cmd4ClearW 
         Caption         =   "X"
         Height          =   375
         Left            =   2400
         TabIndex        =   35
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txt4GoodWeight 
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txt4FullWeight 
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txt4Quantity 
         Height          =   405
         Left            =   240
         TabIndex        =   32
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txt4PWeight 
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txt4PNum 
         Height          =   405
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txt4PackageWeight 
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox txt4Netto 
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3240
         Width           =   2895
      End
      Begin VB.CommandButton cmd6FindCell 
         Caption         =   "..."
         Height          =   375
         Left            =   5280
         TabIndex        =   27
         ToolTipText     =   "����� ������"
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "��� ������ ������"
         Height          =   375
         Left            =   2880
         TabIndex        =   44
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label17 
         Caption         =   "����� � �������� ����"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2880
         TabIndex        =   43
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label12 
         Caption         =   "��� ����� � ��������"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label11 
         Caption         =   "���������� �������"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "��� ������� ��."
         Height          =   255
         Left            =   2880
         TabIndex        =   40
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "������ �"
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "��� ����� �������� ��."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "��� ������ �����"
         Height          =   255
         Left            =   2880
         TabIndex        =   37
         Top             =   3000
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�������� ������"
      Height          =   4335
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   7215
      Begin VB.CommandButton cmd3ClearNum 
         Caption         =   "X"
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         ToolTipText     =   "������ ����� ��� ���"
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
         ToolTipText     =   "�������� ��� �  �����"
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
      Begin VB.TextBox txtMainCell 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3360
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "������"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "��� ����� � ��������"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "��� ����� �����"
         Height          =   375
         Left            =   2760
         TabIndex        =   18
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label18 
         Caption         =   "���������� �������"
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label19 
         Caption         =   "��� �������"
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label20 
         Caption         =   "��� ����� ��������"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "������ ��������� ��������"
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   3000
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "����� ������"
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   6015
      Begin VB.CommandButton Command2 
         Caption         =   "x"
         Height          =   375
         Left            =   5280
         TabIndex        =   24
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "x"
         Height          =   375
         Left            =   5280
         TabIndex        =   23
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txt3Weight 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1680
         Width           =   5055
      End
      Begin VB.TextBox txt3Poddon 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "��� ������� ��."
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "����� �������"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   720
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "��������"
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "�����"
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
Attribute VB_Name = "frmDivideWizard"
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
Private Poddon As ITTPL_DEF
Private isFull As Boolean
Private item_id As String
Private Country As String
Private Factory As String
Private killplace As String
Private qry As String
Private edizm As String
Private exp_date As String
Private made_date As String
Private status As String
Private brak As String
Private partia As String


Private Sub MyBeep(ByVal BeepType As String)
      If Not wave Is Nothing Then
        On Error Resume Next
        wave.OpenFile App.Path & "\" & BeepType & ".wav"
        wave.Play
      End If
End Sub


Private Function CheckPoddon2() As Boolean
On Error Resume Next
  If txt3Poddon <> "" Then
    If Len(txt3Poddon) = 6 Then
      Set Poddon = Nothing
      Set Poddon = FindPoddon(txt3Poddon)
      If Not Poddon Is Nothing Then
        MyBeep "Nomer"
        txt3Weight = Poddon.Weight
      Else
        MsgBox "����� �������: " & txt3Poddon & "  �� ���������������"
      End If
    End If
  End If
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
  f.Country = ""
  f.Country = Country
  f.Factory = ""
  f.Factory = Factory
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
  If StepNo = 3 Then
    If Not After3 Then
      Exit Sub
    End If
  End If
  StepNo = StepNo + 1

  ProcessStatus
End Sub

Private Sub Command1_Click()
  txt3Poddon = ""
End Sub

Private Sub Command2_Click()
  txt3Weight = "0"
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
    If txt3Poddon = "" Then
      txt3Poddon.SetFocus
    End If
    If txt3Weight = "0" Or Not IsNumeric(txt3Weight) Then
      w = GetWeight
      If w > 0 Then
        txt3Weight = Round(w + 0.001, 1)
        MyBeep "Poddon"
      End If
    End If
  End If
  
  
  If StepNo = 3 Then
    
    If txt4FullWeight = "0" Or Not IsNumeric(txt4FullWeight) Then
      w = GetWeight
      If w > 0 Then
        txt4FullWeight = Round(w + 0.001, 1)
        MyBeep "Gruz"
      End If
    End If
  End If

End Sub

Private Sub txt3PNum_Change()
On Error Resume Next
  
  If CheckPoddon Then
    txt3PWeight = Poddon.Weight
    txt3FullWeight = Poddon.CurrentWeightBrutto
    txt3Quantity = Poddon.CaliberQuantity
    txt3PackageWeight = Poddon.PackageWeight / Poddon.CaliberQuantity
  End If
End Sub

Private Function CheckPoddon() As Boolean
On Error Resume Next
  Dim result As Boolean
  result = True
  If txt3PNum <> "" Then
    If Len(txt3PNum) = 6 Then
      Set Poddon = Nothing
      Set Poddon = FindPoddon(txt3PNum)
      If Not Poddon Is Nothing Then
        If Poddon.Application.StatusID = "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" Then
          Dim conn As ADODB.Connection
          Set conn = Manager.GetCustomObjects("refref")
          If conn.State <> adStateOpen Then
            conn.Open
          End If
          
          Dim rs As ADODB.Recordset
          'poddon.CurrentGood
          
          Set rs = conn.Execute("select * from stock where PALLET_STATUS is null and pallet_id=" & Poddon.CorePalette_ID)
          If rs.EOF Then
            MsgBox "����� �������: " & txt3PNum & "  �� ��������� � ���� CORE IMS"
            result = False
          Else
            MyBeep "Nomer"
            Dim lid  As String
            lid = "" & rs!location_id
            
            item_id = rs!item_id
            Country = rs!custom_field6
            Factory = rs!custom_field4
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
              txtMainCell = rs!Code
              txtMainCell.Tag = rs!id
            End If
            txt3Weight = Poddon.Weight
            
         
            
'            Printer.FontBold = False
'            Printer.Print "������ �������������: ";
'            Printer.FontBold = True
'            Printer.Print strs!custom_field6 & ""
'
'            Printer.FontBold = False
'            Printer.Print "�������������: ";
'            Printer.FontBold = True
'            Printer.Print strs!custom_field4 & ""
'
'            Printer.FontBold = False
'            Printer.Print "�����: ";
'            Printer.FontBold = True
'            Printer.Print strs!custom_field11 & ""
'
'            Printer.FontBold = False
'            Printer.Print "������: ";
'            Printer.FontBold = True
'            Printer.Print strs!LOT_SN & ""
            
            
            
          End If
        Else
          MsgBox "��������� �������: " & txt3PNum & "  ����������� ������� (" & Poddon.Application.StatusName & ")"
          result = False
        End If
      Else
        MsgBox "����� �������: " & txt3PNum & "  �� ���������������"
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
  
  cmdNext.Caption = "�����"
  cmdCancel.Caption = "��������"
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
    
    GetWeight4 = 0  ' �� ��������� ������
    Exit Function
    
answer_s1:
    
    ws = MSComm1.Input
    ' ������ ��� ��� ��������
    If Asc(Mid(ws, 1, 1)) >= 128 Then
    
      ''''''''''''''''''''''''''''''''''''
      '�������� !!!
      '
      ' ���� ����� ��������� �������
      start = Timer   ' Set start time.
      Do While Timer < start + 0.3
      Loop
      
      ' ���������� ��� ���
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
    
    GetWeight4 = 0 ' ��� ������� ������
    Exit Function
    
answer_s2:

    ws = MSComm1.Input
    
    ' ������ ��� ��� ���� ��������
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
    
    GetWeight4 = 0 ' ��� �� ��������, ��� ��� ������
    Exit Function
    
answer_w1:

    ' ������ ��������� ����
    ws1 = MSComm1.Input
    
    
    ''''''''''''''''''''''''''''''''''''
    '�������� !!!
    '
    ' ���� ����� ��������� �������
    start = Timer   ' Set start time.
    Do While Timer < start + 0.3
    Loop
    
    ' ���������� ��� ��� ���
    MSComm1.output = Chr(69)
    start = Timer   ' Set start time.
    Do While Timer < start + 0.2
    Loop
    
    If MSComm1.InBufferCount > 0 Then GoTo answer_w2
    start = Timer   ' Set start time
    Do While Timer < start + 0.5
       If MSComm1.InBufferCount > 0 Then GoTo answer_w2
    Loop
    
    GetWeight4 = 0 '  ��� ������
    Exit Function
      
answer_w2:
    ws = MSComm1.Input
  
    If ws1 = ws Then
      GetWeight4 = (Asc(Mid(ws, 2, 1)) * 256 + Asc(Mid(ws, 1, 1))) / 10
    Else
      GetWeight4 = 0 ' ��� �� ��������, ���������� ���������
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
  txt4PNum = txt3Poddon
  txt4PWeight = txt3PWeight
  txt4PackageWeight = txt3PackageWeight
End Sub


Private Sub txt4FullWeight_Change()
  On Error Resume Next
  txt4GoodWeight = Round(MyRound(txt4FullWeight) - MyRound(txt4PWeight) + 0.001, 1)
End Sub

Private Function After3() As Boolean
  Dim result As Boolean
  If MyRound(txt4Netto) <= 0 Then
    MsgBox "�� ����� ��� ����� �� ����� �������"
    After3 = False
    Exit Function
  End If
  
  If MyRound(txt4Netto) > MyRound(txt3GoodWeight) Then
    MsgBox "��� ����� �������������� ��������� ��������"
    After3 = False
    Exit Function
  End If
  
  If MyRound(txt4Quantity) > MyRound(txt3Quantity) Then
    MsgBox "���������� ������� �������������� ��������� ��������"
    After3 = False
    Exit Function
  End If
  
  If MsgBox("���������������� ��������� ������?", vbYesNo) = vbYes Then
  
    
    SaveOLDPoddon FindPoddon(txt3PNum)
    SaveNewPoddon FindPoddon(txt4PNum)
  
    PrintSticker FindPoddon(txt3PNum)
    PrintSticker FindPoddon(txt4PNum)
  
    After3 = True
  End If
  

End Function

Private Sub Txt4PackageWeight_Change()
txt4GoodWeight = MyRound(txt4FullWeight) - (MyRound(txt4PackageWeight) * MyRound(txt4Quantity))
End Sub

Private Sub txt4Quantity_Change()
  txt4Netto = MyRound(txt4GoodWeight) - (MyRound(txt4PackageWeight) * MyRound(txt4Quantity))
End Sub


Private Sub SaveNewPoddon(Poddon As ITTPL_DEF)
On Error Resume Next
  Dim conn As ADODB.Connection
  Set conn = Manager.GetCustomObjects("refref")
  Dim cmd As ADODB.Command
  Dim rs As ADODB.Recordset
  Dim rsitem As ADODB.Recordset
  
  Dim rlID As String
  Dim palID As String
  Dim oid As String
 
  palID = Poddon.CorePalette_ID
  
  
  ' ����������� ��������� ����� � �������� ����
  Dim bzrs As ADODB.Recordset
  Dim loccode As ADODB.Recordset
  Dim bzid As String
  Set conn = Manager.GetCustomObjects("refref")
  If conn.State <> adStateOpen Then
    conn.Open
  End If
  
  
  Dim s As String
  Dim netto As Double
  
  
  
  
  
  netto = MyRound(txt4Netto)
  
  Poddon.Weight = MyRound(txt4PWeight)
  Poddon.CaliberQuantity = MyRound(txt4Quantity)
  Poddon.PackageWeight = MyRound(txt4PackageWeight) * MyRound(txt4Quantity)
  Poddon.CurrentWeightBrutto = MyRound(txt4FullWeight)
  Poddon.CurrentPosition = txt4NewPlace
  
  Poddon.save
  Poddon.Application.StatusID = "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}"
  
  bzid = txt4NewPlace.Tag
  
  
  
  s = "insert into stock(SITE_ID,ITEM_ID,LOCATION_ID,ORDER_ID,QTY_ON_HAND," & _
  "status,UNIT_COST,UOM,LOT_SN,REF_NUM," & _
  "ORD_NUM,PALLET_ID,custom_field1,custom_field6,custom_field11,custom_field5,exp_date,custom_field3,custom_field4,custom_field12)" & _
  "values(" & _
  "1," & item_id & "," & bzid & ",null," & MyRound2(netto) & _
   "," & status & ",0,'" & edizm & "','" & partia & "','" & qry & "'," & _
  "'" & qry & "'," & palID & "," & MyRound2(txt4Quantity) & ",'" & Country & "','" & killplace & "','" & made_date & "'," & MakeMSSQLDate(exp_date) & ",'" & MyRound2(txt4PackageWeight) & "','" & Factory & "','" & brak & "') "

  
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



Private Sub SaveOLDPoddon(ByVal Poddon As ITTPL_DEF)
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
  

  palID = Poddon.CorePalette_ID
  palNum = Poddon.TheNumber
  
  
  Dim strs As ADODB.Recordset
  Dim LCRS As ADODB.Recordset
  Dim rsitem As ADODB.Recordset
  
  Set strs = conn.Execute("select * from STOCK where PALLET_STATUS is null and  PALLET_ID=" & palID)
  Set LCRS = conn.Execute("select * from location where id=" & strs!location_id)
  Set rsitem = conn.Execute("select * from [item] where [id]=" & strs!item_id)
    
  If strs.EOF Then
    MsgBox "�� ���������� ������ � ������"
    Exit Sub
  End If
  
  
  Dim w As Double
  Dim Q As Long
  Dim netto As Double
  
  netto = MyRound(txt4Netto)
  w = MyRound(strs!QTY_ON_HAND)
  Q = MyRound("0" & strs!custom_field1)
  
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
  
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdText
    cmd.CommandText = "update STOCK set QTY_ON_HAND =isnull(QTY_ON_HAND,0)-(" & MyRound2(netto) & "), CUSTOM_FIELD1='" & MyRound2(Q - MyRound(txt4Quantity)) & "' where PALLET_STATUS is null and PALLET_ID=" & palID
    Set cmd.ActiveConnection = conn
    On Error Resume Next
    cmd.Execute
    If err.Number > 0 Then
      MsgBox err.Description
    End If
    
    
    
End Sub
