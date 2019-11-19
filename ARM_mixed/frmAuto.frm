VERSION 5.00
Begin VB.Form frmAuto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ожидание поддона"
   ClientHeight    =   1605
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmAuto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   4680
      Top             =   120
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Закрыть"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtPoddon 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5055
   End
   Begin VB.CommandButton cmd3ClearNum 
      Caption         =   "x"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Номер поддона"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 125

Option Explicit
'форма запуска автоматического режима определения заказа по поддону

Dim poddon As ITTPL_DEF

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    If txtPoddon = "" Then
        txtPoddon.SetFocus
    End If
End Sub

Private Sub txtPoddon_Change()
On Error Resume Next
  CheckPoddon
End Sub

'проверка поддона
Private Function CheckPoddon() As Boolean
On Error Resume Next
  If txtPoddon <> "" Then
    If Len(txtPoddon) = 6 Then
      Set poddon = Nothing
      Set poddon = FindPoddon(txtPoddon)
      If Not poddon Is Nothing Then
        ProcessPoddon
      Else
        MsgBox "Номер паддона: " & txtPoddon & "  не зарегистрирован"
      End If
    End If
  End If
End Function

' состояния для типа:ITTPL Палетта
' "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}" 'Взвешена
' "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" 'На складе с грузом
' "{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}" 'Отправлена с грузом
' "{E9BFB749-A606-4DEF-A429-07D636F108C6}" 'Пустая
' "{588C5203-1E59-408E-92A1-B3DFED8C19FA}" 'Списана


'запуск нужного заказа в зависимости от поддона
Private Sub ProcessPoddon()
    Dim id As String
    Dim ObjIn As ITTIN.Application
    Dim ObjOut As ITTOUT.Application
    
    If poddon.Application.StatusID = "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}" Then
        ' Ищем заказ, для которого взвешен этот паддон
        Dim checkrs As ADODB.Recordset
        Set checkrs = Session.GetData("select * from v_viewITTIN_ITTIN_EPL" & _
        " where  ITTIN_EPL_TheNumber_ID = '" & poddon.id & "' and " & _
        " INTSANCEStatusID in ('{EB3A7D03-EB3F-4541-AD93-D55C92BE02AC}','{49A919F7-94A6-49DE-9280-1EEAC973647B}')")
        If checkrs Is Nothing Then
         MsgBox err.Description
         Exit Sub
        End If
        
        If Not checkrs.EOF Then
            id = checkrs!InstanceID
            Set ObjIn = Manager.GetInstanceObject(id)
            Dim f As frmInWiz2
            Set f = New frmInWiz2
            Set f.Item = ObjIn
            f.INIT
            Load f
            f.Before1
            f.txtQryCode.Tag = ObjIn.ITTIN_DEF.Item(1).QryCode
            f.txtQryCode = GetBRIEFFromXMLField(ObjIn.ITTIN_DEF.Item(1).QryCode)
            f.Before2
            f.StepNo = 3
            f.NoMSG = True
            f.SinglePoddon = True
            f.InPoddon = poddon.code
            f.InWeight = poddon.Weight
            f.ProcessStatus
            f.Show vbModal
            Unload f
            Set f = Nothing
        End If
        txtPoddon = ""
        Exit Sub
    End If
    
    If poddon.Application.StatusID = "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" Then
            
        Dim f2 As frmOutWiz
        Set f2 = New frmOutWiz
        f2.Show vbModal
        Unload f2
        Set f2 = Nothing
        txtPoddon = ""
        Exit Sub
    End If
    
    MsgBox "Паддон: " & txtPoddon & "  в состоянии <" & poddon.Application.StatusName & "> и не может быть обработан."
    txtPoddon = ""
    
End Sub

Private Sub txt3Poddon_Change()

End Sub
