VERSION 5.00
Begin VB.Form frmSSCC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Привязка SSCC кода"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Закрыть"
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtResult 
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   6975
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   6600
      Top             =   0
   End
   Begin VB.TextBox txtSapCode 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6975
   End
   Begin VB.TextBox txtPoddon 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6975
   End
   Begin VB.Label Label3 
      Caption         =   "Результат обработки"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Баркод SAP"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Код поддона"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmSSCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim poddonFound As Boolean
Dim ssccFound As Boolean


Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Timer1_Timer()
  If Len(txtPoddon) < 6 Then
    poddonFound = False
    txtPoddon.Locked = False
    txtPoddon.SetFocus
    txtSapCode.Locked = True
    Exit Sub
  End If
  If Len(txtSapCode) < 38 Then
     ssccFound = False
     txtSapCode.Locked = False
    txtSapCode.SetFocus
    txtPoddon.Locked = True
    Exit Sub
  End If
End Sub

Private Sub txtPoddon_Change()
  
  If Len(txtPoddon) = 6 Then
    If poddonFound Then
      Exit Sub
    End If
    
    ' ищем поддон в открытых заказах на приемку груза
    poddonFound = CheckPoddon(txtPoddon.Text)
    If poddonFound Then
      txtResult = "Поддон <" & txtPoddon.Text & "> обнаружен среди принятых поддонов" & vbCrLf & txtResult.Text
    Else
      txtResult = "Неерный номер поддона <" & txtPoddon.Text & "> " & vbCrLf & txtResult.Text
      txtPoddon.Text = ""
    End If
    
  Else
   If Len(txtPoddon) > 7 Then
    txtPoddon = ""
   End If
    poddonFound = False
    
  End If
  
  
  
End Sub

Private Sub txtSapCode_Change()
 
  If Len(txtSapCode) >= 38 Then
    If ssccFound Then
      Exit Sub
    End If
    
    ' ищем код в файлах транспортного заказа
    If Left(txtSapCode, 2) <> "00" Then
        MsgBox "Отсканируйте другой штрихкод"
        txtSapCode = ""
        Exit Sub
    End If
    
    Dim s0 As String, s1 As String, s2  As String, s3  As String, s4  As String
    
    s0 = Mid(txtSapCode, 3, 18)
    s1 = Mid(txtSapCode, 3, 1)
    s2 = Mid(txtSapCode, 4, 7)
    s3 = Mid(txtSapCode, 11, 9)
    s4 = Mid(txtSapCode, 20, 1)
    
    If CheckSSCC(s0) Then
      If SaveSapCode(txtPoddon.Text, s0) Then
        txtResult = s1 & " " & s2 & " " & s3 & " " & s4 & " обнаружен в файле транспортного заказа" & vbCrLf & txtResult.Text
        txtSapCode = ""
        txtPoddon = ""
        ssccFound = False
        poddonFound = False
      End If
    End If
    
    
  End If
End Sub


Private Function CheckPoddon(poddon As String) As Boolean
  Dim rs As ADODB.Recordset
  Set rs = Session.GetData("select * from v_autoittin_palet where ittin_palet_thenumber='" & poddon & ";' and intsancestatusid <>'E3728A5B-6B62-48BF-9E5A-D4F0BCBFC75B'")
  If Not rs Is Nothing Then
    If Not rs.EOF Then
      CheckPoddon = True
    End If
  End If
End Function

Private Function CheckSSCC(SSCC As String) As Boolean
  Dim rs As ADODB.Recordset
  Set rs = Session.GetData("select * from itttz_lines where sscc ='" & SSCC & "'")
  If Not rs Is Nothing Then
    If Not rs.EOF Then
      CheckSSCC = True
    End If
  End If
End Function


Private Function SaveSapCode(poddon As String, sapCode As String) As Boolean
  Dim rs As ADODB.Recordset
  Set rs = Session.GetData("select * from v_autoittin_palet where ittin_palet_thenumber='" & poddon & ";' and statusname <>'Приемка заершена'")
  If Not rs Is Nothing Then
    If Not rs.EOF Then
      Session.GetData "update ittin_palet set sscc = '" & sapCode & "'  where ittin_paletid='" & rs!ittin_paletid & "'"
      SaveSapCode = True
    End If
  End If
  
    
  
  

End Function
