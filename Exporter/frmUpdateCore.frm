VERSION 5.00
Object = "{1801C003-859D-471D-BF31-D4428050324B}#2.1#0"; "MTZ_PANEL.ocx"
Begin VB.Form frmUpdateCore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Обновить данные в Core по заказу"
   ClientHeight    =   2085
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmUpdateCore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkInsert 
      Caption         =   "Вставить данные"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox txtQryCode 
      Height          =   300
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Код заказа"
      Top             =   480
      Width           =   3375
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
   Begin MTZ_PANEL.DropButton cmdQryCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   3
      Tag             =   "refopen.ico"
      ToolTipText     =   "Код заказа"
      Top             =   480
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Caption         =   ""
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label lblQryCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Код заказа:"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "frmUpdateCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim StepNo As Integer
Dim XMLQryCode As String
Dim XMLTheClient As String
Dim Item As ITTIN.Application



Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    StepNo = 0
    XMLQryCode = "<SQLData>"
    XMLQryCode = XMLQryCode & "<connectionstring>ref</connectionstring>"
    XMLQryCode = XMLQryCode & "<connectionprovider>ref</connectionprovider>"
    XMLQryCode = XMLQryCode & "<query>select A.ID [КОД] , convert(varchar(30),A.NUMBER) +'  от ' + convert(varchar(30),A.ORD_DATE,111)  [Название], B.Name [Клиент]  from RECEIVING_ORDER A left join PARTNER B on A.PARTNER_ID=B.ID  </query>"
    XMLQryCode = XMLQryCode & "<IDFieldName>КОД</IDFieldName>"
    XMLQryCode = XMLQryCode & "<BriefFields>Название</BriefFields>"
    XMLQryCode = XMLQryCode & "</SQLData>"
    
    txtQryCode.Tag = XMLQryCode
    
End Sub
Private Sub cmdQryCode_Click()
  On Error Resume Next

  Dim pars As New NamedValues
  Dim res As NamedValues
  If (txtQryCode.Tag = "") Then
    ' call MsgBox("Нет данных для запроса")
  Else
    txtQryCode.Tag = Replace(txtQryCode.Tag, "%ID%", " 1=1 ")
    Call pars.Add("xml", txtQryCode.Tag)
  End If
  Set res = Manager.GetSQLDataDialog(pars)
  If (Not res Is Nothing) Then
    Dim resStr As String
    resStr = res.Item("RESULT").Value
    If (resStr = "OK") Then
      txtQryCode.Tag = res.Item("xml").Value
      If (txtQryCode.Text <> res.Item("brief").Value) Then
        txtQryCode.Text = res.Item("brief").Value
        'mIDQryCode = res.Item("ID").Value
        FindItem
      End If
    Else
      Dim errStr As String
      errStr = res.Item("ErrorDescription").Value
      If (errStr <> vbNullString) Then
       Call MsgBox("Ошибка исполнения: " & errStr, vbOKOnly + vbCritical)
     End If
    End If
  End If
End Sub




Private Sub FindItem()
On Error Resume Next
  'Найти заказ у в нашей базе
  Dim rs As ADODB.Recordset
  Dim id As String
  Dim qID As String
  qID = Manager.GetIDFromXMLField(txtQryCode.Tag)
  id = ""
  Set rs = Session.GetData("select instanceid from ITTIN_DEF where QryCode like '%<ID>" & qID & "</ID>%'")
  If Not rs Is Nothing Then
    If Not rs.EOF Then
      id = rs!InstanceID
    End If
  End If
  rs.Close
  
  If id <> "" Then
    Set Item = Manager.GetInstanceObject(id)
  End If
End Sub

Private Sub OKButton_Click()
  If Not Item Is Nothing Then
    Dim qry As ITTIN.Application
    Dim qline As ITTIN_QLINE
    Dim pal As ITTIN_PALET
    Dim poddon As ITTPL_DEF
    Dim i As Long, j As Long
    Set qry = Item
    If Not qry Is Nothing Then
      If MsgBox("Если c палет данного заказа производилась отгрузка, то обновление приведет к ошибкам учета " & vbCrLf & _
      "Обновить данные по текущем состоянию документа ?", vbYesNo + vbExclamation, "ВНИМАНИЕ") = vbYes Then
        If chkInsert.Value = vbChecked Then
          CleanRCVAtCore qry
        End If
        
        For i = 1 To qry.ITTIN_QLINE.Count
          For j = 1 To qry.ITTIN_QLINE.Item(i).ITTIN_PALET.Count
            If chkInsert.Value = vbChecked Then
              Set pal = qry.ITTIN_QLINE.Item(i).ITTIN_PALET.Item(j)
              SaveRCVRowToCore qry, qry.ITTIN_QLINE.Item(i), qry.ITTIN_QLINE.Item(i).ITTIN_PALET.Item(j), pal.BufferZonePlace, txtQryCode.Text
            Else
              UpdateMyPalet qry.ITTIN_QLINE.Item(i).ITTIN_PALET.Item(j)
            End If
            Label1 = qry.ITTIN_QLINE.Item(i).ITTIN_PALET.Item(j).brief
            DoEvents
          Next
        Next
        MsgBox "Данные обновлены"
        Unload Me
      End If
    End If
  End If
End Sub
