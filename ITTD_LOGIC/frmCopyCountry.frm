VERSION 5.00
Begin VB.Form frmCopyCountry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " опирование справочника стран"
   ClientHeight    =   5475
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelCopy 
      Caption         =   "..."
      Height          =   375
      Left            =   8880
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtCopyFrom 
      Height          =   375
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   3735
   End
   Begin VB.CommandButton cmdClient 
      Caption         =   "..."
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtClient 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.ListBox lstCountry 
      Height          =   2985
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   1440
      Width           =   9375
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "«аполнить все"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "ќчистить все"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "ќтмена"
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "—копировать"
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "—писок стран дл€ копировани€"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Ќовый поклажедатель"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "ѕоклажедатель, с которого копируютс€ страны"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmCopyCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 25

Option Explicit
Public dic As ITTD.Application
Attribute dic.VB_VarHelpID = 30
Private manager As MTZManager.Main
Private session As MTZSession.session
Private LastSupplier As String


' закрытие формы
Private Sub CancelButton_Click()
  Unload Me
End Sub


' сброс галочек в списке стран
Private Sub cmdClearAll_Click()
  Dim i As Long
  For i = 0 To lstCountry.ListCount - 1
    lstCountry.Selected(i) = False
  Next
End Sub

'  выбрать клиента из базы данных CORE
Private Sub cmdClient_Click()
  Dim f As frmClienList
  Set f = New frmClienList
  Set f.manager = manager
  f.Show vbModal
  If f.ClientText <> "" Then
    txtClient.Text = f.ClientText
    Dim rs As ADODB.Recordset
    LastSupplier = txtCopyFrom.Text
    Set rs = session.GetData("select distinct name from ittd_country where thesupplier like '%" & txtClient.Text & "%' order by name")
    lstCountry.Clear
    While Not rs.EOF
      lstCountry.AddItem rs!Name
      rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
  
  End If
  Unload f
  Set f = Nothing
End Sub

' выбор всех стран
Private Sub cmdSelAll_Click()
  Dim i As Long
  For i = 0 To lstCountry.ListCount - 1
    lstCountry.Selected(i) = True
  Next
End Sub


'  выбрать клиента -источник из базы данных CORE
Private Sub cmdSelCopy_Click()
  Dim f As frmClienList
  Set f = New frmClienList
  Set f.manager = manager
  f.Caption = "¬ыбор поклажедател€ дл€ копировани€ списка"
  f.Show vbModal
  If f.ClientText <> "" Then
    txtCopyFrom.Text = f.ClientText
     Dim rs As ADODB.Recordset
    LastSupplier = txtCopyFrom.Text
    Set rs = session.GetData("select distinct name from ittd_country where thesupplier like '%" & txtCopyFrom.Text & "%' order by name")
    lstCountry.Clear
    While Not rs.EOF
      lstCountry.AddItem rs!Name
      rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
  End If
  Unload f
  Set f = Nothing
End Sub

' начальна€ загрузка формы
Private Sub Form_Load()
  Dim rs As ADODB.Recordset

    ' загрузить весь список стран
    If Not dic Is Nothing Then
      Set manager = dic.manager
      Set session = dic.MTZSession
      Set rs = session.GetData("select distinct name from ittd_country order by name")
      lstCountry.Clear
      While Not rs.EOF
        lstCountry.AddItem rs!Name
        rs.MoveNext
      Wend
      rs.Close
      Set rs = Nothing
    End If

End Sub

' запуск процесса копировани€
Private Sub OKButton_Click()
  If txtClient.Text <> "" Then
    If lstCountry.SelCount > 0 Then
      If MsgBox("—копировать список стран, заводов и боен дл€ нового поклажедател€?", vbYesNo + vbQuestion, "ѕодтверждение") = vbYes Then
          CopyCountries
          
          Unload Me
      End If
    Else
      MsgBox "Ќе выбраны страны дл€ копировани€", vbCritical + vbOKOnly, "ќшибка заполнени€"
    End If
  Else
    MsgBox "Ќеобходимо задать поклажедател€", vbCritical + vbOKOnly, "ќшибка заполнени€"
  End If
  
End Sub

' скопировать страны
Private Sub CopyCountries()

  Dim rs As ADODB.Recordset
  Dim i As Long
  Dim cname As String
  
  For i = 0 To lstCountry.ListCount - 1
    If lstCountry.Selected(i) Then
      cname = lstCountry.List(i)
      CopyOneCountry cname
    
    End If
  Next
 
 
End Sub


' скопировать одну страну
' ѕараметры:
' CountryName - название страны
Private Sub CopyOneCountry(ByVal CountryName As String)
  Dim newCountry As ITTD.ITTD_COUNTRY
  Dim oldCountry As ITTD.ITTD_COUNTRY
  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset
  
  
  Me.Caption = CountryName
  DoEvents
  Dim cl As Integer
  cl = Len(txtClient.Text) + 1
  
  
  
  Set rs = session.GetData("select * from ittd_country where name ='" & Replace(CountryName, "'", "''") & "' and theSupplier like '%" & Replace(LastSupplier, "'", "''") & "%'")
  
  If Not rs.EOF Then
    If Len(rs!TheSupplier) + cl < 127 Then
        session.GetData "update ittd_country set thesupplier = '" & Replace(txtClient.Text, "'", "''") & "," & rs!TheSupplier & "' where ittd_countryid='" & rs!ittd_countryid & "'"
    Else
    
        Set oldCountry = dic.FindRowObject("ITTD_COUNTRY", rs!ittd_countryid)
        
        Set rs2 = session.GetData("select * from ittd_country where name ='" & Replace(CountryName, "'", "''") & "' and theSupplier  like '%" & Replace(txtClient.Text, "'", "''") & "%'")
        
        If rs2.EOF Then
          Set newCountry = dic.ITTD_COUNTRY.Add
          With newCountry
            .Code1 = rs!Code1 & ""
            .Code2 = rs!Code2 & ""
            .TheSupplier = txtClient.Text
            .Name = CountryName
            .Save
          End With
        Else
          Set newCountry = dic.FindRowObject("ITTD_COUNTRY", rs2!ittd_countryid)
        End If
        
        CopyFactories oldCountry, newCountry
    End If
  End If

End Sub

' —копировать заводы
' параметры:
' SrcCountry - »сходна€ страна
' DstCountry - страна в которую переносить
Private Sub CopyFactories(ByRef SrcCountry As ITTD_COUNTRY, ByRef DstCountry As ITTD_COUNTRY)
 Dim i As Long
 Dim fc As ITTD_FACTORY
 
 Dim rs As ADODB.Recordset
 Set rs = session.GetData("select ITTD_factoryid from ITTD_factory where country='" & SrcCountry.ID & "'")
 While Not rs.EOF
 'For i = 1 To dic.ITTD_FACTORY.Count
 
  
  Set fc = dic.ITTD_FACTORY.Item(rs!ittd_factoryid)
  
'  If Not fc.Country Is Nothing Then
'    If fc.Country Is SrcCountry Then
      CopyOneFactory DstCountry, fc
'    End If
'  End If
 'Next
 rs.MoveNext
 Wend
 rs.Close
 Set rs = Nothing


End Sub

' скопировать один завод
' параметры:
' DstCountry - страна куда копировать
' SrcFactory - завод, который надо скопировать
Private Sub CopyOneFactory(ByRef DstCountry As ITTD_COUNTRY, ByRef SrcFactory As ITTD_FACTORY)
 Dim i As Long
 Me.Caption = DstCountry.Name & "->" & SrcFactory.Name
 DoEvents
 
 Dim fc As ITTD_FACTORY
 Dim rs As ADODB.Recordset
 Set rs = session.GetData("select ITTD_factoryid from ITTD_factory where  name ='" & Replace(SrcFactory.Name, "'", "''") & "' and country='" & DstCountry.ID & "'")
 If Not rs.EOF Then
  
  'For i = 1 To dic.ITTD_FACTORY.Count
  Set fc = dic.ITTD_FACTORY.Item(rs!ittd_factoryid)
  'If Not fc.Country Is Nothing Then
  '  If fc.Country Is DstCountry And UCase(fc.Name) = UCase(SrcFactory.Name) Then
      ' така€ страна уже есть в списке стран
      CopyKillplaces SrcFactory, fc
      Exit Sub
   ' End If
  End If
 'Next
 
 ' не нашли такого завода
 Set fc = dic.ITTD_FACTORY.Add
 With fc
  .Name = SrcFactory.Name
  .Code1 = SrcFactory.Code1
  .Code2 = SrcFactory.Code2
  Set .Country = DstCountry
  .Save
 End With
 
 CopyKillplaces SrcFactory, fc
 
End Sub


'  опировать бойни
' параметры:
' SrcFactory - «авод откуда брать бойни
' DstFactory - «авод куда копировать
Private Sub CopyKillplaces(ByRef SrcFactory As ITTD_FACTORY, ByRef DstFactory As ITTD_FACTORY)
 Dim i As Long
 Dim kp As ITTD_KILLPLACE
 
'
' For i = 1 To dic.ITTD_KILLPLACE.Count
'  Set kp = dic.ITTD_KILLPLACE.Item(i)
'
 Dim rs As ADODB.Recordset
 Set rs = session.GetData("select ITTD_KILLPLACEid from ITTD_KILLPLACE where factory='" & SrcFactory.ID & "'")
 While Not rs.EOF
 'For i = 1 To dic.ITTD_FACTORY.Count
 
  
  Set kp = dic.ITTD_KILLPLACE.Item(rs!ITTD_KILLPLACEid)
'  If Not kp.Factory Is Nothing Then
'    If kp.Factory Is SrcFactory Then
      CopyOneKillPlace DstFactory, kp
'    End If
'  End If
' Next
  rs.MoveNext
 Wend
 rs.Close
 Set rs = Nothing

End Sub


' скопировать одну бойню
' параметры:
' DstFactory -  завод  к которому присоединить бойню
' SrcKP - бойн€ дл€ присоединени€
Private Sub CopyOneKillPlace(ByRef DstFactory As ITTD_FACTORY, ByRef SrcKP As ITTD_KILLPLACE)
 Me.Caption = DstFactory.Country.Name & "->" & DstFactory.Name & "->" & SrcKP.Name
 DoEvents
 Dim i As Long
 Dim kp As ITTD_KILLPLACE
 
 Dim rs As ADODB.Recordset
 Set rs = session.GetData("select ITTD_KILLPLACEid from ITTD_KILLPLACE where  name ='" & Replace(SrcKP.Name, "'", "''") & "' and factory='" & DstFactory.ID & "'")
 If Not rs.EOF Then
 
 'For i = 1 To dic.ITTD_KILLPLACE.Count
  'Set kp = dic.ITTD_KILLPLACE.Item(rs!ITTD_KILLPLACEid)
  'If Not kp.Factory Is Nothing Then
  '  If kp.Factory Is DstFactory And UCase(kp.Name) = UCase(SrcKP.Name) Then
      ' така€ бойн€ уже есть в списке стран
      Exit Sub
  '  End If
  End If
 'Next
 
 ' не нашли такой бойни
 Set kp = dic.ITTD_KILLPLACE.Add
 With kp
  .Name = SrcKP.Name
  .Code1 = SrcKP.Code1
  .Code2 = SrcKP.Code2
  Set .Factory = DstFactory
  .Save
 End With
 
End Sub

