VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 750
Option Explicit

'
' ����� ������� GUI ���������� ���������� ���  ���������� ����������� ���������� ������  MDI ����


Private mOpenModal As Boolean
Dim opf As OpenFormInfo
Dim openForms As Collection
Private m_Typename As String
Private STDTypeGUI As Object

' ������������� ������ ��� ����������� ���� ���������
'Parameters:
'[IN]   tn , ��� ���������: String  - ��� ���������
'See Also:
'  CloseClass
'  EraseForm
'  FormCount
'  GuiPath
'  IsModal
'  Show
'  ShowAddForm
'Example:
'  call me.INIT(...���������...)
Public Sub INIT(ByVal tn As String)
Attribute INIT.VB_HelpID = 775
m_Typename = tn
Set STDTypeGUI = CreateObject(tn & "GUI.GUI")
End Sub

Private Sub Class_Initialize()
 Set openForms = New Collection

 
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  CloseClass
End Sub

' �������� ����� ��������������
'Parameters:
'[IN]   mode , ��� ���������: String,
'[IN][OUT]   RowItem , ��� ���������: Object  - ...
'Returns:
' Boolean, ��������� ����������:
'   true  -
'   false -
'See Also:
'  CloseClass
'  EraseForm
'  FormCount
'  GuiPath
'  INIT
'  IsModal
'  Show
'Example:
' dim variable as Boolean
' variable = me.ShowAddForm(...���������...)
Public Function ShowAddForm(ByVal mode As String, RowItem As Object) As Boolean
Attribute ShowAddForm.VB_HelpID = 790
On Error Resume Next
ShowAddForm = STDTypeGUI.ShowAddForm(mode, RowItem)
End Function

' ������������ ������, �������� �������� � ����������� ����
'Parameters:
' ���������� ���
'See Also:
'  EraseForm
'  FormCount
'  GuiPath
'  INIT
'  IsModal
'  Show
'  ShowAddForm
'Example:
'  call me.CloseClass()
Public Sub CloseClass()
Attribute CloseClass.VB_HelpID = 755
  Dim f As OpenFormInfo
  On Error Resume Next
  For Each f In openForms
      Set f.Form.Host = Nothing
      Unload f.Form
      Set f.Form = Nothing
  Next
  Set openForms = Nothing
  STDTypeGUI.CloseClass
  Set STDTypeGUI = Nothing
End Sub


' ����� ����� �������������� � ��������� ����������� ����
Private Function FindForm(ByVal InstanceID As String, ByVal mode As String) As Object
  Dim f As OpenFormInfo
  For Each f In openForms
    If f.InstanceID = InstanceID And f.mode = mode And f.Modal = False Then
      Set FindForm = f
      Exit Function
    End If
  Next
End Function

'���������� �������� ���� ��� �������  ���������
'Parameters:
'[IN]   InstanceID , ��� ���������: String  - ID ���������
'Returns:
'  �������� ���� Long
'See Also:
'  CloseClass
'  EraseForm
'  GuiPath
'  INIT
'  IsModal
'  Show
'  ShowAddForm
'Example:
' dim variable as Long
' variable = me.FormCount(...���������...)
Public Function FormCount(ByVal InstanceID As String) As Long
Attribute FormCount.VB_HelpID = 765
    Dim f As OpenFormInfo
    Dim cnt As Long
    cnt = 0
    For Each f In openForms
      If f.InstanceID = InstanceID Then
        cnt = cnt + 1
      End If
    Next
    FormCount = cnt + STDTypeGUI.FormCount(InstanceID)
End Function


'��������  ���� �������������� ��������� �� ������
'Parameters:
'[IN]   InstanceID , ��� ���������: String,
'[IN]   mode , ��� ���������: String,
'[IN]   ModalMode , ��� ���������: Boolean  - ...
'See Also:
'  CloseClass
'  FormCount
'  GuiPath
'  INIT
'  IsModal
'  Show
'  ShowAddForm
'Example:
'  call me.EraseForm(...���������...)
Public Sub EraseForm(ByVal InstanceID As String, ByVal mode As String, ByVal ModalMode As Boolean)
Attribute EraseForm.VB_HelpID = 760
    On Error Resume Next
    Dim f As OpenFormInfo
    Dim i As Long, cnt As Long
    i = 1
    cnt = FormCount(InstanceID)
    For Each f In openForms
      If f.InstanceID = InstanceID And f.mode = mode And f.Modal = ModalMode Then
        GoTo found
      End If
      i = i + 1
    Next
    STDTypeGUI.EraseForm InstanceID, mode, ModalMode
    Exit Sub
found:
    
    openForms.Remove i
    If Not f.Form Is Nothing Then
    With f.Form
      If Not .Item Is Nothing Then
         .Item.Manager.FreeInstanceObject .Item.id
      End If
    End With
    Set f.Form.Host = Nothing
    End If
    Set f.Form = Nothing
End Sub

'������� �����������
'Parameters:
' ���������� ���
'Returns:
' Boolean, ��������� ����������:
'   true  -
'   false -
'See Also:
'  CloseClass
'  EraseForm
'  FormCount
'  GuiPath
'  INIT
'  Show
'  ShowAddForm
'Example:
' dim variable as Boolean
' variable = me.IsModal
Public Property Get IsModal() As Boolean
Attribute IsModal.VB_HelpID = 780
 IsModal = mOpenModal
End Property


'�������� ����� �������������� ���������
'Parameters:
'[IN]   mode , ��� ���������: String - �����,
'[IN][OUT]   ObjItem , ��� ���������: Object -  Application ����� ���������,
'[IN][OUT]   Optional OpenModal , ��� ���������: Boolean = False  - ��������� ����� ��������
'See Also:
'  CloseClass
'  EraseForm
'  FormCount
'  GuiPath
'  INIT
'  IsModal
'  ShowAddForm
'Example:
'  call me.Show(...���������...)
Public Sub Show(ByVal mode As String, ObjItem As Object, Optional OpenModal As Boolean = False)
Attribute Show.VB_HelpID = 785
 Dim f As Form
 On Error Resume Next
 mOpenModal = OpenModal
If Not OpenModal Then
  If ObjItem.MTZSession.CheckRight(ObjItem.SecureStyleID, m_Typename & ":" & mode & ":" & ObjItem.StatusID) Then
   Set opf = FindForm(ObjItem.id, mode)
   If opf Is Nothing Then
     Set f = New frmObj
     Set opf = New OpenFormInfo
     Set opf.Form = f
     opf.InstanceID = ObjItem.id
     opf.mode = mode
     opf.Modal = OpenModal
     openForms.Add opf
   End If
   Set f = opf.Form
   Set f.Host = Me
   f.INIT ObjItem, mode
   f.Show
   f.WindowState = 0
   f.ZOrder 0
  Else
   MsgBox "������ � ��������� ��������", vbOKOnly + vbCritical, ""
  End If
End If
If OpenModal Then
  STDTypeGUI.Show mode, ObjItem, OpenModal
End If
End Sub

' ���� � ���������� GUI ���������
'Parameters:
' ���������� ���
'Returns:
'  �������� ���� String
'See Also:
'  CloseClass
'  EraseForm
'  FormCount
'  INIT
'  IsModal
'  Show
'  ShowAddForm
'Example:
' dim variable as String
'  variable = me.GuiPath()
Public Function GuiPath() As String
Attribute GuiPath.VB_HelpID = 770
On Error GoTo bye
GuiPath = STDTypeGUI.GuiPath
Exit Function
bye:
GuiPath = App.Path & "\" & App.EXEName & ".exe"
End Function