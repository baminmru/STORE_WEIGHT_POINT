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
' класс подмены GUI интерфейса документов для  реализации отображения документов внутри  MDI окна


Private mOpenModal As Boolean
Dim opf As OpenFormInfo
Dim openForms As Collection
Private m_Typename As String
Private STDTypeGUI As Object

' Инициализация класса для конкретного типа документа
'Parameters:
'[IN]   tn , тип параметра: String  - тип документа
'See Also:
'  CloseClass
'  EraseForm
'  FormCount
'  GuiPath
'  IsModal
'  Show
'  ShowAddForm
'Example:
'  call me.INIT(...параметры...)
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

' Открытие формы редактирования
'Parameters:
'[IN]   mode , тип параметра: String,
'[IN][OUT]   RowItem , тип параметра: Object  - ...
'Returns:
' Boolean, семантика результата:
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
' variable = me.ShowAddForm(...параметры...)
Public Function ShowAddForm(ByVal mode As String, RowItem As Object) As Boolean
Attribute ShowAddForm.VB_HelpID = 790
On Error Resume Next
ShowAddForm = STDTypeGUI.ShowAddForm(mode, RowItem)
End Function

' Освобождение ссылок, выгрузка открытых и загруженных форм
'Parameters:
' параметров нет
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


' поиск формы редактирования в коллекции загруженных форм
Private Function FindForm(ByVal InstanceID As String, ByVal mode As String) As Object
  Dim f As OpenFormInfo
  For Each f In openForms
    If f.InstanceID = InstanceID And f.mode = mode And f.Modal = False Then
      Set FindForm = f
      Exit Function
    End If
  Next
End Function

'количество открытых форм для данного  документа
'Parameters:
'[IN]   InstanceID , тип параметра: String  - ID документа
'Returns:
'  значение типа Long
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
' variable = me.FormCount(...параметры...)
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


'выгрузка  форм редактирования документа из памяти
'Parameters:
'[IN]   InstanceID , тип параметра: String,
'[IN]   mode , тип параметра: String,
'[IN]   ModalMode , тип параметра: Boolean  - ...
'See Also:
'  CloseClass
'  FormCount
'  GuiPath
'  INIT
'  IsModal
'  Show
'  ShowAddForm
'Example:
'  call me.EraseForm(...параметры...)
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

'Признак модальности
'Parameters:
' параметров нет
'Returns:
' Boolean, семантика результата:
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


'Показать форму редактирования документа
'Parameters:
'[IN]   mode , тип параметра: String - режим,
'[IN][OUT]   ObjItem , тип параметра: Object -  Application класс документа,
'[IN][OUT]   Optional OpenModal , тип параметра: Boolean = False  - модальный режим открытия
'See Also:
'  CloseClass
'  EraseForm
'  FormCount
'  GuiPath
'  INIT
'  IsModal
'  ShowAddForm
'Example:
'  call me.Show(...параметры...)
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
   MsgBox "Доступ к документу запрещен", vbOKOnly + vbCritical, ""
  End If
End If
If OpenModal Then
  STDTypeGUI.Show mode, ObjItem, OpenModal
End If
End Sub

' путь к библиотеке GUI документа
'Parameters:
' параметров нет
'Returns:
'  значение типа String
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
