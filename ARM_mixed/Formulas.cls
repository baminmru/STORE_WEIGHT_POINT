VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Formulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 75
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"FormulaData"
Attribute VB_Ext_KEY = "Member0" ,"FormulaData"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Коллекция формул в отчете


'local variable to hold collection
Private mCol As Collection

'Добавить в коллекцию
'Parameters:
'[IN]   Name , тип параметра: String  - ...
'Returns:
'  объект класса FormulaData
'  ,или Nothing
'See Also:
'  Count
'  Item
'  NewEnum
'  Remove
'Example:
' dim variable as FormulaData
' Set variable = me.Add(...параметры...)
Public Function Add(ByVal Name As String) As FormulaData
Attribute Add.VB_HelpID = 80
    'create a new object
    Dim objNewMember As FormulaData
    Set objNewMember = New FormulaData
    objNewMember.Name = Name

    
    mCol.Add objNewMember, Name
    


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

'Получить по имени
'Parameters:
'[IN][OUT]  vntIndexKey , тип параметра: Variant  - ...
'Returns:
'  объект класса FormulaData
'  ,или Nothing
'See Also:
'  Add
'  Count
'  NewEnum
'  Remove
'Example:
' dim variable as FormulaData
' Set variable = me.Item(...параметры...)
Public Property Get Item(vntIndexKey As Variant) As FormulaData
Attribute Item.VB_HelpID = 90
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  On Error Resume Next
  Set Item = mCol(vntIndexKey)
End Property


'Количество элементов
'Parameters:
' параметров нет
'Returns:
'  значение типа Long
'See Also:
'  Add
'  Item
'  NewEnum
'  Remove
'Example:
' dim variable as Long
' variable = me.Count
Public Property Get Count() As Long
Attribute Count.VB_HelpID = 85
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


' Удалить из коллекци
'Parameters:
'[IN][OUT]  vntIndexKey , тип параметра: Variant  - ...
'See Also:
'  Add
'  Count
'  Item
'  NewEnum
'Example:
'  call me.Remove(...параметры...)
Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_HelpID = 100
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_HelpID = 95
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

