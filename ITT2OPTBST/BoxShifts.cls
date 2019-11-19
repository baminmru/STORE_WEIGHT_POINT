VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BoxShifts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 135
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"BoxShift"
Attribute VB_Ext_KEY = "Member0" ,"BoxShift"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Коллекция Перемещений

Private mCol As SortableCollection

'Добавить в коллекцию перемещение
'Parameters:
'[IN][OUT]  ToCode , тип параметра: String,
'[IN][OUT]   FromCode , тип параметра: String  - ...
'Returns:
'  объект класса BoxShift
'  ,или Nothing
'See Also:
'  Base
'  Count
'  Item
'  Remove
'Example:
' dim variable as BoxShift
' Set variable = me.Add(...параметры...)
Public Function Add(ToCode As String, FromCode As String, pallet As String, GoodInfo As String, ptype As String, Stage As String) As BoxShift
Attribute Add.VB_HelpID = 140
    'create a new object
    Dim objNewMember As BoxShift
    Set objNewMember = New BoxShift

  

    'set the properties passed into the method
    objNewMember.ToCode = ToCode
    objNewMember.FromCode = FromCode
    objNewMember.pallet = pallet
    objNewMember.GoodInfo = GoodInfo
    objNewMember.ptype = ptype
    objNewMember.Stage = Stage
    mCol.AddItem objNewMember
  


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

'Получить перемещение по индексу, или номеру
'Parameters:
'[IN][OUT]  vntIndexKey , тип параметра: Variant  - ...
'Returns:
'  объект класса BoxShift
'  ,или Nothing
'See Also:
'  Add
'  Base
'  Count
'  Remove
'Example:
' dim variable as BoxShift
' Set variable = me.Item(...параметры...)
Public Property Get Item(vntIndexKey As Variant) As BoxShift
Attribute Item.VB_HelpID = 155
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    'used when referencing an element in the SortableCollection
    'vntIndexKey contains either the Index or Key to the SortableCollection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol.Item(vntIndexKey)
End Property

'Количество элементов в коллекции
'Parameters:
' параметров нет
'Returns:
'  значение типа Long
'See Also:
'  Add
'  Base
'  Item
'  Remove
'Example:
' dim variable as Long
' variable = me.Count
Public Property Get Count() As Long
Attribute Count.VB_HelpID = 150
    'used when retrieving the number of elements in the
    'SortableCollection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

'Удалить элемент
'Parameters:
'[IN][OUT]  vntIndexKey , тип параметра: Variant  - ...
'See Also:
'  Add
'  Base
'  Count
'  Item
'Example:
'  call me.Remove(...параметры...)
Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_HelpID = 160
    'used when removing an element from the SortableCollection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub




Private Sub Class_Initialize()
    'creates the SortableCollection when this class is created
    Set mCol = New SortableCollection
End Sub


Private Sub Class_Terminate()
    'destroys SortableCollection when this class is terminated
    Set mCol = Nothing
End Sub

'Возвращает элемент базового класса
'Parameters:
' параметров нет
'Returns:
'  объект класса SortableCollection
'  ,или Nothing
'See Also:
'  Add
'  Count
'  Item
'  Remove
'Example:
' dim variable as SortableCollection
' Set variable = me.Base()
Public Function Base() As SortableCollection
Attribute Base.VB_HelpID = 145
  Set Base = mCol
End Function

