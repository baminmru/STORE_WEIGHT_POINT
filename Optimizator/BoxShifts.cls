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
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"BoxShift"
Attribute VB_Ext_KEY = "Member0" ,"BoxShift"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold SortableCollection
Private mCol As SortableCollection

Public Function Add(ToCode As String, FromCode As String) As BoxShift
    'create a new object
    Dim objNewMember As BoxShift
    Set objNewMember = New BoxShift


    'set the properties passed into the method
    objNewMember.ToCode = ToCode
    objNewMember.FromCode = FromCode
    mCol.AddItem objNewMember
  


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As BoxShift
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the SortableCollection
    'vntIndexKey contains either the Index or Key to the SortableCollection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol.Item(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'SortableCollection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
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

Public Function Base() As SortableCollection
  Set Base = mCol
End Function

