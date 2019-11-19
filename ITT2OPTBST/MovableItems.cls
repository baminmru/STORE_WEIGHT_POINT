VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MovableItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MOvableItem"
Attribute VB_Ext_KEY = "Member0" ,"MOvableItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' коллекция перемещаемых паллет


'local variable to hold Collection
Private mCol As Collection

Public Function Add(ByVal PalCode As String, Optional ByVal GoodParam As String = "", Optional ByVal palletSize As Double = 1, Optional ByVal PalType As String = "E") As MOvableItem
    'create a new object
    
    
    Dim objNewMember As MOvableItem
    
    On Error Resume Next
    
    Set objNewMember = Item(PalCode)
    If objNewMember Is Nothing Then
    
      Set objNewMember = New MOvableItem
      mCol.Add objNewMember, PalCode
    End If
    'set the properties passed into the method
    objNewMember.GoodParam = GoodParam
    objNewMember.palletSize = palletSize
    objNewMember.PalType = PalType
    objNewMember.PalCode = PalCode
    

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As MOvableItem
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the Collection
    'vntIndexKey contains either the Index or Key to the Collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol.Item(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'Collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the Collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub





Private Sub Class_Initialize()
    'creates the Collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys Collection when this class is terminated
    Set mCol = Nothing
End Sub



