VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Boxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Box"
Attribute VB_Ext_KEY = "Member0" ,"Box"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold SortableCollection
Private mCol As SortableCollection

Public Function Add(qcur As Integer, qmax As Integer, T As Integer, Z As Integer, Y As Integer, X As Integer, Code As String) As Box
    'create a new object
    Dim objNewMember As Box
    Set objNewMember = New Box


    'set the properties passed into the method
    objNewMember.qcur = qcur
    objNewMember.qmax = qmax
    objNewMember.T = T
    objNewMember.Z = Z
    objNewMember.Y = Y
    objNewMember.X = X
    objNewMember.Code = Code
  
    mCol.AddItem objNewMember, Code
   


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Box
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the SortableCollection
    'vntIndexKey contains either the Index or Key to the SortableCollection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error Resume Next
    Set Item = mCol.Item(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'SortableCollection. Syntax: Debug.Print x.Count
      On Error Resume Next
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
      On Error Resume Next
    Set mCol = New SortableCollection
End Sub


Private Sub Class_Terminate()
    'destroys SortableCollection when this class is terminated
      On Error Resume Next
    Set mCol = Nothing
End Sub

Public Function Base() As SortableCollection
  Set Base = mCol
End Function
