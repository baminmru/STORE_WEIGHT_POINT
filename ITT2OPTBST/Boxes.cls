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
Attribute VB_Exposed = True
Attribute VB_HelpID = 85
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Box"
Attribute VB_Ext_KEY = "Member0" ,"Box"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'��������� �����


'local variable to hold SortableCollection
Private mCol As SortableCollection

'�������� ������� � ���������
'Parameters:
'[IN][OUT]  qcur , ��� ���������: Integer,
'[IN][OUT]   qmax , ��� ���������: Integer,
'[IN][OUT]   T , ��� ���������: Integer,
'[IN][OUT]   Z , ��� ���������: Integer,
'[IN][OUT]   Y , ��� ���������: Integer,
'[IN][OUT]   X , ��� ���������: Integer,
'[IN][OUT]   Code , ��� ���������: String
' ��. Box
'Returns:
'  ������ ������ Box
'  ,��� Nothing
'See Also:
'  Base
'  Count
'  Item
'  Remove
'Example:
' dim variable as Box
' Set variable = me.Add(...���������...)
Public Function Add(qcur As Double, QMAX As Integer, T As Integer, Z As Integer, Y As Integer, X As Integer, Code As String, GoodInfo As String) As Box
Attribute Add.VB_HelpID = 90
    'create a new object
    Dim objNewMember As Box
    Set objNewMember = New Box


    'set the properties passed into the method
    objNewMember.qcur = qcur
    objNewMember.QMAX = QMAX
   
    objNewMember.T = T
    objNewMember.Z = Z
    objNewMember.Y = Y
    objNewMember.X = X
    objNewMember.Code = Code
    objNewMember.GoodInfo = GoodInfo
  
    mCol.AddItem objNewMember, Code
   


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

'�������� ������� ��������� �� �������
'Parameters:
'[IN][OUT]  vntIndexKey , ��� ���������: Variant  - ...
'Returns:
'  ������ ������ Box
'  ,��� Nothing
'See Also:
'  Add
'  Base
'  Count
'  Remove
'Example:
' dim variable as Box
' Set variable = me.Item(...���������...)
Public Property Get Item(vntIndexKey As Variant) As Box
Attribute Item.VB_HelpID = 105
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    'used when referencing an element in the SortableCollection
    'vntIndexKey contains either the Index or Key to the SortableCollection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error Resume Next
    Set Item = mCol.Item(vntIndexKey)
End Property

'���������� ���������
'Parameters:
' ���������� ���
'Returns:
'  �������� ���� Long
'See Also:
'  Add
'  Base
'  Item
'  Remove
'Example:
' dim variable as Long
' variable = me.Count
Public Property Get Count() As Long
Attribute Count.VB_HelpID = 100
    'used when retrieving the number of elements in the
    'SortableCollection. Syntax: Debug.Print x.Count
      On Error Resume Next
    Count = mCol.Count
End Property

'�������� ��������
'Parameters:
'[IN][OUT]  vntIndexKey , ��� ���������: Variant  - ...
'See Also:
'  Add
'  Base
'  Count
'  Item
'Example:
'  call me.Remove(...���������...)
Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_HelpID = 110
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

'������� ����� ���������
'Parameters:
' ���������� ���
'Returns:
'  ������ ������ SortableCollection
'  ,��� Nothing
'See Also:
'  Add
'  Count
'  Item
'  Remove
'Example:
' dim variable as SortableCollection
' Set variable = me.Base()
Public Function Base() As SortableCollection
Attribute Base.VB_HelpID = 95
  Set Base = mCol
End Function