VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 20
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Ячейка


' Код
Public Code As String
Attribute Code.VB_VarHelpID = 25

'Координата Х
Public X As Integer
Attribute X.VB_VarHelpID = 70

'Координата Y
Public Y As Integer
Attribute Y.VB_VarHelpID = 75

'Координата Z
Public Z As Integer
Attribute Z.VB_VarHelpID = 80

Public T As Integer
Attribute T.VB_VarHelpID = 65

' максимальный объем
Public QMAX As Integer
Attribute QMAX.VB_VarHelpID = 60

'текущий объем
Public qcur As Double
Attribute qcur.VB_VarHelpID = 55

' Разрешено брать из ячейки
Public GetFrom As Boolean
Attribute GetFrom.VB_VarHelpID = 35

' Разрешено ставить в яччейку
Public PutTo As Boolean
Attribute PutTo.VB_VarHelpID = 50

'  В ячейку было размещение
Public HavePut As Boolean
Attribute HavePut.VB_VarHelpID = 40


'  В ячейку запрещено размещение !!!
Public HaveMixedPallet As Boolean

Public GoodInfo As String



Private m_pallets As MovableItems

'Список оптимизируемых паллет
Public Property Get Pallets() As MovableItems
  If m_pallets Is Nothing Then
    Set m_pallets = New MovableItems
  End If
  Set Pallets = m_pallets
End Property


'Вычислить стоимость изъятия
'Parameters:
'[IN]   locl , тип параметра: Boolean  - ...
'Returns:
'  значение типа Double
'See Also:
'  Code
'  GetFrom
'  HavePut
'  PutCost
'  PutTo
'  qcur
'  qmax
'  T
'  X
'  Y
'  Z
'Example:
' dim variable as Double
' variable = me.GetCost(...параметры...)
Public Function GetCost(ByVal locl As Boolean) As Double
Attribute GetCost.VB_HelpID = 30
  Dim r As Double
  If locl Then
    r = (Z - 1) / 2 + 1 + (Y - 1) / 100
  Else
    r = 1 + (Z - 1) / 2 + 1 + (Y - 1) / 100
  End If
  GetCost = r
End Function

' Вычислить стоимость постановки
'Parameters:
'[IN]   locl , тип параметра: Boolean  - ...
'Returns:
'  значение типа Double
'See Also:
'  Code
'  GetCost
'  GetFrom
'  HavePut
'  PutTo
'  qcur
'  qmax
'  T
'  X
'  Y
'  Z
'Example:
' dim variable as Double
' variable = me.PutCost(...параметры...)
Public Function PutCost(ByVal locl As Boolean) As Double
Attribute PutCost.VB_HelpID = 45
  Dim r As Double
If locl Then
    r = (Z - 1) / 2 + 1 + (Y - 1) / 100
  Else
    r = 1 + (Z - 1) / 2 + 1 + (Y - 1) / 100
  End If
  PutCost = r
End Function



Private Sub Class_Terminate()
  Set m_pallets = Nothing
End Sub
