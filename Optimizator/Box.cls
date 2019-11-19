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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Code As String
Public X As Integer
Public Y As Integer
Public Z As Integer
Public T As Integer

Public qmax As Integer
Public qcur As Integer

Public GetFrom As Boolean
Public PutTo As Boolean
Public HavePut As Boolean


Public Function GetCost(ByVal locl As Boolean) As Double
  Dim r As Double
  If locl Then
    r = (Z - 1) / 2 + 1 + (Y - 1) / 100
  Else
    r = 1 + (Z - 1) / 2 + 1 + (Y - 1) / 100
  End If
  GetCost = r
End Function

Public Function PutCost(ByVal locl As Boolean) As Double
  Dim r As Double
If locl Then
    r = (Z - 1) / 2 + 1 + (Y - 1) / 100
  Else
    r = 1 + (Z - 1) / 2 + 1 + (Y - 1) / 100
  End If
  PutCost = r
End Function



