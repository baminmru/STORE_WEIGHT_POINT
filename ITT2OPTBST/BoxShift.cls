VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BoxShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 115
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Перемещение

'Откуда
Public FromCode As String
Attribute FromCode.VB_VarHelpID = 125

'Стадия алгоритма
Public Stage As String

'Куда
Public ToCode As String
Attribute ToCode.VB_VarHelpID = 130

' Исключить из коллекции
Public ExcludeShift As Boolean
Attribute ExcludeShift.VB_VarHelpID = 120
  
  
Public pallet As String

Public ptype As String
  
  
Public GoodInfo As String
