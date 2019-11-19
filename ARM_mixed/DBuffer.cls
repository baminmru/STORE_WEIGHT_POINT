VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 20
Option Explicit

' объект для хранения временных данных
Public DateOf As Date
Attribute DateOf.VB_VarHelpID = 25
Public Item As Object
Attribute Item.VB_VarHelpID = 35
Public id As String
Attribute id.VB_VarHelpID = 30
Public Name As String
Attribute Name.VB_VarHelpID = 40
Public OK As Boolean
Attribute OK.VB_VarHelpID = 45
Public Volume As Double
Attribute Volume.VB_VarHelpID = 55
Public ordItem As Object
Attribute ordItem.VB_VarHelpID = 50
