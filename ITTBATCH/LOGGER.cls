VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LOGGER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 105
Option Explicit
'отладочные сообщения

Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Private Type LogRecord
  App As String
  Message As String
  Time As Date
End Type

Private ff As Integer
Private mAppName As String
Private mFileName As String

'Записать сообщение
'Parameters:
'[IN]   s , тип параметра: String  - сообщение
'See Also:
'  CreateLog
'Example:
'  call me.Message(<параметры>)
Public Sub Message(ByVal s As String)
Attribute Message.VB_HelpID = 115
  If GetSetting("ITT", "ITTBATCH", "LOG", "1") = "1" Then
    Dim r As LogRecord
    r.App = mAppName
    r.Time = Now
    r.Message = s
    On Error Resume Next
    ff = FreeFile
    Open mFileName For Append As #ff Len = Len(r)
    Print #ff, r.Time, r.App, r.Message
    Close #ff
    
    Dim s1 As String
    s1 = Now & " LOG: " & s & Chr(0)
    
    OutputDebugString s1
  End If
  Debug.Print s
  
End Sub

'создать новый лог
'Parameters:
'[IN]   Name , тип параметра: String,
'[IN]   AppName , тип параметра: String  - ...
'See Also:
'  Message
'Example:
'  call me.CreateLog(<параметры>)
Public Sub CreateLog(ByVal Name As String, ByVal AppName As String)
Attribute CreateLog.VB_HelpID = 110
  On Error Resume Next
  Dim r As LogRecord
  
  mFileName = App.Path & "\" & Name
  mAppName = AppName
End Sub

Private Sub Class_Initialize()
  ff = -1
  mAppName = "LOGGER"
  mFileName = App.Path & "\log.txt"
End Sub


