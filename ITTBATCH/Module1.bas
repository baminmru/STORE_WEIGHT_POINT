Attribute VB_Name = "Module1"
Option Explicit

Public Sub Main()
  Dim s As Object
  Set s = CreateObject("ITTBATCH.SYNC")
  If Not s Is Nothing Then
  s.setup
  End If
  
  If MsgBox("Запустить тестовый прогон синхронизации?", vbYesNo, "Тест") = vbYes Then
    s.BeforeSync
    s.Aftersync
    
  End If
  Set s = Nothing

End Sub
