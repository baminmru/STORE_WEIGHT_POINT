Attribute VB_Name = "mailSupport"
Attribute VB_HelpID = 1265
Option Explicit

'Подготовка почтового сообщения
'Parameters:
'[IN]   aSubj , тип параметра: String - тема,
'[IN]   aBody , тип параметра: String - тело,
'[IN]   FileName , тип параметра: String  - файл для присоединения
'Example:
'  call me.MailThisFile(...параметры...)
Public Sub MailThisFile(ByVal aSubj As String, ByVal aBody As String, ByVal FileName As String)
Attribute MailThisFile.VB_HelpID = 1270
    Dim mail As STDMail.Application
    Dim i As Long
    Dim idmail As String
    idmail = CreateGUID2()
    Manager.NewInstance idmail, "STDMail", "Оповещение " & Now
    Set mail = Manager.GetInstanceObject(idmail)
    If Not mail Is Nothing Then
    
      For i = 1 To ITTDic.ITTD_EMAIL.Count
        If ITTDic.ITTD_EMAIL.Item(i).IgnoreAddress = Boolean_Net Then
          With mail.STDMail_To.Add
            .TheTo = ITTDic.ITTD_EMAIL.Item(i).EMail
            .TheType = MailSenderType_Komu
            .save
          End With
        End If
      Next
      
      With mail.STDMail_Attach.Add
        .TheFile = FileToArray(FileName)
        .TheFile_EXT = GetFileExtension2(FileName)
        .TheName = "report." & GetFileExtension2(FileName)
        .save
      End With
      
      
      With mail.STDMail_Info.Add
        .Subject = aSubj
        .TheBody = aBody
        .Sended = Boolean_Net
        .save
      End With
      
    End If
End Sub
