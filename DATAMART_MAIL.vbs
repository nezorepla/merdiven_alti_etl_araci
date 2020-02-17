Set OApp = CreateObject("Outlook.Application")
Set OMail = OApp.CreateItem(0)
        signature = OMail.body
    With OMail
    .To = "alper@mail.com" 
    .Subject = "Datamart Daily Refresh" 
    .body = "Merhaba,"&chr(10)&"Datamart tabloları güncellenmiştir."&chr(10)&"##tablo  İyi çalışmalar. " & vbNewLine & signature 
  .Send
    End With
