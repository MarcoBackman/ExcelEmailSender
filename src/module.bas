Type smptDataType
    smtpServer As String
    smtpPort As Integer
    smtpUser As String
    smtpPassword As String
End Type

Public smptData As smptDataType
Public subject As String
Public body As String



Public Sub setSmptData()
    Sheets("Sender detail").Select
    ' SMPT info
    smptData.smtpServer = Sheets("Sender detail").Range("E4").Value
    Debug.Print "smptData.smtpServer: " & smptData.smtpServer
    smptData.smtpUser = Sheets("Sender detail").Range("E5").Value
    smptData.smtpPassword = Sheets("Sender detail").Range("E6").Value
    smptData.smtpPort = Sheets("Sender detail").Range("E7").Value
End Sub

Sub setMailContent(name As String, position As String, company As String)
    
    subject = "안녕하세요! " & company & " " & position & " " & name & "님"
    body = "<html><body>" & _
          "<font face=""verdana, sans-serif"" style=""font-family: verdana, sans-serif; color: rgb(55, 55, 55);"">" & _
          "<strong>" & company & " " & position & " " & name & "님 안녕하세요!</strong><br><br>" & _
          "content.<br>" & _
          "content.<br>" & _
          "content.<br><br>" & _
          "content.<br><br>" & _
          "content!<br>" & _
          "content.<br><br>" & _
          "content<br><br>" & _
          "</font>" & _
          "<hr style=""height:2px;border-width:0;color:gray;background-color:gray""><br><br>" & _
          "<img src=""img_url""><br><br>" & _
          "<font face=""verdana, sans-serif"" style=""font-family: verdana, sans-serif; color: rgb(102, 102, 102);"">" & _
            "개발자:  <br>" & _
            "연락처:  <br>" & _
            "이메일:  <br>" & _
            "주소:  <br>" & _
          "</font>" & _
          "</body></html>"

    ' MsgBox content
End Sub

Public Sub SendMailTo(num As Integer)

    Dim sender As String
    Dim bodyFormat As Integer
    Dim recipientName As String
    Dim recipientAddress As String
    Dim position As String
    Dim companyName As String
    
    ' User data - set recipient
    Workbooks("BusinesscardFormatter.xlsm").Activate
    Sheets("Contact list").Select
    recipientName = Sheets("Contact list").Range("A" & num).Value
    position = Sheets("Contact list").Range("B" & num).Value
    companyName = Sheets("Contact list").Range("C" & num).Value
    recipientAddress = Sheets("Contact list").Range("D" & num).Value
    
    
    ' Mail content
    Call setMailContent(recipientName, position, companyName)
    
    
    ' Sender info
    sender = Sheets("Sender detail").Range("E8")
    

    Debug.Print "recipientName: " & recipientName
    Debug.Print "recipientAddress: " & recipientAddress
    
    bodyFormat = 1
    
    Debug.Print "body: " & body

    Set oSmtp = New EASendMailObjLib.Mail
    oSmtp.LicenseCode = "TryIt" ' Here goes your license code for the software; for now, we are using the trial version

    ' Please change the server address, username, and password to the ones you will be using
    oSmtp.ServerAddr = smptData.smtpServer
    oSmtp.UserName = smptData.smtpUser
    oSmtp.Password = smptData.smtpPassword
    oSmtp.ServerPort = smptData.smtpPort
    
    Debug.Print "ServerAddr: " & oSmtp.ServerAddr
    Debug.Print "UserName: " & oSmtp.UserName
    Debug.Print "Password: " & oSmtp.Password
    Debug.Print "ServerPort: " & oSmtp.ServerPort

    ' Using TryTLS,
    ' If the SMTP server supports TLS, then a TLS connection is used; otherwise, a normal TCP connection is used.
    ' https://www.emailarchitect.net/easendmail/sdk/?ct=connecttype
    oSmtp.ConnectType = 3

    oSmtp.FromAddr = sender
    oSmtp.AddRecipient recipientName, recipientAddress, 0

    oSmtp.subject = subject
    oSmtp.bodyFormat = bodyFormat
    oSmtp.BodyText = body
    
    
    ' file attachment
    oSmtp.AddAttachment "file_path"

    oSmtp.Asynchronous = 0
    oSmtp.SendMail
    Set oSmtp = Nothing
    
End Sub

Sub Main()
    Call setSmptData
    Dim i As Integer
    For i = 2 To 5
        Call SendMailTo(i)
    Next i
End Sub
