Imports Microsoft.Exchange.WebServices.Data

Public Class Form1
    Private Function RedirectionUrlValidationCallback(Optional redirectionUrl = "") As Boolean

        Dim result As Boolean = False

        Dim redirectionUri = New Uri(redirectionUrl)

        If redirectionUri.Scheme = "https" Then
            result = True
        End If

        Return result
    End Function


    Private Function SendExchangeEmail(sender_address As String, sender_password As String,
                                       email_recipient As String, email_subject As String,
                                       email_body As String, email_attachments As String) As Boolean
        Notify("Establishing Exchange Connection...")

        Dim service As New ExchangeService(ExchangeVersion.Exchange2007_SP1)
        service.Url = New Uri("https://mail.h01.hostedmail.net/EWS/Exchange.asmx")
        service.Credentials = New WebCredentials(sender_address, sender_password)
        service.PreAuthenticate = True


        'service.AutodiscoverUrl(sender_address, AddressOf RedirectionUrlValidationCallback)

        Notify("Connected to " & service.Url.ToString())

        Dim email As New EmailMessage(service)
        email.ToRecipients.Add(email_recipient)
        email.Subject = email_subject
        email.Body = New MessageBody(Microsoft.Exchange.WebServices.Data.BodyType.Text, email_body)

        'Load up attachments if present -- update 20190211 djb
        If email_attachments <> "|" Then
            Dim attachments As String() = Split(email_attachments, "|")
            For Each filename As String In attachments
                If filename <> "" Then
                    email.Attachments.AddFileAttachment(filename)
                End If
            Next

        End If


        Notify("Sending Email to '" & email_recipient & "'")


        Dim err_string As String = ""
        Dim result As String = ""
        Dim retVal As Boolean = False

        Try
            email.SendAndSaveCopy()
        Catch ex As Microsoft.Exchange.WebServices.Data.ServiceRequestException
            err_string = "Cannot connect to Exchange Webservice defined by " &
                service.Url.ToString &
                "Exact Error: " & ex.Message

        Catch ex As Exception
            err_string = "General Error: " & ex.Message
        End Try

        If err_string <> "" Then
            result = "Error: " & err_string
            retVal = False
        Else
            result = "Email Sent OK!"
            retVal = True
        End If

        Notify(result)

        Return retVal

    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Interval = 200
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False

        'build stuff from command line.
        'Command Line Example:
        'ExchangeEmailSender.exe username password recipient "subject" "message-body-filename"
        Dim login As String = Environment.GetCommandLineArgs(1)
        Dim password As String = Environment.GetCommandLineArgs(2)
        Dim recipient As String = Environment.GetCommandLineArgs(3)
        Dim subject As String = Environment.GetCommandLineArgs(4)
        Dim body_filename As String = Environment.GetCommandLineArgs(5)
        Dim email_attachments As String = Environment.GetCommandLineArgs(6)

        Dim body As String = ""

        For Each line As String In IO.File.ReadLines(body_filename)
            body &= line & vbNewLine
        Next

        If SendExchangeEmail(login, password, recipient, subject, body, email_attachments) Then
            End
        End If
    End Sub

    Private Sub Notify(message As String)
        TextBox1.Text &= message & vbNewLine
        TextBox1.Refresh()
    End Sub

End Class
