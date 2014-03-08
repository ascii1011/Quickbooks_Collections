Attribute VB_Name = "modEmail"
Option Explicit



Function SendOCEmail(SendToAddress As String, _
                    SentFromEmail As String, _
                    ProductDescription As String, _
                    ProductFile As String, _
                    EmailSubject As String) As Boolean
    
    Dim cdo2configuration As New CDO.Configuration
    Dim cdo2fields As ADODB.Fields
    Dim cdo2message As New CDO.Message
    Dim strFile As String

On Error GoTo errhandler

    SendOCEmail = False

    Set cdo2fields = cdo2configuration.Fields
    
    'prcReportLocal Now & " Setting up Email Params..."
    With cdo2fields
        .Item(cdoSMTPServer) = "192.168.168.28"
        .Item(cdoSendUsingMethod) = cdoSendUsingPort
        .Update
    End With
                
    Set cdo2message.Configuration = cdo2configuration

    
    cdo2message.From = SentFromEmail
    cdo2message.To = SendToAddress
    cdo2message.Subject = EmailSubject
    
    If ProductFile <> "0" Then  'if equal to zero then order copy and no file and no description
        'strFile = ProductFile
            'strFile = "c:\MCsvrLog.txt"
        'cdo2message.AddAttachment strFile
    Else
        'cdo2message.CC = SentFromEmail
    End If
    
    cdo2message.HTMLBody = ProductDescription
        
    
    cdo2message.Send        'sending email now
    
    SendOCEmail = True
    
    Exit Function
    
errhandler:
    Select Case (Err.Number)
        Case Else
            MsgBox Err.Number & ": " & Err.Description & vbNewLine & "There was a problem with this email process."
    End Select
End Function
