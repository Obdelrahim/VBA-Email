Attribute VB_Name = "PersonalRefund"
Option Explicit

Sub PersonalRefund()

    Dim oApp As Outlook.Application
    Set oApp = New Outlook.Application
    Dim oMail As Outlook.MailItem
        Set oMail = oApp.CreateItemFromTemplate("template location")

    oMail.Display
    
    Dim attachSubject As String
            attachSubject = Sheet1.Range("Cell1").Value

    Dim emTo As String, emAttach As String, emSubject As String
    
            emTo = Sheet1.Range("Cell2").Value
            emAttach = Sheet1.Range("Cell3").Value
            emSubject = Sheet1.Range("Cell4").Value & "-"
    
    With oMail
     .To = emTo
     .Subject = emSubject & attachSubject
     .Attachments.Add emAttach
                .CC = "email@example.com"
    End With

    
    Set oApp = Nothing
    Set oMail = Nothing
End Sub
