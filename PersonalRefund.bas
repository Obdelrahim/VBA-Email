Attribute VB_Name = "PersonalRefund"
Option Explicit

Sub PersonalRefund()

    Dim oApp As Outlook.Application
    Set oApp = New Outlook.Application
    Dim oMail As Outlook.MailItem
    Set oMail = oApp.CreateItemFromTemplate("C:\Users\omer\OneDrive\Documents\Refund Income Tax Return Template.oft")

    oMail.Display
    
    Dim attachSubject As String
     attachSubject = Sheet1.Range("U4").Value

    Dim emTo As String, emAttach As String, emSubject As String
    
    emTo = Sheet1.Range("S79").Value
    emAttach = Sheet1.Range("T79").Value
    emSubject = Sheet1.Range("A79").Value & "-"
    
    With oMail
     .To = emTo
     .Subject = emSubject & attachSubject
     .Attachments.Add emAttach
     .CC = "richard@morrisandkim.com"
    End With

    
    Set oApp = Nothing
    Set oMail = Nothing
End Sub
