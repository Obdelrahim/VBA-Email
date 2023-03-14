Attribute VB_Name = "Business_Return"
Option Explicit

Sub Business_Returns()

    Dim oApp As Outlook.Application
    Set oApp = New Outlook.Application
    Dim oMail As Outlook.MailItem
    Set oMail = oApp.CreateItemFromTemplate("C:\Users\omer\AppData\Roaming\Microsoft\Templates\xxx - 2022 Business Income Tax Return(s) - Morris  Kim LLP.oft")

    oMail.Display
    
    Dim attachSubject As String
     attachSubject = Sheet1.Range("U5").Value

    Dim emTo As String, emAttach As String, emSubject As String
    
    emTo = Sheet1.Range("S90").Value
    emAttach = Sheet1.Range("T90").Value
    emSubject = Sheet1.Range("A90").Value & "-"
    
    With oMail
     .To = emTo
     .Subject = emSubject & attachSubject
     .Attachments.Add emAttach
     .CC = "richard@morrisandkim.com"
    End With

    
    Set oApp = Nothing
    Set oMail = Nothing
End Sub
