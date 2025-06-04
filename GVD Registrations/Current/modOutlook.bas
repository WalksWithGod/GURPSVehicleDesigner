Attribute VB_Name = "modOutlook"
Option Explicit

'**************************************
' Name: A better example of Automation w
'     ith Outlook
' Description:This *<*improved*>*
'     demonstration program gives examples of
'     how you can control Outlook using AUTOMA
'     TION to create mail, contacts, folders a
'     nd appointments.
'It shows how To create contacts In a specific group of your choosing which serves as a good example of folder manipulation in Outlook.
'You can adapt this code To create the other outlook items.


'Don 't forget that you can find out more by reading the Outlook VBA documents or by pressing F1 in front of options you don't understand.
' By: John Edward Colman
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/xq/ASP/txtCode
'     Id.12124/lngWId.1/qx/vb/scripts/ShowCode
'     .htm'for details.'**************************************


'Create an object to refererence the Out
'     look App.
'This is simular to a pointer and is dec
'     lared in this way...
'...to allow early binding, making the c
'     ode more efficient.
Private o1 As Outlook.Application


Private Function CreateContactsFolder(Name As String) As Outlook.MAPIFolder
    On Error Resume Next
    'You can use a similar syntax to create
    '     folders in any part of outlook.
    ' (MAPI is the only valid string)
    Set CreateContactsFolder = o1.GetNamespace("MAPI").GetDefaultFolder(olFolderContacts).Folders.Add(Name)
    'Note: To create further sub folders:
    'Set CreateContactsFolder = o1.GetNamesp
    '     ace("MAPI").GetDefaultFolder(olFolderCon
    '     tacts).Folders("parent folder").Add(Name
    '     )
End Function


Private Sub Form_Load()
    'Create an instance of Outlook
    Set o1 = New Outlook.Application
End Sub


Private Sub Form_Terminate()
    'Comment out this line if you don't want
    '     to close Outlook
    o1.Quit
    'The next line frees up the memory used
    Set o1 = Nothing
End Sub


Private Sub CreateEmail(Recipient As String, Subject As String, Body As String, Attach As String)
    'Create a reference to a mail item
    Dim e1 As Outlook.MailItem
    'Create a new mail item
    Set e1 = o1.CreateItem(olMailItem)
    'Set a few of the many possible message
    '     parameters.
    e1.To = Recipient
    e1.Subject = Subject
    e1.Body = Body
    'This is how you add attatchments


    If Attach <> vbNullString Then
        e1.Attachments.Add Path
    End If
    'Commit the message
    e1.Send
    'Free up the space
    Set e1 = Nothing
End Sub
'This routine improved:
'-Now there is an option to create the c
'     ontact in a particular folder.
'-Returns a reference to the contact ite
'     m


Private Function CreateContact(Name As String, Nick As String, Email As String, Optional Folder As Outlook.MAPIFolder = Nothing) As Outlook.ContactItem
    'Create a new contact item


    If Folder Is Nothing Then
        Set CreateContact = o1.CreateItem(olContactItem)
    Else
        Set CreateContact = Folder.Items.Add(olContactItem)
    End If
    'Set a few of the many possible contact
    '     parameters.
    CreateContact.FullName = Name
    CreateContact.NickName = Nick
    CreateContact.Email1Address = Email
    'Commit the contact
    CreateContact.Save
End Function


Private Sub CreateAppointment(StartTime As Date, Endtime As Date, Subject As String, Location As String)
    'Create a reference to a Appointment ite
    '     m
    Dim e1 As Outlook.AppointmentItem
    'Create a new appointment item
    Set e1 = o1.CreateItem(olAppointmentItem)
    'Set a few of the many possible appointm
    '     ent parameters.
    e1.Start = StartTime
    e1.End = Endtime
    e1.Subject = Subject
    e1.Location = Location
    'If you want to set a list of recipients
    '     , do it like this
    'e1.Recipients.Add Name
    'Commit the appointment
    e1.Send
    'Free up the space
    Set e1 = Nothing
End Sub


Private Sub Command1_Click()
    'This variable holds a reference to a fo
    '     lder
    Dim f1 As Outlook.MAPIFolder
    Dim e1 As Outlook.ContactItem
    Dim e2 As Outlook.ContactItem
    'Use one of the following methods to cre
    '     ate a folder and put contacts in it
    Set f1 = CreateContactsFolder("My Friends")
    CreateContact "john", "johnzinho", "j@k", f1
    ' Or....
    CreateContactsFolder ("My Family")
    CreateContact "Dad", "My dad", "d@k", o1.GetNamespace("MAPI").GetDefaultFolder(olFolderContacts).Folders("My Family")
    ' Or....
    Set f1 = CreateContactsFolder("Business Contacts")
    Set e1 = CreateContact("John", "Big John", "jc@hot.com")
    'Make a copy to put in a different list
    Set e2 = e1.Copy
    e2.Move f1
    'The following line opens the folder for
    '     a user
    f1.Display
End Sub

