Attribute VB_Name = "modMapi"

'Mapi Control constancts
Private Const mapSuccessSuccess = 32000 'Action returned successfully
Private Const mapUserAbort = 32001 'User canceled process
Private Const mapFailure = 32002 'Unspecified failure
Private Const mapLoginFail = 32003 'Login failure
Private Const mapDiskFull = 32004 ' Disk full
Private Const mapInsufficientMem = 32005 'Insufficient memory
Private Const mapAccessDenied = 32006 'Access denied
Private Const mapGeneralFailure = 32007 'General failure
Private Const mapTooManySessions = 32008 'Too many sessions
Private Const mapTooManyFiles = 32009 'Too many files
Private Const mapTooManyRecipients = 32010 'Too many recipients
Private Const mapAttachmentNotFound = 32011 'Attachment not found
Private Const mapAttachmentOpenFailure = 32012 'Attachment open failure
Private Const mapAttachmentWriteFailure = 32013 'Attachment write failure
Private Const mapUnknownRecipient = 32014 'Unknown recipient
Private Const mapBadRecipType = 32015 'Invalid recipient type
Private Const mapNoMessages = 32016 'No message
Private Const mapInvalidMessage = 32017 'Invalid message
Private Const mapTextTooLarge = 32018 'Text too large
Private Const mapInvalidSession = 32019 'Invalid session
Private Const mapTypeNotSupported = 32020 'Type not supported
Private Const mapAmbiguousRecipient = 32021 'Ambiguous recipient
Private Const mapMessageInUse = 32022 'Message in use
Private Const mapNetworkFailure = 32023 'Network failure
Private Const mapInvalidEditFields = 32024 'Invalid editfields
Private Const mapInvalidRecips = 32025 'Invalid Recipients
Private Const mapNotSupported = 32026 'Current action not supported
Private Const mapSessionExist = 32050 'Session ID already exists
Private Const mapInvalidBuffer = 32051 'Read-only in read buffer
Private Const mapInvalidReadBufferAction = 32052 'Valid in compose buffer only
Private Const mapNoSession = 32053 'No valid session ID
Private Const mapInvalidRecipient = 32054 'Originator information not available
Private Const mapInvalidComposeBufferAction = 32055 'Action not valid for Compose Buffer
Private Const mapControlFailure = 32056 'No messages in list
Private Const mapNoRecipients = 32057 'No recipients
Private Const mapNoAttachment = 32058  'attachments


Sub Publish()
    On Error GoTo errorhandler
    Dim tempbyte() As Byte
    Dim bFlag As Boolean
    Dim i As Long
    Dim sName As String
    Dim AttachName1 As String
    Dim AttachName2 As String
    Dim AttachPath1 As String
    Dim AttachPath2 As String
    Dim iAttachCount As Integer
    Dim fsoFile As File
    Dim fso As FileSystemObject
    Dim sMessageSubject As String
    Const MAX_IMAGESIZE_IN_BYTES = 25000
        
    
    'do a reg check
    '//reg check
    tempbyte = ChopCheck
    If (IsEmpty(tempbyte) = False) And (IsEmpty(gsRegNum) = False) And (UBound(tempbyte) - LBound(tempbyte) = UBound(gsRegNum) - LBound(gsRegNum)) Then
        For i = 1 To UBound(gsRegNum)
            If tempbyte(i) = gsRegNum(i) Then
                bFlag = True
            Else
                bFlag = False
                Exit For
            End If
        Next
    Else
        bFlag = False
    End If
    If bFlag = False Then
        MsgBox "This program has not been properly registered.  The 'Publish' feature is only available in Registered versions."
        Exit Sub
    '//check if there's even a file loaded
    ElseIf sVehicleFile = "untitled" Then
        MsgBox "You must save the Vehicle before it can be published."
        Exit Sub
    Else
        '//now check that the Vehicle's Name, Version, Author, Category and Subcategory fields have been filled in
        If (Vehicle.Description.NickName = "") Or (Vehicle.Description.author = "") Or (Vehicle.Description.version = 0) Or (Vehicle.Description.category = "") Or (Vehicle.Description.subcategory = "") Then
            MsgBox "You must fill in the 'Name', 'Category', 'SubCategory' and 'Author' fields in the Options dialog before you can publish this vehicle."
            Exit Sub
        End If
        '//proceed to process the attached files
        sName = "TL" & Vehicle.Components(BODY_KEY).TL & " " & Vehicle.Description.NickName & " v" & Vehicle.Description.version & " -" & Vehicle.Description.author
        sMessageSubject = "VEHICLE" & "|" & Vehicle.Description.category & "|" & Vehicle.Description.subcategory
        AttachName1 = sName & ".veh"
        AttachPath1 = sVehicleFile
        
        '//check to see if their is an image of the file
        If Vehicle.Description.VehicleImageFileName <> "" Then
            AttachPath2 = App.Path & "\images\" & Vehicle.Description.VehicleImageFileName
            '//see if the image still exists
            Set fso = New FileSystemObject
            Set fsoFile = fso.GetFile(AttachPath2)
                '//see if it fits our max file size for uploading
                If fsoFile.size <= MAX_IMAGESIZE_IN_BYTES Then
                    '//rename our image attachment to match the first attachment except
                    '  use a JPG extension
                    AttachName2 = sName & ".jpg"
                Else
                    AttachPath2 = ""
                End If
        End If
        DoEvents

        With frmDesigner.MAPISession1
            .LogonUI = True
            '.NewSession = True
            '.SessionID = 1
            .SignOn
            .Tag = "on"
        End With
        DoEvents
        With frmDesigner.MAPIMessages1
            .SessionID = frmDesigner.MAPISession1.SessionID
            .Compose
            .RecipAddress = Settings.PublishEmailAddress
            .MsgSubject = sMessageSubject
            .MsgNoteText = "SUBMIT VEHICLE"
            .AttachmentIndex = 0
            .AttachmentName = AttachName1
            .AttachmentPathName = AttachPath1
            If AttachPath2 <> "" Then
                .AttachmentIndex = 1
                .AttachmentName = AttachName2
                .AttachmentPathName = AttachPath2
            End If
            .MsgReceiptRequested = True
            .send
        End With
    End If
errorhandler:

    Select Case err.Number
        Case 0
        Case mapSuccessSuccess  'Action returned successfully
        Case mapUserAbort  'User canceled process
        Case mapFailure  'Unspecified failure
        Case mapLoginFail  'Login failure
        Case mapDiskFull  ' Disk full
        Case mapInsufficientMem  'Insufficient memory
        Case mapAccessDenied  'Access denied
        Case mapGeneralFailure  'General failure
        Case mapTooManySessions  'Too many sessions
        Case mapTooManyFiles 'Too many files
        Case mapTooManyRecipients  'Too many recipients
        Case mapAttachmentNotFound  'Attachment not found
        Case mapAttachmentOpenFailure  'Attachment open failure
        Case mapAttachmentWriteFailure  'Attachment write failure
        Case mapUnknownRecipient  'Unknown recipient
        Case mapBadRecipType  'Invalid recipient type
        Case mapNoMessages 'No message
        Case mapInvalidMessage  'Invalid message
        Case mapTextTooLarge  'Text too large
        Case mapInvalidSession  'Invalid session
        Case mapTypeNotSupported 'Type not supported
        Case mapAmbiguousRecipient  'Ambiguous recipient
        Case mapMessageInUse  'Message in use
        Case mapNetworkFailure  'Network failure
        Case mapInvalidEditFields  'Invalid editfields
        Case mapInvalidRecips  'Invalid Recipients
        Case mapNotSupported  'Current action not supported
        Case mapSessionExist 'Session ID already exists
        Case mapInvalidBuffer  'Read-only in read buffer
        Case mapInvalidReadBufferAction  'Valid in compose buffer only
        Case mapNoSession  'No valid session ID
            MsgBox "Could not establish a MAPI session. Check that you are currently logged into your internet service provider."
            Exit Sub
        Case mapInvalidRecipient 'Originator information not available
        Case mapInvalidComposeBufferAction 'Action not valid for Compose Buffer
        Case mapControlFailure  'No messages in list
        Case mapNoRecipients  'No recipients
        Case mapNoAttachment
        
        Case Else
            MsgBox "Could not establish a MAPI session. The 'publish' feature requires that you have a MAPI enabled email client installed."
    End Select

   If frmDesigner.MAPISession1.Tag = "on" Then
        frmDesigner.MAPISession1.SignOff
        frmDesigner.MAPISession1.Tag = "off"
    End If
    
End Sub


