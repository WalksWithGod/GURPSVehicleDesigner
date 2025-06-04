Option Strict Off
Option Explicit On
Module modMapi
	
	'Mapi Control constancts
	Private Const mapSuccessSuccess As Short = 32000 'Action returned successfully
	Private Const mapUserAbort As Short = 32001 'User canceled process
	Private Const mapFailure As Short = 32002 'Unspecified failure
	Private Const mapLoginFail As Short = 32003 'Login failure
	Private Const mapDiskFull As Short = 32004 ' Disk full
	Private Const mapInsufficientMem As Short = 32005 'Insufficient memory
	Private Const mapAccessDenied As Short = 32006 'Access denied
	Private Const mapGeneralFailure As Short = 32007 'General failure
	Private Const mapTooManySessions As Short = 32008 'Too many sessions
	Private Const mapTooManyFiles As Short = 32009 'Too many files
	Private Const mapTooManyRecipients As Short = 32010 'Too many recipients
	Private Const mapAttachmentNotFound As Short = 32011 'Attachment not found
	Private Const mapAttachmentOpenFailure As Short = 32012 'Attachment open failure
	Private Const mapAttachmentWriteFailure As Short = 32013 'Attachment write failure
	Private Const mapUnknownRecipient As Short = 32014 'Unknown recipient
	Private Const mapBadRecipType As Short = 32015 'Invalid recipient type
	Private Const mapNoMessages As Short = 32016 'No message
	Private Const mapInvalidMessage As Short = 32017 'Invalid message
	Private Const mapTextTooLarge As Short = 32018 'Text too large
	Private Const mapInvalidSession As Short = 32019 'Invalid session
	Private Const mapTypeNotSupported As Short = 32020 'Type not supported
	Private Const mapAmbiguousRecipient As Short = 32021 'Ambiguous recipient
	Private Const mapMessageInUse As Short = 32022 'Message in use
	Private Const mapNetworkFailure As Short = 32023 'Network failure
	Private Const mapInvalidEditFields As Short = 32024 'Invalid editfields
	Private Const mapInvalidRecips As Short = 32025 'Invalid Recipients
	Private Const mapNotSupported As Short = 32026 'Current action not supported
	Private Const mapSessionExist As Short = 32050 'Session ID already exists
	Private Const mapInvalidBuffer As Short = 32051 'Read-only in read buffer
	Private Const mapInvalidReadBufferAction As Short = 32052 'Valid in compose buffer only
	Private Const mapNoSession As Short = 32053 'No valid session ID
	Private Const mapInvalidRecipient As Short = 32054 'Originator information not available
	Private Const mapInvalidComposeBufferAction As Short = 32055 'Action not valid for Compose Buffer
	Private Const mapControlFailure As Short = 32056 'No messages in list
	Private Const mapNoRecipients As Short = 32057 'No recipients
	Private Const mapNoAttachment As Short = 32058 'attachments
	
	
	Sub Publish()
		Dim frmDesigner As Object
		Dim Vehicle As Object
		Dim sVehicleFile As Object
		On Error GoTo errorhandler
		Dim tempbyte() As Byte
		Dim bFlag As Boolean
		Dim i As Integer
		Dim sName As String
		Dim AttachName1 As String
		Dim AttachName2 As String
		Dim AttachPath1 As String
		Dim AttachPath2 As String
		Dim iAttachCount As Short
		'UPGRADE_ISSUE: File object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim fsoFile As File
		'UPGRADE_ISSUE: FileSystemObject object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim fso As FileSystemObject
		Dim sMessageSubject As String
		Const MAX_IMAGESIZE_IN_BYTES As Short = 25000
		
		
		'do a reg check
		'//reg check
		tempbyte = VB6.CopyArray(ChopCheck)
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If (IsNothing(tempbyte) = False) And (IsNothing(gsRegNum) = False) And (UBound(tempbyte) - LBound(tempbyte) = UBound(gsRegNum) - LBound(gsRegNum)) Then
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
			MsgBox("This program has not been properly registered.  The 'Publish' feature is only available in Registered versions.")
			Exit Sub
			'//check if there's even a file loaded
			'UPGRADE_WARNING: Couldn't resolve default property of object sVehicleFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf sVehicleFile = "untitled" Then 
			MsgBox("You must save the Vehicle before it can be published.")
			Exit Sub
		Else
			'//now check that the Vehicle's Name, Version, Author, Category and Subcategory fields have been filled in
			'UPGRADE_WARNING: Couldn't resolve default property of object Vehicle.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (Vehicle.Description.NickName = "") Or (Vehicle.Description.author = "") Or (Vehicle.Description.version = 0) Or (Vehicle.Description.category = "") Or (Vehicle.Description.subcategory = "") Then
				MsgBox("You must fill in the 'Name', 'Category', 'SubCategory' and 'Author' fields in the Options dialog before you can publish this vehicle.")
				Exit Sub
			End If
			'//proceed to process the attached files
			'UPGRADE_WARNING: Couldn't resolve default property of object Vehicle.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Vehicle.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sName = "TL" & Vehicle.Components(BODY_KEY).TL & " " & Vehicle.Description.NickName & " v" & Vehicle.Description.version & " -" & Vehicle.Description.author
			'UPGRADE_WARNING: Couldn't resolve default property of object Vehicle.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sMessageSubject = "VEHICLE" & "|" & Vehicle.Description.category & "|" & Vehicle.Description.subcategory
			AttachName1 = sName & ".veh"
			'UPGRADE_WARNING: Couldn't resolve default property of object sVehicleFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AttachPath1 = sVehicleFile
			
			'//check to see if their is an image of the file
			'UPGRADE_WARNING: Couldn't resolve default property of object Vehicle.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Vehicle.Description.VehicleImageFileName <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Vehicle.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AttachPath2 = My.Application.Info.DirectoryPath & "\images\" & Vehicle.Description.VehicleImageFileName
				'//see if the image still exists
				fso = New FileSystemObject
				'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				fsoFile = fso.GetFile(AttachPath2)
				'//see if it fits our max file size for uploading
				'UPGRADE_WARNING: Couldn't resolve default property of object fsoFile.size. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If fsoFile.size <= MAX_IMAGESIZE_IN_BYTES Then
					'//rename our image attachment to match the first attachment except
					'  use a JPG extension
					AttachName2 = sName & ".jpg"
				Else
					AttachPath2 = ""
				End If
			End If
			System.Windows.Forms.Application.DoEvents()
			
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPISession1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			With frmDesigner.MAPISession1
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPISession1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.LogonUI = True
				'.NewSession = True
				'.SessionID = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPISession1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.SignOn()
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPISession1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Tag = "on"
			End With
			System.Windows.Forms.Application.DoEvents()
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			With frmDesigner.MAPIMessages1
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPISession1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.SessionID = frmDesigner.MAPISession1.SessionID
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Compose()
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.RecipAddress = Settings.PublishEmailAddress
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.MsgSubject = sMessageSubject
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.MsgNoteText = "SUBMIT VEHICLE"
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.AttachmentIndex = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.AttachmentName = AttachName1
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.AttachmentPathName = AttachPath1
				If AttachPath2 <> "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.AttachmentIndex = 1
					'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.AttachmentName = AttachName2
					'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.AttachmentPathName = AttachPath2
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.MsgReceiptRequested = True
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPIMessages1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.send()
			End With
		End If
errorhandler: 
		
		Select Case Err.Number
			Case 0
			Case mapSuccessSuccess 'Action returned successfully
			Case mapUserAbort 'User canceled process
			Case mapFailure 'Unspecified failure
			Case mapLoginFail 'Login failure
			Case mapDiskFull ' Disk full
			Case mapInsufficientMem 'Insufficient memory
			Case mapAccessDenied 'Access denied
			Case mapGeneralFailure 'General failure
			Case mapTooManySessions 'Too many sessions
			Case mapTooManyFiles 'Too many files
			Case mapTooManyRecipients 'Too many recipients
			Case mapAttachmentNotFound 'Attachment not found
			Case mapAttachmentOpenFailure 'Attachment open failure
			Case mapAttachmentWriteFailure 'Attachment write failure
			Case mapUnknownRecipient 'Unknown recipient
			Case mapBadRecipType 'Invalid recipient type
			Case mapNoMessages 'No message
			Case mapInvalidMessage 'Invalid message
			Case mapTextTooLarge 'Text too large
			Case mapInvalidSession 'Invalid session
			Case mapTypeNotSupported 'Type not supported
			Case mapAmbiguousRecipient 'Ambiguous recipient
			Case mapMessageInUse 'Message in use
			Case mapNetworkFailure 'Network failure
			Case mapInvalidEditFields 'Invalid editfields
			Case mapInvalidRecips 'Invalid Recipients
			Case mapNotSupported 'Current action not supported
			Case mapSessionExist 'Session ID already exists
			Case mapInvalidBuffer 'Read-only in read buffer
			Case mapInvalidReadBufferAction 'Valid in compose buffer only
			Case mapNoSession 'No valid session ID
				MsgBox("Could not establish a MAPI session. Check that you are currently logged into your internet service provider.")
				Exit Sub
			Case mapInvalidRecipient 'Originator information not available
			Case mapInvalidComposeBufferAction 'Action not valid for Compose Buffer
			Case mapControlFailure 'No messages in list
			Case mapNoRecipients 'No recipients
			Case mapNoAttachment
				
			Case Else
				MsgBox("Could not establish a MAPI session. The 'publish' feature requires that you have a MAPI enabled email client installed.")
		End Select
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPISession1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If frmDesigner.MAPISession1.Tag = "on" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPISession1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmDesigner.MAPISession1.SignOff()
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MAPISession1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmDesigner.MAPISession1.Tag = "off"
		End If
		
	End Sub
End Module