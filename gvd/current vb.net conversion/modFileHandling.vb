Option Strict Off
Option Explicit On
Module modFileHandling
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef hpvDest As Any, ByRef hpvSource As Any, ByVal cbCopy As Integer)
	
	'======================================================
	' This is the UDT for the INI file
	'======================================================
	Private Structure udtRegInfo
		Dim RegName() As Byte
		Dim RegNum() As Byte
		Dim RegID As Integer
	End Structure
	
	'UPGRADE_WARNING: Arrays in structure RegInfo may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Private RegInfo As udtRegInfo
	'======================================================
	
	'======================================================
	' This is .VEH file  header information
	'======================================================
	Dim FormatSignature As Byte ' used with SIG_128 and SIG_129 to determine if this is a new file format and whether it has One or Two file headers (i.e. Header and Header2 UDT's)
	Const SIG_128 As Short = 10 ' only contains first header
	Const SIG_129 As Short = 11 ' contains first and second header
	Const OFFSET1 As Short = 14 'offset for start of second header
	Const OFFSET2 As Short = 1356 'offset for start of vehicle data
	Dim m_lngOffset As Integer
	
	' GVD 1st header data
	Private Structure Header
		Dim CRC32 As Integer
		Dim Major As Short
		Dim Minor As Short
		Dim Revision As Short
		Dim RegID As Short
	End Structure
	
	' 2nd header, contains user vehicle file info
	Private Structure Header2
		Dim TL As Byte
		Dim version As Single
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(39),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=39)> Public GUID() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(50),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=50)> Public Category() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(50),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=50)> Public subcategory() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(150),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=150)> Public Name() As Char ' vehicle name and not FILENAME
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		'UPGRADE_NOTE: Class was upgraded to Class_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		<VBFixedString(150),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=150)> Public Class_Renamed() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(100),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=100)> Public author() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(100),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=100)> Public email() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(200),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=200)> Public url() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(255),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=255)> Public jpgfilename() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(255),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=255)> Public Description() As Char 'max description that will be visible on site is only 255
	End Structure
	
	'======================================================
	' This is the UDT for the File Save/Load file
	'======================================================
	
	Private Structure uComponent
		Dim TreeInfo As String
		Dim Properties() As String
	End Structure
	
	Private Components() As uComponent
	'========================================================
	
	
	Const GVDLicenseFile As String = "GVD.lic"
	Const GVDINIFile As String = "GVD.ini"
	
	Const FLAG_NOZIP As Short = 0 ' set to 0 for RELEASE builds
	
	Private z As String
	
	Public Sub ExportFile(ByRef sType As String)
		Dim frmConfigure As Object
		Dim frmDesigner As Object
		' Code to export and view the file as either Text, HTML-classic gurps style or HTML-tables
		Dim Cancel As Boolean
		Dim sFileName As String
		Dim sExtension As String
		Dim sFilter As String
		Dim sTemp As String
		Dim oCDLG As clsCmdlg
		
		If sType = "Text" Or sType = "Text Slim" Then
			sExtension = ".txt"
			sFilter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
		Else
			sExtension = ".html"
			sFilter = "HTML files (*.htm; *.html)|*.htm; *.html|All files (*.*)|*.*"
		End If
		On Error GoTo errorhandler
		Cancel = False
		'UPGRADE_ISSUE: FileSystemObject object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oFile As FileSystemObject
		With oCDLG
			' todo: eventually use simpler code to check for existance?  dont really need to
			' include the scripting runtime object if i stop using the filesystemobjects
			oFile = New FileSystemObject
			
			If sType = "Text" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object oFile.FolderExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If oFile.FolderExists(Settings.TextExportPath) Then
					.InitialDir = Settings.TextExportPath
				Else
					.InitialDir = My.Application.Info.DirectoryPath
				End If
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object oFile.FolderExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If oFile.FolderExists(Settings.HTMLExportPath) Then
					.InitialDir = Settings.HTMLExportPath
				Else
					.InitialDir = My.Application.Info.DirectoryPath
				End If
			End If
			.DefaultFilename = ""
			'.DefaultExt = sExtension
			.Filter_Renamed = sFilter
			.CancelError = True
			.MultiSelect = False
			'.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
		End With
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.hwnd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Cancel = oCDLG.ShowSave(frmDesigner.hwnd)
		If Not Cancel Then
			' A fileName was selected. Add the code to save the file here
			'UPGRADE_WARNING: Couldn't resolve default property of object oCDLG.cFileName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sFileName = oCDLG.cFileName.Item(0)
			
			'save the path
			If sType = "Text" Then
				Settings.TextExportPath = ExtractPathFromFile(sFileName)
			Else
				Settings.HTMLExportPath = ExtractPathFromFile(sFileName)
			End If
			System.Windows.Forms.Application.DoEvents()
		Else
			Exit Sub
		End If
		
		' now generate the actual file
		FileOpen(2, sFileName, OpenMode.Output)
		sTemp = createGURPSText(sType)
		PrintLine(2, sTemp)
		FileClose(2)
		
		Dim retval As Integer
		Dim sProgramPath As String
		
		'make sure the path is set
		If sType = "Text" Or sType = "Text Slim" Then
			sProgramPath = Settings.TextViewerPath
		Else
			sProgramPath = Settings.HTMLBrowserPath
		End If
		
		If sProgramPath = "" Then
			MsgBox("No viewer specified.")
			'UPGRADE_WARNING: Couldn't resolve default property of object frmConfigure.Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmConfigure.Show(VB6.FormShowConstants.Modal, frmDesigner)
			'UPGRADE_NOTE: Object frmConfigure may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			frmConfigure = Nothing
			
		Else 'attempt to launch the program
			retval = StartDoc(sProgramPath)
			If retval <= 32 Then ' Error
				MsgBox("Web Page not Opened", MsgBoxStyle.Exclamation, "URL Failed")
			End If
		End If
		Exit Sub
		
		' check to see if the user Cancels out instead of electing to save a file
errorhandler: 
		modHelper.InfoPrint(1, "Error in ExportFile:  " & CStr(Err.Number) & " " & Err.Description)
		Resume Next
	End Sub
	
	Function LoadRecords(ByVal FileName As String) As Boolean
		Dim m_oCurrentVeh As Object
		Dim frmDesigner As Object
		Dim Vehicles As Object
		Dim Vehicle As Object
		' This function just loads the file data into an array of
		' "components".  They are not actually added to the "Vehicle" at this point
		' that is done in the RebuildComponentStructure function which is called at the end of this sub
		' if the vehicle data was able to be read in properly
		
		Dim sVersInfo() As String
		Const Major As Short = 0
		Const Minor As Short = 1
		Const Revision As Short = 2
		Const REG_ID As Short = 3
		
		Dim iFreeFile As Integer
		Dim sLine As String
		
		
		On Error GoTo errorhandler
		
		' Destroy any old vehicle object and create new
		'UPGRADE_NOTE: Object Vehicle may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Vehicle = Nothing
		Vehicle = New Vehicles.clsVehicle
		'todo: MUST use obtptr() of IComponent interface of Vehicle for Key in tree for Vehicle
		' Clear the treeview of any nodes that might already exist
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.treeVehicle.Nodes.Clear()
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.GetOverDrive. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		z = m_oCurrentVeh.GetOverDrive
		
		'//show the "loading" in the status bar
		With frmDesigner
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MousePointer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.MousePointer = System.Windows.Forms.Cursors.WaitCursor
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ListView1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.ListView1.MousePointer = System.Windows.Forms.Cursors.WaitCursor
		End With
		
		' set the status bar panels
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.StatusBar1.Panels(1).Text = "Reading file data  0%"
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ImageList1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.StatusBar1.Panels(1).Picture = frmDesigner.ImageList1.ListImages(11).Picture
		
		' determine whether we are dealing with a old file format or new
		If NewFileFormat(FileName) Then
			If LoadComponents_NewFormat(FileName) = False Then GoTo errorhandler
			Debug.Print("LoadRecords: GVD User Registration ID = " & gsRegID) 'MPJ 07/04/2000
		Else
			'open the file
			iFreeFile = FreeFile
			FileOpen(iFreeFile, FileName, OpenMode.Input)
			'get the first line which has the version info
			sLine = LineInput(iFreeFile)
			FileClose(iFreeFile)
			
			sLine = DecryptINI(DecryptINI(sLine, z), z & Str(5982)) 'the first line is double encrypted
			sVersInfo = Split(sLine, ",")
			gsMajor = sVersInfo(Major)
			gsMinor = sVersInfo(Minor)
			gsRevision = sVersInfo(Revision)
			
			If LoadComponents_OldFormat(FileName) = False Then GoTo errorhandler
		End If
		
		'//reset the status bar panels
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.StatusBar1.Panels(1).Text = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.StatusBar1.Panels(1).Picture = Nothing
		
		'//Rebuild Vehicle Structure
		If RebuildComponentStructure = False Then GoTo errorhandler
		
		With frmDesigner
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MousePointer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.MousePointer = System.Windows.Forms.Cursors.Default
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ListView1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.ListView1.MousePointer = System.Windows.Forms.Cursors.Default
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.treeVehicle.Nodes(BODY_KEY).Expanded = True
		End With
		
		'//load was successful
		LoadRecords = True
		Exit Function
		
errorhandler: 
		With frmDesigner
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MousePointer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.MousePointer = System.Windows.Forms.Cursors.Default
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ListView1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.ListView1.MousePointer = System.Windows.Forms.Cursors.Default
		End With
		
		MsgBox("Unable to load file.  Vehicle file is either invalid or corrupt.")
		LoadRecords = False
		FileClose(iFreeFile)
	End Function
	
	Function NewFileFormat(ByRef sFileName As String) As Boolean
		
		Dim iFree As Integer
		Dim bSig As Byte
		Dim uHeader As Header
		Dim uHeader2 As Header2
		
		On Error GoTo errorhandler
		
		iFree = FreeFile
		FileOpen(iFree, sFileName, OpenMode.Binary)
		
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(iFree, bSig)
		
		' returns True if bSig matches our Signature constant
		If (bSig <> SIG_128) And (bSig <> SIG_129) Then
			NewFileFormat = False
		Else
			'get the rest of the header
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileGet(iFree, uHeader, 2)
			With uHeader
				gsMajor = CStr(.Major)
				gsMinor = CStr(.Minor)
				gsRevision = CStr(.Revision)
			End With
			
			NewFileFormat = True
			
			' set the appropriate offset for where our actual Vehicle data starts
			If bSig = SIG_129 Then
				m_lngOffset = OFFSET2
				
				' get the second Header of the new file format so we can
				'obtain the GUID
				'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileGet(iFree, uHeader2, OFFSET1)
				p_sGUID.Value = uHeader2.GUID
				
			ElseIf bSig = SIG_128 Then 
				m_lngOffset = OFFSET1
			Else
				NewFileFormat = False
			End If
		End If
		
		FileClose(iFree)
		Exit Function
		
errorhandler: 
		NewFileFormat = False
		FileClose(iFree)
	End Function
	
	
	Function LoadComponents_NewFormat(ByVal sFileName As String) As Boolean
		Dim frmDesigner As Object
		'//accepts a filename and loads in all the components
		Dim iFree As Integer
		Dim sTemp As String
		Dim sArray() As String
		Dim b() As Byte
		
		Dim i As Integer
		Dim j As Integer
		Dim k As Integer
		Dim lngDataLen As Integer
		Dim lngBytesRead As Integer
		'UPGRADE_ISSUE: cZlib object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oZip As cZlib
		oZip = New cZlib
		
		
		
		On Error GoTo errorhandler
		
		'UPGRADE_WARNING: Lower bound of array uRet was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim uRet(1) As Object
		iFree = FreeFile
		
		
		FileOpen(iFree, sFileName, OpenMode.Binary)
		
		'determine the length of the data
		lngDataLen = FileLen(sFileName) - m_lngOffset
		'UPGRADE_WARNING: Lower bound of array b was changed from 0 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim b(lngDataLen - 1)
		
		'read it all in and decompress it
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(iFree, b, m_lngOffset)
		
		If FLAG_NOZIP <> True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oZip.UncompressB. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oZip.UncompressB(b)
		End If
		
		' convert to string and split this up into our seperate component lines
		'UPGRADE_ISSUE: Constant vbUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		sTemp = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(b), vbUnicode)
		
		sArray = Split(sTemp, Chr(254))
		
		
		k = 0
		Do While k <= UBound(sArray)
			
			'fill the variant array that we will be passing into the FileLoader class
			If Left(sArray(k), 1) = "[" Then
				'now remove the leading and trailing characters
				i = i + 1
				j = 0
				'UPGRADE_WARNING: Lower bound of array Components was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				ReDim Preserve Components(i)
				sArray(k) = Mid(sArray(k), 2, Len(sArray(k)) - 2)
				Components(i).TreeInfo = sArray(k)
				
			Else
				j = j + 1
				'UPGRADE_WARNING: Lower bound of array Components(i).Properties was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				ReDim Preserve Components(i).Properties(j)
				Components(i).Properties(j) = sArray(k)
			End If
			k = k + 1
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmDesigner.StatusBar1.Panels(1).Text = "Reading file data  " & CInt(k / UBound(sArray) * 100) & "%"
		Loop 
		
		
		FileClose(iFree)
		'UPGRADE_NOTE: Object oZip may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oZip = Nothing
		LoadComponents_NewFormat = True
		Debug.Print("LoadComponents_NewFormat: " & sTemp)
		Exit Function
		
errorhandler: 
		Debug.Print("LoadComponents_NewFormat: " & Err.Description)
		'UPGRADE_NOTE: Object oZip may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oZip = Nothing
		LoadComponents_NewFormat = False
	End Function
	
	Function LoadComponents_OldFormat(ByVal sFileName As String) As Boolean
		Dim frmDesigner As Object
		'//accepts a filename and loads in all the components
		Dim iFree As Integer
		Dim sTemp As String
		Dim i As Integer
		Dim j As Integer
		Dim lngFileLen As Integer
		Dim lngBytesRead As Integer
		
		On Error GoTo errorhandler
		
		'UPGRADE_WARNING: Lower bound of array uRet was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim uRet(1) As Object
		iFree = FreeFile
		
		lngFileLen = FileLen(sFileName)
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.StatusBar1.Panels(1).Text = "Reading file data  0%"
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ImageList1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.StatusBar1.Panels(1).Picture = frmDesigner.ImageList1.ListImages(11).Picture
		
		FileOpen(iFree, sFileName, OpenMode.Input)
		
		'load the first line and skip it
		sTemp = LineInput(iFree)
		
		
		Do While Not EOF(iFree)
			sTemp = LineInput(iFree)
			lngBytesRead = lngBytesRead + Len(sTemp)
			sTemp = DecryptINI(sTemp, z)
			'fill the variant array that we will be passing into the FileLoader class
			If Left(sTemp, 1) = "[" Then
				'now remove the leading and trailing characters
				i = i + 1
				j = 0
				'UPGRADE_WARNING: Lower bound of array Components was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				ReDim Preserve Components(i)
				sTemp = Mid(sTemp, 2, Len(sTemp) - 2)
				Components(i).TreeInfo = sTemp
				
			Else
				j = j + 1
				'UPGRADE_WARNING: Lower bound of array Components(i).Properties was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				ReDim Preserve Components(i).Properties(j)
				Components(i).Properties(j) = sTemp
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmDesigner.StatusBar1.Panels(1).Text = "Reading file data  " & CInt(lngBytesRead / lngFileLen * 100) & "%"
		Loop 
		
		
		FileClose(iFree)
		
		LoadComponents_OldFormat = True
		Exit Function
		
errorhandler: 
		
		LoadComponents_OldFormat = False
		
	End Function
	
	Function RebuildComponentStructure() As Boolean
		Dim m_oCurrentVeh As Object
		Dim tvwChild As Object
		Dim frmDesigner As Object
		' This is where the read in saved vehicle data is turned into the actual Vehicle heirarchy.
		' It makes calls to Vehicle.AddObject for creating the correct objects based on the parsed datatypes
		Dim tobj As clsFileLoader
		Dim vc As Object
		Dim A As Integer
		Dim sKey As String
		Dim sParent As String
		Dim dType As Short
		Dim memberID As String
		Dim propvalue As Object
		Dim sDescription As String
		Dim icon1 As Short
		Dim arrkey As Object
		Dim i As Integer
		Dim j As Integer
		Dim iCount As Integer
		Dim iPropCount As Integer
		Dim lngUpper As Integer
		
		On Error GoTo errorhandler
		'/show our progress meter
		lngUpper = UBound(Components)
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.StatusBar1.Panels(1).Text = "Building tree  0%"
		
		' create an instance of the FileLoader class
		tobj = New clsFileLoader
		
		'//load up our tree and Component object
		For iCount = 1 To lngUpper
			'UPGRADE_WARNING: Couldn't resolve default property of object vc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vc = Split(Components(iCount).TreeInfo, "|")
			'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sDescription = vc(0)
			'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sKey = vc(1)
			'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sParent = vc(2)
			'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = vc(3)
			'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			icon1 = vc(4)
			
			'create the vehicle component object
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.addObject. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If m_oCurrentVeh.addObject(dType, sKey, sParent, icon1, sDescription, True) Then
				'create the tree node (unless its like a weapon link, performance, etc)
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				With frmDesigner.treeVehicle
					If sKey = BODY_KEY Then 'if its the body, then its the root node and doesnt have a parent
						'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Nodes.Add( ,  , sKey, sDescription, icon1)
					ElseIf (dType = PERFORMANCEPROFILE) Or (dType = WeaponLink) Then 
						'performance profiles and weapon links do NOT get added to the tree
						'todo: They will have to now!
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Nodes.Add(sParent, tvwChild, sKey, sDescription, icon1)
					End If
				End With
			Else
				GoTo errorhandler
			End If
			
			For iPropCount = 1 To UBound(Components(iCount).Properties)
				'UPGRADE_WARNING: Couldn't resolve default property of object vc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vc = Split(Components(iCount).Properties(iPropCount), "|")
				'check for keychain.
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If vc(1) = "[" Then
					
					j = 1
					'UPGRADE_WARNING: Lower bound of array arrkey was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
					ReDim arrkey(UBound(vc) - 1)
					For i = 2 To UBound(vc)
						'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object arrkey(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						arrkey(j) = vc(i)
						j = j + 1
					Next 
					'UPGRADE_WARNING: Couldn't resolve default property of object arrkey. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object propvalue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					propvalue = arrkey
					'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					memberID = vc(0)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					tobj.LetProperties(m_oCurrentVeh.Components(sKey), memberID, propvalue)
				Else
					'fill the properties for this object
					'UPGRADE_WARNING: Couldn't resolve default property of object vc(0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					System.Diagnostics.Debug.Assert(vc(0) <> "CombinedComponentVolume", "")
					'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					memberID = vc(0)
					'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object propvalue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					propvalue = vc(1)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					tobj.LetProperties(m_oCurrentVeh.Components(sKey), memberID, propvalue)
				End If
			Next 
			
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmDesigner.StatusBar1.Panels(1).Text = "Building tree  " & CInt(iCount / lngUpper * 100) & "%"
		Next 
		
		RebuildComponentStructure = True
		'destroy the instance of the fileloader class
		'UPGRADE_NOTE: Object tobj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		tobj = Nothing
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.StatusBar1.Panels(1).Picture = Nothing
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.StatusBar1.Panels(1).Text = ""
		Exit Function
		
errorhandler: 
		Debug.Print("RebuildComponentStructure: " & Err.Description)
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.StatusBar1.Panels(1).Text = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.StatusBar1.Panels(1).Picture = Nothing
		RebuildComponentStructure = False
		
	End Function
	'//////////////////////////////////////////////////////
	Function CreatePrintString(ByVal vc As Object) As String()
		
		Dim i As Object
		Dim j As Integer
		Dim dType As Short
		Dim sKey As String
		Dim sParent As String
		Dim skeychain As String
		Dim vType As Integer
		Dim svType As String
		Dim sDescription As String
		Dim icon1 As Short
		Dim icon2 As Short
		Dim retval() As String
		Dim SIZE As Integer
		
		'find the key and datatype
		For i = 1 To UBound(vc)
			'UPGRADE_WARNING: Couldn't resolve default property of object vc(i, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (vc(i, 0) = "Datatype") Or (vc(i, 0) = "datatype") Then
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dType = vc(0, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(i, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf (vc(i, 0) = "Key") Or (vc(i, 0) = "key") Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sKey = vc(0, i)
				If sKey = BODY_KEY Then sParent = "0_"
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(i, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf (vc(i, 0) = "Parent") Or (vc(i, 0) = "parent") Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sParent = vc(0, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(i, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf (vc(i, 0) = "customdescription") Or (vc(i, 0) = "Customdescription") Or (vc(i, 0) = "CustomDescription") Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sDescription = vc(0, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(i, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf (vc(i, 0) = "Image") Or (vc(i, 0) = "image") Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				icon1 = vc(0, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(i, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf (vc(i, 0) = "SelectedImage") Or (vc(i, 0) = "selectedimage") Or (vc(i, 0) = "Selectedimage") Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				icon2 = vc(0, i)
			End If
			If (dType <> VariantType.Empty) And (sKey <> "") And (sDescription <> "") And (icon1 <> VariantType.Empty) And (icon2 <> VariantType.Empty) And (sParent <> "") Then Exit For
		Next 
		
		'store this line in the array
		'UPGRADE_WARNING: Lower bound of array retval was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim retval(1)
		retval(1) = "[" & sDescription & "|" & sKey & "|" & sParent & "|" & Str(dType) & "|" & Str(icon1) & "|" & Str(icon2) & "]"
		
		'UPGRADE_WARNING: Lower bound of array retval was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim Preserve retval(UBound(vc) + 1)
		SIZE = 1
		
		For i = 1 To UBound(vc)
			SIZE = SIZE + 1
			'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			vType = VarType(vc(0, i))
			svType = Str(vType)
			'check for keychains (vbarray + vbvariant)
			'reset the skeychain string
			skeychain = ""
			If vType = VariantType.Array + VariantType.Object Then
				'found a keychain.  Place the "[" which inidcates a keychain
				skeychain = skeychain & "|" & "["
				'Seperate the individual keys into a string
				For j = 1 To UBound(vc(0, i))
					
					'UPGRADE_WARNING: Couldn't resolve default property of object vc()(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					skeychain = skeychain & "|" + vc(0, i)(j)
				Next 
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				retval(SIZE) = vc(i, 0) + skeychain
			ElseIf vType = VariantType.Array + VariantType.String Then 
				'found a keychain.  Place the "[" which inidcates a keychain
				skeychain = skeychain & "|" & "["
				'Seperate the individual keys into a string
				For j = 1 To UBound(vc(0, i))
					
					'UPGRADE_WARNING: Couldn't resolve default property of object vc()(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					skeychain = skeychain & "|" + vc(0, i)(j)
				Next 
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				retval(SIZE) = vc(i, 0) + skeychain
			ElseIf vType <> VariantType.String Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				retval(SIZE) = vc(i, 0) + "|" + Str(vc(0, i))
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(0, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				retval(SIZE) = vc(i, 0) + "|" + vc(0, i)
			End If
		Next 
		
		
		CreatePrintString = VB6.CopyArray(retval)
	End Function
	
	Sub CreateRecords()
		Dim Vehicle As Object
		Dim p_nIndex As Object
		Dim frmDesigner As Object
		Dim m_oCurrentVeh As Object
		
		Dim tobj As clsFileLoader
		Dim vc As Object
		Dim iIndex As Short
		Dim i As Integer
		Dim j As Integer
		Dim k As Integer
		Dim keychainkeys As Object
		Dim sKey As String
		
		On Error GoTo errorhandler
		
		' create an instance of the FileLoader class
		tobj = New clsFileLoader
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.GetOverDrive. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		z = m_oCurrentVeh.GetOverDrive
		
		'todo: Here instead of using GetFirstParent, we will start with the Body node and only
		' itterate thru all children/sub children from this node.
		'NOTE: We do have to itterate thru the tree and not simply itterate thru the
		' m_oCurrentVeh.components collection.  Itterating thru the components collection would result in
		'components being read and potentially the parent to which they must be added not being in the tree yet.
		'Proof: Lets say you add a Weapon to the Body and then a turret to the body.  Now lets say you move the weapon
		' to the turret.  That puts the weapon ahead of the Turret in the collection so when reading in the weapon
		' it would fail when trying to addobject to the parent turret which hasnt been installed yet.
		GetFirstParent() 'Find a root node in the treeview
		'get the index of the root node that is at the top of the treeview
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iIndex = frmDesigner.treeVehicle.Nodes(p_nIndex).FirstSibling.index
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sKey = frmDesigner.treeVehicle.Nodes(iIndex).Key
		' debug I really dont need a select case here since the first node is always the Body
		'sName = TypeName(m_oCurrentVeh.Components.item(sKey))
		'UPGRADE_WARNING: Couldn't resolve default property of object Vehicle.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object tobj.GetProperties(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object vc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		vc = tobj.GetProperties(Vehicle.Components(sKey))
		i = 1
		'UPGRADE_WARNING: Lower bound of array Components was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim Components(i)
		Components(1).Properties = CreatePrintString(vc)
		
		'If the Node has Children call the sub that writes the children
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If frmDesigner.treeVehicle.Nodes(iIndex).children > 0 Then
			WriteChild(iIndex)
		End If
		
		'Now save the Performance Profiles which are not
		'visually represented by the Tree
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object keychainkeys. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		keychainkeys = m_oCurrentVeh.Components(BODY_KEY).PerformanceProfileKeychain
		'UPGRADE_WARNING: Couldn't resolve default property of object keychainkeys(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (UBound(keychainkeys) >= 1) And keychainkeys(1) <> "" Then
			For k = 1 To UBound(keychainkeys)
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object tobj.GetProperties(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object vc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vc = tobj.GetProperties(m_oCurrentVeh.Components(keychainkeys(k)))
				i = UBound(Components) + 1
				'UPGRADE_WARNING: Lower bound of array Components was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				ReDim Preserve Components(i)
				Components(i).Properties = CreatePrintString(vc)
			Next 
		End If
		'Now save the Weapon Links which are not
		'visually represented by the Tree
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object keychainkeys. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		keychainkeys = m_oCurrentVeh.Components(BODY_KEY).WeaponLinkKeychain
		'UPGRADE_WARNING: Couldn't resolve default property of object keychainkeys(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (UBound(keychainkeys) >= 1) And keychainkeys(1) <> "" Then
			For k = 1 To UBound(keychainkeys)
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object tobj.GetProperties(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object vc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vc = tobj.GetProperties(m_oCurrentVeh.Components(keychainkeys(k)))
				i = UBound(Components) + 1
				'UPGRADE_WARNING: Lower bound of array Components was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				ReDim Preserve Components(i)
				Components(i).Properties = CreatePrintString(vc)
			Next 
		End If
		
		'destroy the instance of the fileloader class
		'UPGRADE_NOTE: Object tobj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		tobj = Nothing
		Exit Sub
		
errorhandler: 
		
		
	End Sub
	Sub WriteRecord(ByVal FileName As String)
		Dim frmDesigner As Object
		'now get ready to print the properties for this object
		'first delete the contents  of the file
		Dim iFree As Integer
		Dim i As Object
		Dim j As Integer
		Dim k As Integer
		Dim tempbyte() As Byte
		Dim bFlag As Boolean
		Dim lngUpper As Integer
		Dim s As String
		'UPGRADE_ISSUE: cZlib object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oZip As cZlib
		Dim b() As Byte
		Dim sJoin() As String
		Dim uHeader As Header
		Dim uHeader2 As Header2
		
		iFree = FreeFile
		
		'//reg check
		tempbyte = VB6.CopyArray(ChopCheck)
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If (IsNothing(tempbyte) = False) And (UBound(tempbyte) - LBound(tempbyte) = UBound(gsRegNum) - LBound(gsRegNum)) Then
			For i = 1 To UBound(gsRegNum)
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If tempbyte(i) = gsRegNum(i) Then
					bFlag = True
				Else
					bFlag = False
					GoTo reghandler
				End If
			Next 
		Else
			bFlag = False
			GoTo reghandler
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.StatusBar1.Panels(1).Text = "Reading vehicle data..."
		CreateRecords()
		
		'delete the existing file
		FileOpen(iFree, FileName, OpenMode.Random)
		FileClose(iFree)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.StatusBar1.Panels(1).Text = "Writing file  0%"
		'open the file for writing
		FileOpen(iFree, FileName, OpenMode.Binary)
		
		oZip = New cZlib
		
		lngUpper = UBound(Components)
		
		'UPGRADE_WARNING: Lower bound of array sJoin was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim sJoin(1)
		
		For i = 1 To lngUpper
			'Print #iFree, EncryptINI$(Components(i).TreeInfo, z)
			's = s & Components(i).TreeInfo & Chr(254)
			If k > 0 Then
				'UPGRADE_WARNING: Lower bound of array sJoin was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ReDim Preserve sJoin(UBound(sJoin) + UBound(Components(i).Properties))
			Else
				'UPGRADE_WARNING: Lower bound of array sJoin was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ReDim sJoin(UBound(Components(i).Properties))
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For j = 1 To UBound(Components(i).Properties)
				k = k + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sJoin(k) = Components(i).Properties(j)
				'Print #iFree, EncryptINI$(Components(i).Properties(j), z)
				
			Next 
			
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.StatusBar1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmDesigner.StatusBar1.Panels(1).Text = "Writing file  " & CInt(i / lngUpper * 100) & "%"
		Next 
		
		
		'ReDim Preserve B(Len(B) - 1)
		
		s = Join(sJoin, Chr(254))
		'remove the last chr(254) from the end
		's = Mid(s, 1, Len(s) - 1)
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		b = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(s, vbFromUnicode))
		
		
		If FLAG_NOZIP <> True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oZip.CompressB. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oZip.CompressB(b)
		End If
		
		'create our file header
		'first print the header version info
		With uHeader
			.CRC32 = 100 'todo: need to calc the crc first
			.Major = My.Application.Info.Version.Major
			.Minor = My.Application.Info.Version.Minor
			.Revision = My.Application.Info.Version.Revision
			.RegID = gsRegID
		End With
		
		' create our second file header
		'todo:
		'With m_oCurrentVeh.Components(BODY_KEY)
		'   uHeader2.TL = .TL
		'   uHeader2.version = m_oCurrentVeh.Description.version
		'   'uHeader2.GUID = 'todo: Eh?  why commented out.  Cant remember but i do need to store that GUID right?
		'   uHeader2.category = m_oCurrentVeh.Description.category
		'   uHeader2.subcategory = m_oCurrentVeh.Description.subcategory
		'   uHeader2.Class = m_oCurrentVeh.Description.Classname
		'   uHeader2.name = m_oCurrentVeh.Description.NickName
		'   uHeader2.author = m_oCurrentVeh.Description.author
		'   uHeader2.email = m_oCurrentVeh.Description.email
		'   uHeader2.url = m_oCurrentVeh.Description.url
		'   uHeader2.jpgfilename = m_oCurrentVeh.Description.VehicleImageFileName
		'   uHeader2.Description = m_oCurrentVeh.Description.VehicleDescription
		'End With
		
		' put the file header
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(iFree, SIG_129) 'file version signature to help identify between all the legacy .veh formats
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(iFree, uHeader, 2)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(iFree, uHeader2, OFFSET1)
		
		' now put our actual data
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(iFree, b, OFFSET2)
		'UPGRADE_NOTE: Object oZip may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oZip = Nothing
		
		FileClose(iFree)
		
		
		' Writes the vehicle collection to a file specified by the user
reghandler: 
		FileClose(iFree)
		Exit Sub
	End Sub
	
	Private Sub WriteChild(ByVal iNodeIndex As Short)
		Dim m_oCurrentVeh As Object
		Dim frmDesigner As Object
		' Write the child nodes to the table. This sub uses recursion
		' to loop through the child nodes. It receives the Index of
		' the node that has the children
		Dim tobj As clsFileLoader
		Dim i As Integer
		Dim iTempIndex As Short
		Dim Temp() As String
		Dim vc As Object
		Dim k As Integer
		
		' create an instance of the FileLoader class
		tobj = New clsFileLoader
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iTempIndex = frmDesigner.treeVehicle.Nodes(iNodeIndex).Child.FirstSibling.index
		'Loop through all a Parents Child Nodes
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		i = frmDesigner.treeVehicle.Nodes(iNodeIndex).children
		k = UBound(Components)
		
		'UPGRADE_WARNING: Lower bound of array Components was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim Preserve Components(k + i)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For i = 1 To frmDesigner.treeVehicle.Nodes(iNodeIndex).children
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object tobj.GetProperties(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vc = tobj.GetProperties(m_oCurrentVeh.Components(frmDesigner.treeVehicle.Nodes(iTempIndex).Key))
			k = k + 1
			Components(k).Properties = CreatePrintString(vc)
			
			'If the Node we are on has a child call the Sub again
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If frmDesigner.treeVehicle.Nodes(iTempIndex).children > 0 Then
				WriteChild((iTempIndex))
			End If
			
			'If we are not on the last child move to the next child Node
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If i <> frmDesigner.treeVehicle.Nodes(iNodeIndex).children Then
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iTempIndex = frmDesigner.treeVehicle.Nodes(iTempIndex).Next.index
			End If
		Next 
		
		'destroy the instance of the fileloader class
		'UPGRADE_NOTE: Object tobj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		tobj = Nothing
		
	End Sub
	
	Public Function PasteNode(ByRef sSourceKey() As String, ByRef sDestinationKey() As String, ByRef sKey() As String) As Object
		Dim tvwChild As Object
		Dim frmDesigner As Object
		Dim m_oCurrentVeh As Object
		'//to copy a node, you simply create a new object of the same type as the Source.
		Dim dType As Short
		Dim sText As String
		Dim iImage1 As Short
		Dim tobj As clsFileLoader
		Dim vc As Object
		Dim i As Integer
		Dim memberID As String
		Dim propvalue As Object
		Dim iCount As Integer
		Dim sOldKey As String
		Dim sNewKey As String
		Dim m As Integer
		
		On Error GoTo errorhandler
		
		Dim sLoc As String
		For iCount = 1 To UBound(sSourceKey)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			With m_oCurrentVeh.Components(sSourceKey(iCount))
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dType = .Datatype
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sText = .CustomDescription
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iImage1 = .Image
				
				
			End With
			
			'//first add it to the tree
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmDesigner.treeVehicle.Nodes.Add(sDestinationKey(iCount), tvwChild, sKey(iCount), sText, iImage1)
			'//now attempt to create the copied node and add it to the vehicle
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.addObject. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If m_oCurrentVeh.addObject(dType, sKey(iCount), sDestinationKey(iCount), iImage1, sText, True) = False Then
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				frmDesigner.treeVehicle.Nodes.Remove(sKey(iCount))
				MsgBox("Error copying node. Copy/Paste operation aborted.")
				Exit Function
			End If
			
			'//restore the key and parent key since it doesnt get properly applied
			'//in a paste operation
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_oCurrentVeh.Components(sKey(iCount)).Parent = sDestinationKey(iCount)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_oCurrentVeh.Components(sKey(iCount)).Key = sKey(iCount)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_oCurrentVeh.Components(sKey(iCount)).Datatype = dType
			
			'//we need to run the location check ourselves since
			'//we are by passing it in the AddObject using the LoadedFlag = TRUE
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If m_oCurrentVeh.Components(sKey(iCount)).LocationCheck = False Then
				' remove the object from the tree since the AddObject failed
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				frmDesigner.treeVehicle.Nodes.Remove(sKey(iCount))
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_oCurrentVeh.Components.Remove(sKey(iCount))
				MsgBox("Invalid paste location for this component. Paste aborted.")
				Exit Function
			End If
			
			'get the location so we can restore it
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sLoc = m_oCurrentVeh.Components(sKey(iCount)).location
			
			'//now we need to restore all the values EXCEPT for the Parent and Key values
			'//crap and also power system values
			tobj = New clsFileLoader
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object tobj.GetProperties(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vc = tobj.GetProperties(m_oCurrentVeh.Components(sSourceKey(iCount)))
			
			For i = 1 To UBound(vc)
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				memberID = vc(i, 0)
				'UPGRADE_WARNING: Couldn't resolve default property of object vc(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object propvalue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				propvalue = vc(0, i)
				If (memberID = "Location") Or (memberID = "Key") Or (memberID = "Parent") Then
					'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				ElseIf (VarType(propvalue) = VariantType.Array + VariantType.Object) Or (VarType(propvalue) = VariantType.Array + VariantType.String) Then 
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					tobj.LetProperties(m_oCurrentVeh.Components(sKey(iCount)), memberID, propvalue)
				End If
			Next 
			
			'finally, we can add our keychain keys
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.keymanager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_oCurrentVeh.keymanager.AddKeyChainKeys(sKey(iCount))
		Next 
		
		Exit Function
errorhandler: 
		If Err.Number = 35602 Then
			'//we need to change all the keys
			sNewKey = GetNextKey
			sOldKey = sKey(iCount)
			For m = iCount To UBound(sKey)
				If sKey(m) = sOldKey Then
					sKey(m) = sNewKey
				End If
			Next 
			Resume 
		End If
		
	End Function
	
	
	Function EncryptINI(ByVal Strg As String, ByVal Password As String) As String
		Dim b, s As String
		Dim i, j As Short
		Dim A2, A1, A3 As Short
		Dim P As String
		j = 1
		For i = 1 To Len(Password)
			P = P & Asc(Mid(Password, i, 1))
		Next 
		
		For i = 1 To Len(Strg)
			A1 = Asc(Mid(P, j, 1))
			j = j + 1 : If j > Len(P) Then j = 1
			A2 = Asc(Mid(Strg, i, 1))
			A3 = A1 Xor A2
			b = Hex(A3)
			If Len(b) < 2 Then b = "0" & b
			s = s & b
		Next 
		EncryptINI = s
	End Function
	
	Function DecryptINI(ByVal Strg As String, ByVal Password As String) As String
		Dim b, s As String
		Dim i, j As Short
		Dim A2, A1, A3 As Short
		Dim P As String
		j = 1
		For i = 1 To Len(Password)
			P = P & Asc(Mid(Password, i, 1))
		Next 
		
		For i = 1 To Len(Strg) Step 2
			A1 = Asc(Mid(P, j, 1))
			j = j + 1 : If j > Len(P) Then j = 1
			b = Mid(Strg, i, 2)
			A3 = Val("&H" & b)
			A2 = A1 Xor A3
			s = s & Chr(A2)
		Next 
		DecryptINI = s
	End Function
	
	
	Sub ReadINI()
		'UPGRADE_ISSUE: FileSystemObject object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oFile As FileSystemObject
		Dim oINI As cINI
		' fill our settings with default values first in case some of the INI values are missing
		With Settings
			.InitialDir = GVDPath
			.DesktopX = VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width)
			.DesktopY = VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height)
			'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
			.windowstate = vbNormal
			.FormTop = 0
			.FormLeft = 0
			.FormHeight = 600
			.FormWidth = 800
			.Splitter1 = 200
			.Splitter2 = 340
			.HSplitter = 340
			.bUseSurfaceAreaTable = True
			.bUseDefaultTextViewer = True
			.bUseDefaultWebBrowser = True
			.AuthorName = ""
			.Copyright = ""
			.email = ""
			.url = ""
			.Header = ""
			.Footer = ""
			.PublishEmailAddress = "veh@makosoft.com" 'todo: 02/16/02 Is this ok to have hardcoded?
			.DecimalPlaces = 2
			.FormatString = "standard"
			'.bQuickStart = False '02/16/02 MPJ (OBSOLETE)
			'.bSoundOff = False   '02/16/02 MPJ (obsolete)
			.bAssociateExt = False
			.TextExportPath = GVDPath
			.HTMLExportPath = GVDPath
			.VehiclesOpenPath = GVDPath
			.VehiclesSavePath = GVDPath
		End With
		
		oFile = New FileSystemObject
		oINI = New cINI
		
		' Make sure the INI file exists
		'UPGRADE_WARNING: Couldn't resolve default property of object oFile.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If oFile.FileExists(GVDPath & "\" & GVDINIFile) Then
			
			oINI.FileName = GVDPath & "\" & GVDINIFile
			
			'read in the settings and store them into our Settings UDT
			With Settings
				
				.TextExportPath = oINI.ReadString("Paths", "TextExportPath")
				.HTMLExportPath = oINI.ReadString("Paths", "HTMLExportPath")
				.VehiclesOpenPath = oINI.ReadString("Paths", "VehiclesOpenPath")
				.VehiclesSavePath = oINI.ReadString("Paths", "VehiclesSavePath")
				
				.Recent1 = oINI.ReadString("Recent", "Recent1")
				.Recent2 = oINI.ReadString("Recent", "Recent2")
				.Recent3 = oINI.ReadString("Recent", "Recent3")
				.Recent4 = oINI.ReadString("Recent", "Recent4")
				.Recent5 = oINI.ReadString("Recent", "Recent5")
				
				.HTMLBrowserPath = oINI.ReadString("Viewers", "HTMLBrowserPath")
				.TextViewerPath = oINI.ReadString("Viewers", "TextViewerPath")
				
				.DesktopX = oINI.ReadInteger("Display", "DesktopX")
				.DesktopY = oINI.ReadInteger("Display", "DesktopY")
				.windowstate = oINI.ReadInteger("Display", "State")
				.FormTop = oINI.ReadInteger("Display", "Top")
				.FormLeft = oINI.ReadInteger("Display", "Left")
				.FormWidth = oINI.ReadInteger("Display", "Width")
				.FormHeight = oINI.ReadInteger("Display", "Height")
				.Splitter1 = oINI.ReadInteger("Display", "Splitter1")
				If .Splitter1 <= 0 Then .Splitter1 = 4220
				
				.Splitter2 = oINI.ReadInteger("Display", "Splitter2")
				If .Splitter2 <= 0 Then .Splitter2 = 6720
				
				.HSplitter = oINI.ReadInteger("Display", "HSplitter")
				If .HSplitter <= 0 Then .HSplitter = 7000
				
				.bUseSurfaceAreaTable = oINI.ReadInteger("Config", "UseSurfaceAreaTable")
				.bUseDefaultTextViewer = oINI.ReadInteger("Config", "UseDefaultTextViewer")
				.bUseDefaultWebBrowser = oINI.ReadInteger("Config", "UseDefaultWebBrowser")
				.PublishEmailAddress = oINI.ReadString("Config", "PublishEmail")
				.DecimalPlaces = oINI.ReadInteger("Config", "DecimalPlaces")
				.FormatString = oINI.ReadString("Config", "FormatCode")
				'.bQuickStart = oINI.ReadInteger("Config", "QuickStart") '02/16/02 MPJ (obsolete)
				'.bSoundOff = oINI.ReadInteger("Config", "DisableSound") '02/16/02 MPJ (obsolete)
				.bAssociateExt = oINI.ReadInteger("Config", "AssociateExt")
				
				.AuthorName = oINI.ReadString("Author", "Name")
				.email = oINI.ReadString("Author", "Email")
				.url = oINI.ReadString("Author", "URL")
				.Copyright = oINI.ReadString("Author", "Copyright")
				.Header = oINI.ReadString("Author", "Header")
				.Footer = oINI.ReadString("Author", "Footer")
			End With
		End If
		
		
	End Sub
	
	Public Sub ReadLicenseFile()
		'UPGRADE_ISSUE: FileSystemObject object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oFile As FileSystemObject
		Dim iFree As Integer
		
		On Error GoTo errorhandler
		
		oFile = New FileSystemObject
		
		iFree = FreeFile
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oFile.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If oFile.FileExists(GVDPath & "\" & GVDLicenseFile) Then ' debug  this line needs to change because its saving the INI in the wrong place
			'retreive the data from the file and store it into the Settings udt
			FileOpen(iFree, GVDPath & "\" & GVDLicenseFile, OpenMode.Binary)
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileGet(iFree, RegInfo, 1)
			FileClose(iFree)
			
			With RegInfo
				gsRegName = VB6.CopyArray(.RegName)
				gsRegID = .RegID
				gsRegNum = VB6.CopyArray(.RegNum)
			End With
		End If
		
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If IsNothing(gsRegName) Then
			'UPGRADE_WARNING: Lower bound of array gsRegName was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim gsRegName(1)
		End If
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If IsNothing(gsRegNum) Then
			'UPGRADE_WARNING: Lower bound of array gsRegNum was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim gsRegNum(1)
		End If
		Exit Sub
		
errorhandler: 
		' must make sure our byte arrays are filled
		'UPGRADE_WARNING: Lower bound of array gsRegName was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim gsRegName(1)
		'UPGRADE_WARNING: Lower bound of array gsRegNum was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim gsRegNum(1)
		
		'first close the file
		FileClose(iFree)
		' now delete it
		'UPGRADE_WARNING: Couldn't resolve default property of object oFile.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If oFile.FileExists(GVDPath & "\" & GVDLicenseFile) Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object oFile.DeleteFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oFile.DeleteFile(GVDPath & "\" & GVDLicenseFile)
		End If
		
		System.Windows.Forms.Application.DoEvents()
		
	End Sub
	
	Public Sub WriteLicenseFile()
		
		Dim iFree As Integer
		
		iFree = FreeFile
		
		'first delete the file if it already exists
		FileOpen(iFree, GVDPath & "\" & GVDLicenseFile, OpenMode.Random)
		FileClose(iFree)
		
		iFree = FreeFile
		
		' open the License for binary write
		FileOpen(iFree, GVDPath & "\" & GVDLicenseFile, OpenMode.Binary)
		
		' update the relevant settings before we save it
		With RegInfo
			.RegID = gsRegID
			.RegName = VB6.CopyArray(gsRegName)
			.RegNum = VB6.CopyArray(gsRegNum)
		End With
		
		' save the Settings data and close the file
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(iFree, RegInfo)
		FileClose(iFree)
		
	End Sub
	
	Sub WriteINI()
		Dim frmDesigner As Object
		On Error Resume Next
		Dim oINI As cINI
		oINI = New cINI
		
		oINI.FileName = GVDPath & "\" & GVDINIFile
		
		With Settings
			Call oINI.WriteInteger("Display", "DesktopX", VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width))
			Call oINI.WriteInteger("Display", "DesktopY", VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height))
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.windowstate. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call oINI.WriteInteger("Display", "State", frmDesigner.windowstate)
			
			'JAW 2000.05.22
			'Splitter positions were not being saved when window was maximized.
			'If (frmDesigner.windowstate <> vbMaximized) And (frmDesigner.windowstate <> vbMinimized) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.windowstate. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (frmDesigner.windowstate <> System.Windows.Forms.FormWindowState.Minimized) Then
				Call oINI.WriteInteger("Display", "Top", Settings.FormTop)
				Call oINI.WriteInteger("Display", "Left", Settings.FormLeft)
				Call oINI.WriteInteger("Display", "Width", Settings.FormWidth)
				Call oINI.WriteInteger("Display", "Height", Settings.FormHeight)
				Call oINI.WriteInteger("Display", "Splitter1", Settings.Splitter1)
				' Call oINI.WriteInteger("Display", "Splitter2", frmDesigner.ListView1.Left + frmDesigner.ListView1.Width)
				Call oINI.WriteInteger("Display", "HSplitter", Settings.HSplitter)
			End If
			
			Call oINI.WriteInteger("Config", "UseSurfaceAreaTable", Settings.bUseSurfaceAreaTable)
			Call oINI.WriteInteger("Config", "UseDefaultTextViewer", Settings.bUseDefaultTextViewer)
			Call oINI.WriteInteger("Config", "UseDefaultWebBrowser", Settings.bUseDefaultWebBrowser)
			Call oINI.WriteString("Config", "PublishEmail", .PublishEmailAddress)
			Call oINI.WriteInteger("Config", "DecimalPlaces", .DecimalPlaces)
			Call oINI.WriteString("Config", "FormatCode", "standard")
			' Call oINI.WriteInteger("Config", "QuickStart", Settings.bQuickStart) '02/16/02 MPJ (obsolete)
			' Call oINI.WriteInteger("Config", "DisableSound", Settings.bSoundOff) '02/16/02 MPJ (obsolete)
			Call oINI.WriteInteger("Config", "AssociateExt", Settings.bAssociateExt)
			
			Call oINI.WriteString("Author", "Name", Settings.AuthorName)
			Call oINI.WriteString("Author", "Email", Settings.email)
			Call oINI.WriteString("Author", "URL", Settings.url)
			Call oINI.WriteString("Author", "Copyright", Settings.Copyright)
			Call oINI.WriteString("Author", "Header", Settings.Header)
			Call oINI.WriteString("Author", "Footer", Settings.Footer)
			
			Call oINI.WriteString("Paths", "App", GVDPath)
			Call oINI.WriteString("Paths", "TextExportPath", .TextExportPath)
			Call oINI.WriteString("Paths", "HTMLExportPath", .HTMLExportPath)
			Call oINI.WriteString("Paths", "VehiclesOpenPath", .VehiclesOpenPath)
			Call oINI.WriteString("Paths", "VehiclesSavePath", .VehiclesSavePath)
			
			Call oINI.WriteString("Viewers", "HTMLBrowserPath", .HTMLBrowserPath)
			Call oINI.WriteString("Viewers", "TextViewerPath", .TextViewerPath)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.mnuRecent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call oINI.WriteString("Recent", "Recent1", frmDesigner.mnuRecent(1).Caption)
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.mnuRecent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call oINI.WriteString("Recent", "Recent2", frmDesigner.mnuRecent(2).Caption)
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.mnuRecent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call oINI.WriteString("Recent", "Recent3", frmDesigner.mnuRecent(3).Caption)
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.mnuRecent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call oINI.WriteString("Recent", "Recent4", frmDesigner.mnuRecent(4).Caption)
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.mnuRecent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call oINI.WriteString("Recent", "Recent5", frmDesigner.mnuRecent(5).Caption)
		End With
	End Sub
	
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''OBSOLETE MPJ Oct.6.2002
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''Sub SaveComponent(ByVal sKey As String, ByVal FileName As String)
	''''''''takes the Key and FileName of a weapon component and saves it
	'''''''' Writes the vehicle collection to a file specified by the user
	'''''''
	'''''''Dim tobj As clsFileLoader
	'''''''Dim vc As Variant
	'''''''Dim Temp() As String
	'''''''Dim i As Long
	'''''''Dim iFree As Long
	'''''''
	'''''''On Error GoTo errorhandler
	'''''''
	'''''''' create an instance of the FileLoader class
	'''''''Set tobj = New clsFileLoader
	'''''''z = m_oCurrentVeh.GetOverDrive
	'''''''
	''''''''now get ready to print the properties for this object
	''''''''first delete the contents  of the file
	'''''''iFree = FreeFile
	'''''''Open FileName For Random As #iFree
	'''''''Close #iFree
	'''''''
	''''''''open the file for writing
	'''''''Open FileName For Output As #iFree
	'''''''
	''''''''print the first line
	'''''''With m_oCurrentVeh.Components(sKey)
	'''''''    Print #iFree, .Datatype, .Image, .SelectedImage, .CustomDescription
	'''''''End With
	''''''''output it all
	'''''''vc = tobj.GetProperties(m_oCurrentVeh.Components(sKey))
	'''''''Temp = CreatePrintString(vc)
	'''''''
	'''''''
	'''''''For i = 1 To UBound(Temp)
	'''''''
	'''''''    Print #iFree, EncryptINI$(Temp(i), z)
	'''''''Next
	'''''''
	''''''''close the file
	'''''''Close #iFree
	'''''''
	''''''''destroy the instance of the fileloader class
	'''''''Set tobj = Nothing
	'''''''Exit Sub
	'''''''
	'''''''errorhandler:
	'''''''
	'''''''    Close #iFree
	'''''''
	'''''''End Sub
	'''''''
	'''''''Function RestoreSavedItem(ByVal sFileName As String, ByVal sKey As String, ByVal sParent As String, ByRef sLocation As String)
	''''''''loads a saved component with a given filename and
	''''''''restores its properties
	'''''''    Dim tobj As clsFileLoader
	'''''''    Dim vc As Variant
	'''''''    Dim memberID As String
	'''''''    Dim propvalue As Variant
	'''''''    Dim arrkey As Variant
	'''''''    Dim i As Long
	'''''''    Dim j As Long
	'''''''    Dim iCount As Long
	'''''''    Dim iPropCount As Long
	'''''''    Dim lngUpper As Long
	'''''''    Dim iFree As Long
	'''''''    Dim sTemp As String
	'''''''
	'''''''    On Error GoTo errorhandler
	'''''''    z = m_oCurrentVeh.GetOverDrive
	'''''''    ReDim uRet(1)
	'''''''    iFree = FreeFile
	'''''''
	'''''''    Open sFileName For Input As iFree
	'''''''
	'''''''     '//load the file data
	'''''''    Line Input #iFree, sTemp '//skip the first line
	'''''''    Do While Not EOF(iFree)
	'''''''        Line Input #iFree, sTemp
	'''''''        sTemp = DecryptINI$(sTemp, z)
	'''''''        'fill the variant array that we will be passing into the FileLoader class
	'''''''        If Left(sTemp, 1) = "[" Then
	'''''''            'now remove the leading and trailing characters
	'''''''            i = i + 1
	'''''''            j = 0
	'''''''            ReDim Preserve Components(i)
	'''''''            sTemp = Mid(sTemp, 2, Len(sTemp) - 2)
	'''''''            Components(i).TreeInfo = sTemp
	'''''''
	'''''''        Else
	'''''''            j = j + 1
	'''''''            ReDim Preserve Components(i).Properties(j)
	'''''''            Components(i).Properties(j) = sTemp
	'''''''        End If
	'''''''    Loop
	'''''''    Close #iFree
	'''''''
	'''''''   '//restore the values
	'''''''    ' create an instance of the FileLoader class
	'''''''    Set tobj = New clsFileLoader
	'''''''
	'''''''    For iPropCount = 1 To UBound(Components(1).Properties)
	'''''''        vc = Split(Components(1).Properties(iPropCount), "|")
	'''''''        'do NOT restore keychains
	'''''''        If vc(1) = "[" Then
	'''''''
	'''''''        Else
	'''''''            'fill the properties for this object
	'''''''            Debug.Assert vc(0) <> "CombinedComponentVolume"
	'''''''            memberID = vc(0)
	'''''''            propvalue = vc(1)
	'''''''            '//we dont want to restore non relevant values
	'''''''            Select Case memberID
	'''''''                Case "Parent", "Key", "Location", "ParentDatatype", "PrintOutput"  '<-- Property Exclusion
	'''''''                Case Else
	'''''''                    tobj.LetProperties m_oCurrentVeh.Components(sKey), memberID, propvalue
	'''''''            End Select
	'''''''        End If
	'''''''    Next
	'''''''
	'''''''    ' set the actual key and parent key and location ' Though we exclude these out from above its probably not even necessary
	'''''''     'set the new key!
	'''''''    With m_oCurrentVeh.Components(sKey)
	'''''''        .Key = sKey
	'''''''        .Parent = sParent
	'''''''        .location = sLocation '//restore the user specified location since it got overwritten after restoring attributes+stats from disk
	'''''''    End With
	'''''''
	'''''''    ' 05/18/02 MPJ - Since AddObject now handles this, no need to AddKeyChainKeys are it results in TWO sets
	'''''''    ' of keys being added.
	'''''''    'finally, we can add our relevant keychain keys
	'''''''    'm_oCurrentVeh.keymanager.AddKeyChainKeys sKey
	'''''''
	'''''''Exit Function
	'''''''errorhandler:
	'''''''
	'''''''
	'''''''End Function
End Module