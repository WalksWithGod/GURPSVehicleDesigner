Option Strict Off
Option Explicit On
Module modListViewHelper
	
	'constants and declarations for the listview column resizing
	Private Const LVM_First As Integer = &H1000s
	Private Const LVSCW_AUTOSIZE As Short = -1
	Private Const LVM_SETCOLUMNWIDTH As Integer = LVM_First + 30
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
	Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Integer) As Integer
	
	Private Const SHGFI_DISPLAYNAME As Short = &H200s
	Private Const SHGFI_EXETYPE As Short = &H2000s
	Private Const SHGFI_SYSICONINDEX As Short = &H4000s 'system icon index
	Private Const SHGFI_LARGEICON As Short = &H0s 'large icon
	Private Const SHGFI_SMALLICON As Short = &H1s 'small icon
	Private m_shinfo As SHFILEINFO
	
	Private Const SHGFI_SHELLICONSIZE As Short = &H4s
	Private Const SHGFI_TYPENAME As Short = &H400s
	Private Const BASIC_SHGFI_FLAGS As Boolean = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
	
	Private Const MAX_PATH As Short = 255
	Private Const ILD_TRANSPARENT As Short = &H1s 'display transparent
	
	'UPGRADE_WARNING: Structure SHFILEINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function SHGetFileInfo Lib "shell32.dll"  Alias "SHGetFileInfoA"(ByVal pszPath As String, ByVal dwFileAttributes As Integer, ByRef psfi As SHFILEINFO, ByVal cbSizeFileInfo As Integer, ByVal uFlags As Integer) As Integer
	
	Private Structure SHFILEINFO
		Dim hIcon As Integer
		Dim iIcon As Integer
		Dim dwAttributes As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(MAX_PATH),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=MAX_PATH)> Public szDisplayName() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(80),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=80)> Public szTypeName() As Char
	End Structure
	
	
	Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Integer, ByVal i As Integer, ByVal hDCDest As Integer, ByVal x As Integer, ByVal y As Integer, ByVal flags As Integer) As Integer
	
	Public Function ReadComponentXML(ByRef uCmp As udtComponent, ByRef oXNode As MSXML2.IXMLDOMNode, ByRef oXML As cXML) As Boolean
		Dim MSXML2 As Object
		
		'UPGRADE_ISSUE: cXML object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oXDef As cXML
		Dim sXPath As String
		Dim hImageSmall As Integer
		
		On Error GoTo err_Renamed
		'--
		'UPGRADE_WARNING: Couldn't resolve default property of object oXML.GetXPath. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sXPath = oXML.GetXPath(oXNode)
		'uCmp.DefPath = App.Path & oXNode.Attributes.getNamedItem(XML_NODE_DEFPATH).nodeValue
		'UPGRADE_WARNING: Couldn't resolve default property of object oXNode.selectSingleNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		uCmp.DefPath = My.Application.Info.DirectoryPath & oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_DEFPATH & "']").nodeTypedValue
		'UPGRADE_WARNING: Couldn't resolve default property of object oXNode.selectSingleNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		uCmp.GUID = oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_GUID & "']").nodeTypedValue
		'UPGRADE_WARNING: Couldn't resolve default property of object oXNode.selectSingleNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		uCmp.Text = oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_NAME & "']").nodeTypedValue
		
		
		'Now we need to open the XML Defintion file
		oXDef = New cXML
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oXDef.OpenFromFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If oXDef.OpenFromFile(uCmp.DefPath, True) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oXDef.GetRootNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oXNode = oXDef.GetRootNode
			'UPGRADE_WARNING: Couldn't resolve default property of object oXNode.selectSingleNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oXNode = oXNode.selectSingleNode(XML_NODE_OBJECT)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object oXNode.selectSingleNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If uCmp.GUID = oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_GUID & "']").nodeTypedValue Then
				'UPGRADE_WARNING: Couldn't resolve default property of object oXNode.selectSingleNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				uCmp.IconPath = My.Application.Info.DirectoryPath & oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_IMAGE & "']").nodeTypedValue
				'UPGRADE_WARNING: Couldn't resolve default property of object oXNode.selectSingleNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				uCmp.Classname = oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_CLASSNAME & "']").nodeTypedValue
				
				
				'note: We need to release the file handle it references or else
				'       if the file happens to be listed more than ONCE in the same
				'       vehicle, the next instance of XML reader wont be able to open it
				'UPGRADE_NOTE: Object oXDef may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				oXDef = Nothing
				ReadComponentXML = True
				
				AddIconToImageList(uCmp.IconPath)
			Else
				Debug.Print("modListViewHelper:ReadComponentXML() -- Error:  '" & uCmp.Text & ".cmp' GUID does not match it's .def GUID.")
			End If
		End If
		'-
		Exit Function
err_Renamed: 
		Debug.Print("modListViewHelper:ReadComponentXML -- Error #" & Err.Number & " " & Err.Description)
		Resume Next
	End Function
	
	Sub LoadListView(ByVal sComponentType As String)
		Dim frmDesigner As Object
		Dim MSXML2 As Object
		Dim i As Integer
		Dim sFileName As String
		Dim sPath As String
		Dim sKey As String
		Dim sIconKey As String
		Dim uComponent As udtComponent
		'--------
		'UPGRADE_ISSUE: cXML object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oXML As cXML
		Dim oXRoot As MSXML2.IXMLDOMNode
		'UPGRADE_ISSUE: MSXML2.IXMLDOMNode object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oXNode As MSXML2.IXMLDOMNode
		
		' prevent listview updates til we're done updating
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ListView1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LockWindowUpdate(frmDesigner.ListView1.hwnd)
		' Clear the listview
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ListView1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.ListView1.ListItems.Clear()
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ListView1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ImageList1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.ListView1.SmallIcons = frmDesigner.ImageList1 '<-- add new icons to this list
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ListView1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.ListView1.Sorted = True
		sPath = My.Application.Info.DirectoryPath & "\components\" & sComponentType & "\" 'todo: need better use of constants for path
		' start searching for components and add them.
		' NOTE: It only displays the top level component even when a file contains nested components
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		sFileName = Dir(sPath & "*.cmp")
		
		
		oXML = New cXML
		
		Do While sFileName <> ""
			'UPGRADE_WARNING: Couldn't resolve default property of object oXML.OpenFromFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If oXML.OpenFromFile(sPath & sFileName, True) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object oXML.GetRootNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oXRoot = oXML.GetRootNode()
				'UPGRADE_WARNING: Couldn't resolve default property of object oXRoot.selectSingleNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oXNode = oXRoot.selectSingleNode(XML_NODE_OBJECT & "[@" & XML_ATTRIB_HANDLE & "='0_']")
				'--
				If Not oXNode Is Nothing Then
					Debug.Print("modListViewHelper.LoadListView() -- Found root object in CMP.XML file '" & sPath & sFileName & "'.  Attempting to load DEF.XML")
					If ReadComponentXML(uComponent, oXNode, oXML) Then
						' all checks out, we can finally add it to the tree
						sIconKey = uComponent.IconPath
						
						sKey = sPath & sFileName
						uComponent.ComponentPath = sKey
						AddListViewItem(sKey, uComponent.Text, sIconKey)
					End If
				Else
					Debug.Print("modListViewHelper.LoadListView() -- Could not find root object in XML file '" & sPath & sFileName & "'.  Most like caused by object handle not set to '0_'")
				End If
			End If
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			sFileName = Dir()
			uComponent.ComponentPath = sPath & sFileName
		Loop 
		
		'set the column widths
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ListView1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For i = 0 To frmDesigner.ListView1.ColumnHeaders.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ListView1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call SendMessage(frmDesigner.ListView1.hwnd, LVM_SETCOLUMNWIDTH, i, 0)
		Next 
		
		'UPGRADE_NOTE: Object oXRoot may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oXRoot = Nothing
		'UPGRADE_NOTE: Object oXNode may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oXNode = Nothing
		LockWindowUpdate(0)
		Exit Sub
		
errorhandler: 
		LockWindowUpdate(0)
		'UPGRADE_WARNING: Couldn't resolve default property of object Error.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Error.Number. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Debug.Print("modListViewHelper:LoadListView -- Error " & ErrorToString().Number & " " & ErrorToString().Description)
		
	End Sub
	
	Public Function KeyFromLong(ByVal l As Integer) As String
		KeyFromLong = CStr(l) & "_"
	End Function
	
	Private Function ImageListKeyExists(ByRef sKey As String) As Boolean
		Dim frmDesigner As Object
		On Error GoTo errorhandler
		Dim s As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ImageList1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		s = frmDesigner.ImageList1.ListImages.Item(sKey).Tag
		ImageListKeyExists = True
		Exit Function
errorhandler: 
		If Err.Number = 35601 Then
			ImageListKeyExists = False
		Else
			' we shouldnt be encountering any other error types here....
			modHelper.InfoPrint(1, "modListViewHelper:ImageListKeyExists -- Unexpected Error# " & Err.Number & " " & Err.Description)
			ImageListKeyExists = False
		End If
	End Function
	Private Function StripExtensionFromFileName(ByVal sName As String) As String
		Dim lngLength As Integer
		Dim lngExtensionPosition As Integer
		
		Dim i As Integer
		
		lngLength = Len(sName)
		
		For i = lngLength To 1 Step -1
			If Mid(sName, i, 1) = "." Then Exit For
		Next 
		
		If (i = 1) And Mid(sName, 1, 1) <> "." Then
			' there is no extension to strip off
			StripExtensionFromFileName = sName
		Else
			lngExtensionPosition = i - 1
			StripExtensionFromFileName = Left(sName, lngExtensionPosition)
		End If
	End Function
	
	Public Function CreateImageFromIconFile(ByRef sIconPath As String) As Integer
		Dim frmDesigner As Object
		Dim lngRet As Integer
		Dim hIcon As Integer
		
		hIcon = SHGetFileInfo(sIconPath, 0, m_shinfo, Len(m_shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
		
		'IMPORTANT: picIcon's width and height in pixels needs to be 16x16 else when Image_List draw is used
		'it will do some funky scaling and will screw the image up
		'frmDesigner.picIcon.ScaleMode = vbPixels
		'frmDesigner.picIcon.Height = 16
		'frmDesigner.picIcon.Width = 16
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.picIcon. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.picIcon.Picture = Nothing
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.picIcon. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngRet = ImageList_Draw(hIcon, m_shinfo.iIcon, frmDesigner.picIcon.hdc, 0, 0, ILD_TRANSPARENT)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.picIcon. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.picIcon.Picture = frmDesigner.picIcon.Image
		
		CreateImageFromIconFile = lngRet
		
	End Function
	
	Public Function AddIconToImageList(ByRef sIconKey As String) As Integer
		Dim frmDesigner As Object
		'UPGRADE_ISSUE: ListImage object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim imgX As ListImage
		Dim hImageSmall As Integer
		Dim lRet As Integer
		
		If Not ImageListKeyExists(sIconKey) Then
			hImageSmall = CreateImageFromIconFile(sIconKey)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.picIcon. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ImageList1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			imgX = frmDesigner.ImageList1.ListImages.Add( , sIconKey, frmDesigner.picIcon.Picture)
			'also add it to our treex
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.picIcon. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lRet = frmDesigner.treeVehicle.AddImage(sIconKey, frmDesigner.picIcon.Picture)
			'todo: do we need to ensure this got added to the tree's built in image list?
			' do we need to check that its not already added to the tree's built in control?
			' wouldn't it error if it were already in?  Only if "somehow" it was in the tree but magically not in the imagelist1
			Debug.Print("modListViewHelper:AddIconToImageList -- " & lRet)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ImageList1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			imgX = frmDesigner.ImageList1.ListImages.Item(sIconKey)
		End If
		
		
		'UPGRADE_NOTE: Object imgX may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		imgX = Nothing
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ImageList1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddIconToImageList = frmDesigner.ImageList1.ListImages(sIconKey).index
		Exit Function
errorhandler: 
		'UPGRADE_NOTE: Object imgX may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		imgX = Nothing
		AddIconToImageList = False
	End Function
	
	Private Function AddListViewItem(ByRef sKey As String, ByRef sNodeName As String, ByRef sIconKey As String) As Integer
		Dim frmDesigner As Object
		'UPGRADE_ISSUE: ListItem object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim itmX As ListItem
		
		If ImageListKeyExists(sIconKey) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ListView1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			itmX = frmDesigner.ListView1.ListItems.Add( , sKey, sNodeName)
			'since the treeview and listview share the same imagelist, and uses the same filename=KEY convention
			'we dont have to worry about coordinating listview and treeview icons.
			'UPGRADE_WARNING: Couldn't resolve default property of object itmX.SmallIcon. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ImageList1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			itmX.SmallIcon = frmDesigner.ImageList1.ListImages(sIconKey).Key
			'UPGRADE_WARNING: Couldn't resolve default property of object itmX.Tag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			itmX.Tag = sIconKey
			'UPGRADE_NOTE: Object itmX may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			itmX = Nothing
			AddListViewItem = True
		Else
			AddListViewItem = False
		End If
		
	End Function
End Module