Option Strict Off
Option Explicit On
Friend Class clsComponentInfo
	
	
	Private m_lngCount As Integer
	
	Public Function LoadComponent(ByRef sPath As String) As Integer
		Dim sFileName As Object
		Dim hImgSmall As Object
		Dim SHGFI_SMALLICON As Object
		Dim BASIC_SHGFI_FLAGS As Object
		Dim m_shinfo As Object
		Dim SHGetFileInfo As Object
		Dim sIconPath As Object
		Dim lngRet As Object
		Dim pavAuto As Object
		Dim m_oXML As Object
		
		
		m_oXML = New cXML
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oXML.Initialize. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object lngRet. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngRet = m_oXML.Initialize(pavAuto)
		'UPGRADE_WARNING: Couldn't resolve default property of object lngRet. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If lngRet = 0 Then
			modHelper.InfoPrint(1, "Error:  Cannot load components.  Reason:  Could not initialize XML Parser.  Solution?:  Install Microsoft XML parser.")
			Exit Function
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sFileName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oXML.OpenFromFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Dim o As Object
		If m_oXML.OpenFromFile(sPath & sFileName, True) Then
			'TODO: Best way to do this, is to also load the DEF file info first, then overwrite it with .CMP if
			' applicable.  That way we are guaranteed to load all data since the DEF file is required to have
			' EVERYTHING except things such as attributes output and print output stuff?
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oXML.FindNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			o = m_oXML.FindNode("/GVD_SAVED_COMP/Component/Definition/Name")
			'sComponentName = oXML.ReadNode("/GVD_SAVED_COMP/Component/Definition/Name")
			'sComponentName = oxml.NodeCount
			
			
			'resortXML.XMLDocument.documentElement.childNodes.item(1).
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oXML.ReadNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sIconPath. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sIconPath = My.Application.Info.DirectoryPath & m_oXML.ReadNode("/GVD_SAVED_COMP/Component/Definition/IconFile")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object SHGetFileInfo(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object hImgSmall. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hImgSmall = SHGetFileInfo(sIconPath, 0, m_shinfo, Len(m_shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
			
			
		End If
	End Function
	
	
	Public ReadOnly Property ItemCount() As Integer
		Get
			ItemCount = m_lngCount
		End Get
	End Property
End Class