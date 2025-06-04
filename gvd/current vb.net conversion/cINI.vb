Option Strict Off
Option Explicit On
Friend Class cINI
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Private Declare Function GetPrivateProfileInt Lib "kernel32"  Alias "GetPrivateProfileIntA"(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
	Private Declare Function GetPrivateProfileSection Lib "kernel32"  Alias "GetPrivateProfileSectionA"(ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function WritePrivateProfileString Lib "kernel32"  Alias "WritePrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
	Private Declare Function WritePrivateProfileSection Lib "kernel32"  Alias "WritePrivateProfileSectionA"(ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
	
	Private Const MAX_SECTION_SIZE As Short = 2000
	Dim m_sFileName As String
	
	'----------------------------------------------------------------------------
	'NOTE:  This class requires the project its in to reference the "Microsoft Scripting Runtime" (scrrun.dll)
	'------------------------------------------------------------------------------
	
	
	' TODO:  This class needs a flag for class initialization e.g. m_bLibraryInitialied
	'        which should be checked when any external call is made.
	
	Public Property FileName() As String
		Get
			FileName = m_sFileName
		End Get
		Set(ByVal Value As String)
			'//here the user sets the name of the INI file
			'//we are working with
			m_sFileName = Value
			
			'//first check to make sure this file exists.
			'//if it doesnt, create it
			'UPGRADE_ISSUE: FileSystemObject object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
			Dim oFile As FileSystemObject
			oFile = New FileSystemObject
			
			'UPGRADE_WARNING: Couldn't resolve default property of object oFile.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not oFile.FileExists(m_sFileName) Then
				'//todo: first needs to build the path or error ensues
				
				'UPGRADE_WARNING: Couldn't resolve default property of object oFile.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oFile.CreateTextFile(m_sFileName)
			End If
			'UPGRADE_NOTE: Object oFile may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oFile = Nothing
		End Set
	End Property
	
	Public Sub DeleteSection(ByVal sSectionName As String)
		'//this accepts a section name and deletes it and all keys in it
		Dim iFree As Integer
		Dim sArray() As String
		Dim s As String
		Dim i As Integer
		Dim j As Integer
		
		iFree = FreeFile
		ReDim sArray(0)
		i = 0
		
		FileOpen(iFree, m_sFileName, OpenMode.Input)
		
		Do While Not EOF(iFree)
			s = LineInput(iFree)
			If s <> "" Then
				ReDim Preserve sArray(i)
				sArray(i) = s
				i = i + 1
			End If
		Loop 
		FileClose(iFree)
		
		'//now find the section in our array and delete it
		For i = 0 To UBound(sArray)
			If sArray(i) = "[" & sSectionName & "]" Then
				sArray(i) = ""
				For j = i + 1 To UBound(sArray)
					If Left(sArray(j), 1) <> "[" Then
						sArray(j) = ""
					Else
						Exit For
					End If
				Next 
			End If
		Next 
		
		'//now write a new file
		iFree = FreeFile
		
		FileOpen(iFree, m_sFileName, OpenMode.Output)
		
		For i = 0 To UBound(sArray)
			If sArray(i) <> "" Then
				PrintLine(iFree, sArray(i))
			End If
		Next 
		FileClose(iFree)
	End Sub
	
	'UPGRADE_WARNING: ParamArray vList was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Public Sub WriteSection(ByVal sSectionName As String, ParamArray ByVal vList() As Object)
		'//this writes an entire section of data
		'//the user provides the section name and the list of keys
		'//NOTE: This does not write in any Key Values...
		
		Dim sArray() As String
		Dim s As String
		Dim i As Integer
		
		For i = 0 To UBound(vList)
			'UPGRADE_WARNING: Couldn't resolve default property of object vList(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sArray = Split(vList(i), "=")
			s = s & sArray(0) & Chr(10) & sArray(1) & Chr(10) & Chr(10)
		Next 
		
		Call WritePrivateProfileSection(sSectionName, s, m_sFileName)
		
	End Sub
	
	Public Function GetNextIndexedSection(ByVal sSectionName As String) As String
		'//this function creates indexed section names.  It finds the first available section
		'//and creates it and returns the name of the section it created
		Dim j, i, k As Object
		Dim l As Integer
		Dim iFree As Integer
		Dim sName As String
		Dim s As String
		Dim sArray() As String
		
		iFree = FreeFile
		ReDim sArray(0)
		'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		i = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		j = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		k = 0
		l = 0
		
		FileOpen(iFree, m_sFileName, OpenMode.Input)
		'//get a list of all sections within this file
		Do While Not EOF(iFree)
			s = LineInput(iFree)
			If Left(s, 1) = "[" Then
				ReDim Preserve sArray(i)
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sArray(i) = Mid(s, 2, Len(s) - 2)
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				i = i + 1
			End If
		Loop 
		FileClose(iFree)
		
		'//now find a free section
		'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sName = sSectionName & k
		For i = 0 To UBound(sArray)
			For j = 0 To UBound(sArray)
				'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If sName = sArray(j) Then
					l = l + 1
					Exit For
				End If
			Next 
			'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If l = k Then Exit For
			'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			k = k + 1
			'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sName = sSectionName & k
		Next 
		'//create the current section
		GetNextIndexedSection = sName
	End Function
	
	Public Sub WriteString(ByVal sSectionName As String, ByVal sKeyName As String, ByVal s As String)
		'//accepts a string and writes it to the ini file
		Call WritePrivateProfileString(sSectionName, sKeyName, s, m_sFileName)
	End Sub
	
	Public Sub WriteInteger(ByVal sSectionName As String, ByVal sKeyName As String, ByVal l As Integer)
		'//accepts a string and writes it to the ini file
		Dim s As String
		
		s = Str(l)
		Call WritePrivateProfileString(sSectionName, sKeyName, s, m_sFileName)
	End Sub
	
	Public Sub WriteBoolean(ByVal sSectionName As String, ByVal sKeyName As String, ByVal b As Boolean)
		'//this sub accepts a boolean and writes it as a string "TRUE" or "FALSE" to the
		'//INI file
		Dim s As String
		If b = True Then
			s = "TRUE"
		Else
			s = "FALSE"
		End If
		Call WritePrivateProfileString(sSectionName, sKeyName, s, m_sFileName)
	End Sub
	
	Public Function ReadString(ByVal sSectionName As String, ByRef sKeyName As String) As String
		'//accepts the Section and Keyname and returns the value of the key
		Dim sReturnString As String
		sReturnString = Space(255)
		ReadString = Left(sReturnString, GetPrivateProfileString(sSectionName, sKeyName, " ", sReturnString, 255, m_sFileName))
	End Function
	
	Public Function ReadInteger(ByVal sSectionName As String, ByRef sKeyName As String) As Integer
		'//accepts the Section and Keyname and returns the integer value of the key
		ReadInteger = GetPrivateProfileInt(sSectionName, sKeyName, 0, m_sFileName)
	End Function
	
	Public Function ReadSection(ByVal sSectionName As String) As String()
		'//this function accepts the name of a section and then retreives
		'//all of the keynames
		
		Dim i As Object
		Dim j As Integer
		Dim s As String
		Dim sRet() As String
		Dim lngNullOffset As Integer
		Dim lngKeyValueOffset As Integer
		
		On Error GoTo errorhandler
		
		s = New String(Chr(0), MAX_SECTION_SIZE) '//allocate space in our string
		j = 0
		
		'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		i = GetPrivateProfileSection(sSectionName, s, MAX_SECTION_SIZE, m_sFileName)
		
		'//load the sections data into the array
		Do 
			lngNullOffset = InStr(s, Chr(0))
			If lngNullOffset > 1 Then
				ReDim Preserve sRet(j)
				
				' only save the keyname, not the value
				lngKeyValueOffset = InStr(s, "=")
				sRet(j) = Mid(s, 1, lngKeyValueOffset - 1)
				'sRet(j) = Mid(s, 1, lngNullOffset - 1)
				s = Mid(s, lngNullOffset + 1)
				j = j + 1
			End If
		Loop While lngNullOffset > 1
		
		ReadSection = VB6.CopyArray(sRet)
		Exit Function
errorhandler: 
		Exit Function
	End Function
	
	Public Function RetreiveSectionNames() As String()
		'//this function opens the INI and returns the list of
		'//all the section names in an string array
		Dim iFileNum As Integer
		iFileNum = FreeFile
		Dim s As String
		Dim sRet() As String
		Dim i As Integer
		
		FileOpen(iFileNum, m_sFileName, OpenMode.Input)
		
		Do While Not EOF(iFileNum)
			s = LineInput(iFileNum)
			If Left(s, 1) = "[" Then
				ReDim Preserve sRet(i)
				s = Mid(s, 2, Len(s) - 2)
				sRet(i) = s
				i = i + 1
			End If
		Loop 
		FileClose(iFileNum)
		RetreiveSectionNames = VB6.CopyArray(sRet)
	End Function
End Class