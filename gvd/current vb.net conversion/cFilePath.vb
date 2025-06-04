Option Strict Off
Option Explicit On
Friend Class cFilePath
	
	'todo: this src ripped from planet source code.  Ideally, id like to make this a module
	' and not a class.  Sometimes its a pain to instance an object just to call a simple function.
	' on the other hand, it doesnt hang in memory until i need it... think about this some other time
	
	' Variable to hold the Path.
	'
	Private sPath As String
	
	'
	' API declares
	'
	
	' Maximum path length
	Private Const MAX_PATH As Integer = 260
	
	Private Declare Function lstrlen Lib "kernel32"  Alias "lstrlenA"(ByVal lpString As String) As Integer
	Private Declare Function lstrlenptr Lib "kernel32"  Alias "lstrlenA"(ByVal lpString As Integer) As Integer
	Private Declare Sub CopyMemoryLpToLp Lib "kernel32"  Alias "RtlMoveMemory"(ByVal pDest As Integer, ByVal pSrc As Integer, ByVal ByteLen As Integer)
	
	
	
	' Directory functions
	Private Declare Function PathIsDirectory Lib "Shlwapi"  Alias "PathIsDirectoryA"(ByVal sPath As String) As Boolean
	Private Declare Function PathRemoveFileSpec Lib "Shlwapi"  Alias "PathRemoveFileSpecA"(ByVal sPath As String) As Boolean
	Private Declare Function PathAddBackslash Lib "Shlwapi"  Alias "PathAddBackslashA"(ByVal sPath As String) As Integer
	Private Declare Function PathRemoveBackslash Lib "Shlwapi"  Alias "PathRemoveBackslashA"(ByVal sPath As String) As Integer
	Private Declare Function PathIsPrefix Lib "Shlwapi"  Alias "PathIsPrefixA"(ByVal sPrefix As String, ByVal sPath As String) As Boolean
	Private Declare Function PathIsRelative Lib "Shlwapi"  Alias "PathIsRelativeA"(ByVal sPath As String) As Boolean
	
	' File functions
	Private Declare Function PathFileExists Lib "Shlwapi"  Alias "PathFileExistsA"(ByVal sPath As String) As Boolean
	Private Declare Function PathIsFileSpec Lib "Shlwapi"  Alias "PathIsFileSpecA"(ByVal sPath As String) As Boolean
	Private Declare Sub PathStripPath Lib "Shlwapi"  Alias "PathStripPathA"(ByVal sPath As String)
	Private Declare Function PathFindFileName Lib "Shlwapi"  Alias "PathFindFileNameA"(ByVal pPath As String) As Integer
	
	' Extension functions
	Private Declare Function PathAddExtension Lib "Shlwapi"  Alias "PathAddExtensionA"(ByVal sPath As String, ByVal sExtension As String) As Boolean
	Private Declare Sub PathRemoveExtension Lib "Shlwapi"  Alias "PathRemoveExtensionA"(ByVal sPath As String)
	Private Declare Function PathFindExtension Lib "Shlwapi"  Alias "PathFindExtensionA"(ByVal pPath As String) As Integer
	Private Declare Function PathRenameExtension Lib "Shlwapi"  Alias "PathRenameExtensionA"(ByVal sPath As String, ByVal sExtension As String) As Boolean
	Private Declare Function PathMatchSpec Lib "Shlwapi"  Alias "PathMatchSpecA"(ByVal sFileParam As String, ByVal sSpec As String) As Boolean
	
	' UNC/URL functions
	Private Declare Function PathIsUNC Lib "Shlwapi"  Alias "PathIsUNCA"(ByVal sPath As String) As Boolean
	Private Declare Function PathIsUNCServer Lib "Shlwapi"  Alias "PathIsUNCServerA"(ByVal sPath As String) As Boolean
	Private Declare Function PathIsUNCServerShare Lib "Shlwapi"  Alias "PathIsUNCServerShareA"(ByVal sPath As String) As Boolean
	Private Declare Function PathIsURL Lib "Shlwapi"  Alias "PathIsURLA"(ByVal sPath As String) As Boolean
	
	'Root/Drive functions
	Private Declare Function PathIsRoot Lib "Shlwapi"  Alias "PathIsRootA"(ByVal sPath As String) As Boolean
	Private Declare Function PathStripToRoot Lib "Shlwapi"  Alias "PathStripToRootA"(ByVal szRoot As String) As Boolean
	Private Declare Function PathSkipRoot Lib "Shlwapi"  Alias "PathSkipRootA"(ByVal sPath As String) As Integer
	Private Declare Function PathGetDriveNumber Lib "Shlwapi"  Alias "PathGetDriveNumberA"(ByVal sPath As String) As Integer
	
	'Building functions
	Private Declare Function PathAppend Lib "Shlwapi"  Alias "PathAppendA"(ByVal sPath As String, ByVal sMore As String) As Boolean
	Private Declare Function PathCombine Lib "Shlwapi"  Alias "PathCombineA"(ByVal sDest As String, ByVal sDir As String, ByVal sFile As String) As Integer
	
	'Formatting functions
	Private Declare Sub PathQuoteSpaces Lib "Shlwapi"  Alias "PathQuoteSpacesA"(ByVal s As String)
	Private Declare Sub PathUnquoteSpaces Lib "Shlwapi"  Alias "PathUnquoteSpacesA"(ByVal s As String)
	Private Declare Function PathCompactPath Lib "Shlwapi"  Alias "PathCompactPathA"(ByVal hdc As Integer, ByVal sPath As String, ByVal dx As Short) As Boolean
	Private Declare Function PathCompactPathEx Lib "Shlwapi"  Alias "PathCompactPathExA"(ByVal sOut As String, ByVal sSrc As String, ByVal cchMax As Short, ByRef dwFlags As Integer) As Boolean
	
	' Adds a backslash to the end of a string to create the correct syntax
	' for a path.
	Public Function AddBackslash() As String
		Call PathAddBackslash(sPath)
		AddBackslash = Path
	End Function
	
	' Adds a file extension to a path string.
	Public Function AddExtension(ByVal NewExtension As String) As String
		Call PathAddExtension(sPath, NewExtension)
		AddExtension = Path
	End Function
	
	' Appends one path to the end of another.
	Public Function Append(ByVal Filename As String) As String
		Call PathAppend(sPath, Trim(Filename))
		Append = Path
	End Function
	
	'Concatenates two strings that represent properly formed paths into
	'one path, as well as any relative path pieces.
	Public Function Combine(ByVal Directory As String, ByVal Filename As String) As String
		Call PathCombine(sPath, Directory, Filename)
		Combine = Path
	End Function
	
	' Truncates a file path to fit within a given pixel width by replacing
	' path components with ellipses.
	Public Function Compact(ByVal hdc As Integer, ByVal Width As Integer) As String
		Dim sTemp As String
		sTemp = sPath
		Call PathCompactPath(hdc, sTemp, Width)
		Compact = TrimNull(sTemp)
	End Function
	
	' Truncates a path to fit within a certain number of characters by
	' replacing path components with ellipses.
	Public Function CompactEx(ByVal Chars As Short, Optional ByVal Separator As String = "\") As String
		Dim sDest As String
		sDest = New String(Chr(0), MAX_PATH)
		Call PathCompactPathEx(sDest, sPath, Chars, Asc(Left(Separator, 1)))
		CompactEx = TrimNull(sDest)
	End Function
	
	' This function tests the validity of the file and path. It works only
	' on the local file system or on a remote drive that has been mounted
	' to a drive letter.
	Public Function Exists() As Boolean
		Exists = CBool(PathFileExists(sPath))
	End Function
	
	' Returns the directory only.
	Public Function GetDirectory() As String
		Dim sTemp As String
		sTemp = sPath
		Call PathRemoveFileSpec(sTemp)
		GetDirectory = TrimNull(sTemp)
	End Function
	
	' Retrieve the drive letter from the current path.
	Public Function GetDrive() As String
		GetDrive = Chr(Asc("A") + PathGetDriveNumber(sPath))
	End Function
	
	' Retrieve the drive number from the current path.
	Public Function GetDriveNumber() As Short
		GetDriveNumber = PathGetDriveNumber(sPath)
	End Function
	
	' Retrieve the file extension from the current path.
	Public Function GetExtension() As String
		Dim lpszExt As Integer
		Dim bBuffer(MAX_PATH) As Byte
		Dim sTemp As String
		
		lpszExt = PathFindExtension(sPath)
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		Call CopyMemoryLpToLp(VarPtr(bBuffer(0)), lpszExt, lstrlenptr(lpszExt))
		
		'UPGRADE_ISSUE: Constant vbUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		sTemp = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(bBuffer), vbUnicode)
		GetExtension = TrimNull(sTemp)
	End Function
	
	' Retrieve the file name from the current path.
	Public Function GetFilename() As String
		Dim lpszName As Integer
		Dim bBuffer(MAX_PATH) As Byte
		Dim sTemp As String
		
		lpszName = PathFindFileName(sPath)
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		Call CopyMemoryLpToLp(VarPtr(bBuffer(0)), lpszName, lstrlenptr(lpszName))
		
		'UPGRADE_ISSUE: Constant vbUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		sTemp = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(bBuffer), vbUnicode)
		GetFilename = TrimNull(sTemp)
	End Function
	
	' Retrieve the root folder from the current path.
	Public Function GetRoot() As String
		Dim sTemp As String
		sTemp = sPath
		Call PathStripToRoot(sTemp)
		GetRoot = TrimNull(sTemp)
	End Function
	
	' Returns TRUE if the path is a valid directory, or FALSE otherwise.
	Public ReadOnly Property IsDirectory() As Boolean
		Get
			IsDirectory = CBool(PathIsDirectory(sPath))
		End Get
	End Property
	
	' Indicates wetwer the path is considered to be a File Spec path
	' (i.e. if there are no path delimiting characters present).
	Public ReadOnly Property IsFileSpec() As Boolean
		Get
			IsFileSpec = CBool(PathIsFileSpec(sPath))
		End Get
	End Property
	
	' Searches a path and determines if it is relative.
	Public ReadOnly Property IsRelative() As Boolean
		Get
			IsRelative = CBool(PathIsRelative(sPath))
		End Get
	End Property
	
	' Parses a path to determine if it is a directory root.
	Public ReadOnly Property IsRoot() As Boolean
		Get
			IsRoot = CBool(PathIsRoot(sPath))
		End Get
	End Property
	
	' Determines if the string is a valid UNC (universal naming convention)
	' for a server and share path.
	Public ReadOnly Property IsUNC() As Boolean
		Get
			IsUNC = CBool(PathIsUNC(sPath))
		End Get
	End Property
	
	' Determines if a string is a valid UNC (universal naming convention)
	' for a server path only.
	Public ReadOnly Property IsUNCServer() As Boolean
		Get
			IsUNCServer = CBool(PathIsUNCServer(sPath))
		End Get
	End Property
	
	' Determines if a string is a valid universal naming convention (UNC)
	' share path, \\server\share.
	Public ReadOnly Property IsUNCServerShare() As Boolean
		Get
			IsUNCServerShare = CBool(PathIsUNCServerShare(sPath))
		End Get
	End Property
	
	' Tests a given string to determine if it conforms to a valid URL format.
	Public ReadOnly Property IsURL() As Boolean
		Get
			IsURL = CBool(PathIsURL(sPath))
		End Get
	End Property
	
	' Set/returns the current path.
	
	Public Property Path() As String
		Get
			Path = TrimNull(sPath)
		End Get
		Set(ByVal Value As String)
			sPath = Left(Trim(Value) & New String(Chr(0), MAX_PATH), MAX_PATH)
		End Set
	End Property
	
	' Searches a path to determine if it contains a valid prefix of the
	' type passed by sPrefix. (e.g. "C:\", ".", "..", "..\".)
	Public Function IsPrefix(ByVal Prefix As String) As Boolean
		IsPrefix = CBool(PathIsPrefix(Prefix, sPath))
	End Function
	
	' Indicates wether or not the path matches the spec.
	Public Function MatchSpec(ByVal Spec As String) As Boolean
		MatchSpec = CBool(PathMatchSpec(sPath, Spec))
	End Function
	
	' Searches a path for spaces. If spaces are found, the entire path is
	' enclosed in quotation marks.
	Public Function QuotePath() As String
		Call PathQuoteSpaces(sPath)
		QuotePath = Path
	End Function
	
	' Removes the trailing backslash from a given path.
	Public Function RemoveBackslash() As String
		PathRemoveBackslash(sPath)
		RemoveBackslash = Path
	End Function
	
	' Removes the file extension from a path, if there is one.
	Public Function RemoveExtension() As String
		Call PathRemoveExtension(sPath)
		RemoveExtension = Path
	End Function
	
	' Removes the trailing file name and backslash from a path, if it
	' has them.
	Public Function RemoveFilename() As String
		Call PathRemoveFileSpec(sPath)
		RemoveFilename = Path
	End Function
	
	' Removes the directory portion of a fully qualified path.
	Public Function RemoveDirectory() As String
		Call PathStripPath(sPath)
		RemoveDirectory = Path
	End Function
	
	
	' Replaces the extension of a file name with a new extension.
	' If the file name does not contain an extension, the extension will be
	' attached to the end of the string.
	Public Function RenameExtension(ByVal NewExtension As String) As String
		Call PathRenameExtension(sPath, NewExtension)
		RenameExtension = Path
	End Function
	
	' Parses a path, ignoring the drive letter or UNC server/share path
	' parts.
	Public Function SkipRoot() As String
		Dim lpszRoot As Integer
		Dim bBuffer(MAX_PATH) As Byte
		Dim sTemp As String
		
		lpszRoot = PathSkipRoot(sPath)
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		Call CopyMemoryLpToLp(VarPtr(bBuffer(0)), lpszRoot, lstrlenptr(lpszRoot))
		
		'UPGRADE_ISSUE: Constant vbUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		sTemp = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(bBuffer), vbUnicode)
		SkipRoot = TrimNull(sTemp)
	End Function
	
	Private Function TrimNull(ByVal szSource As String) As String
		Dim Pos As Short
		Pos = InStr(1, szSource, Chr(0), CompareMethod.Binary)
		If Pos > 1 Then
			TrimNull = Left(szSource, Pos - 1)
		Else
			TrimNull = ""
		End If
	End Function
	
	' Removes quotes from the beginning and end of a path.
	Public Function UnquotePath() As String
		Call PathUnquoteSpaces(sPath)
		UnquotePath = Path
	End Function
End Class