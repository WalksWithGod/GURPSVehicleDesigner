Option Strict Off
Option Explicit On
Friend Class clsAssociateExt
	
	
	Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer
	Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Integer) As Integer
	Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal hIcon As Integer) As Integer
	Private Declare Function ExtractIcon Lib "shell32.dll"  Alias "ExtractIconA"(ByVal hInst As Integer, ByVal lpszExeFileName As String, ByVal nIconIndex As Integer) As Integer
	
	Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
	Private Declare Function RegCreateKeyEx Lib "advapi32.dll"  Alias "RegCreateKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal Reserved As Integer, ByVal lpClass As String, ByVal dwOptions As Integer, ByVal samDesired As Integer, ByVal lpSecurityAttributes As Integer, ByRef phkResult As Integer, ByRef lpdwDisposition As Integer) As Integer
	Private Declare Function RegDeleteValue Lib "advapi32.dll"  Alias "RegDeleteValueA"(ByVal hKey As Integer, ByVal lpValueName As String) As Integer
	Private Declare Function RegOpenKeyEx Lib "advapi32.dll"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function RegQueryValueEx Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Any, ByRef lpcbData As Integer) As Integer ' Note that if you declare the lpData parameter as String, you must pass it By Value.
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function RegSetValueEx Lib "advapi32.dll"  Alias "RegSetValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal Reserved As Integer, ByVal dwType As Integer, ByRef lpData As Any, ByVal cbData As Integer) As Integer ' Note that if you declare the lpData parameter as String, you must pass it By Value.
	Private Declare Function RegDeleteKey Lib "advapi32.dll"  Alias "RegDeleteKeyA"(ByVal hKey As Integer, ByVal lpSubKey As String) As Integer
	
	Private Const HKEY_CLASSES_ROOT As Integer = &H80000000
	Private Const KEY_ALL_ACCESS As Short = &H3Fs
	Private Const REG_OPTION_NON_VOLATILE As Short = 0
	Private Const REG_SZ As Short = 1
	
	Dim aDesc As String
	Dim aError As String
	Dim aExt As String
	Dim aIcon As String
	Dim aOpen As String
	Dim aPrint As String
	Dim aRootFile As String
	
	Public Function DeleteGVDAssociation() As Object
		
		Dim e As Integer
		On Error Resume Next
		
		e = RegDeleteKey(HKEY_CLASSES_ROOT, ".veh")
		e = RegDeleteKey(HKEY_CLASSES_ROOT, "vehfile")
		e = RegDeleteKey(HKEY_CLASSES_ROOT, "veh_auto_file")
		
		
	End Function
	
	Public Function DrawAssociatedIcon(ByRef h As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
		DrawAssociatedIcon = False
		Dim aIconFile As String
		Dim k, hIcon As Integer
		If Len(aIcon) Then
			' %SystemRoot%\system32\shell32.dll,-152
			k = InStr(aIcon, ",")
			If k Then
				aIconFile = Left(aIcon, k - 1)
				k = Val(Mid(aIcon, k + 1))
				hIcon = ExtractIcon(0, aIconFile, k)
				If hIcon Then
					x = x \ VB6.TwipsPerPixelX
					y = y \ VB6.TwipsPerPixelY
					DrawIcon(h, x, y, hIcon)
					DeleteObject(hIcon)
					DrawAssociatedIcon = True
				End If
			End If
		End If
	End Function
	
	Private Sub GetRootFile()
		If aExt = "" Then Exit Sub
		
		Dim aKey As String
		Dim hKey, lRet, lSize As Integer
		aRootFile = ""
		aKey = aExt
		lRet = RegOpenKeyEx(HKEY_CLASSES_ROOT, aKey, 0, KEY_ALL_ACCESS, hKey)
		If hKey Then
			aRootFile = Space(255) : lSize = 255
			lRet = RegQueryValueEx(hKey, "", 0, REG_SZ, aRootFile, lSize)
			lRet = RegCloseKey(hKey)
			lSize = InStr(aRootFile, Chr(0))
			If lSize Then
				aRootFile = Left(aRootFile, lSize - 1)
			Else
				aRootFile = ""
			End If
		End If
		If aRootFile = "" Then
			aRootFile = Mid(aExt, 2) & "file"
			SetKey(aExt, aRootFile)
		End If
	End Sub
	
	Private Sub SetKey(ByRef aKey As String, ByRef aDat As String)
		Dim hKey, lRet, lDisp As Integer
		lRet = RegCreateKeyEx(HKEY_CLASSES_ROOT, aKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, hKey, lDisp)
		If hKey Then
			If Len(aDat) Then
				lRet = RegSetValueEx(hKey, "", 0, REG_SZ, aDat, Len(aDat))
			Else
				lRet = RegDeleteValue(hKey, "")
			End If
			lRet = RegCloseKey(hKey)
		End If
	End Sub
	
	Public Sub SetAssociation()
		aError = ""
		If aExt = "" Then aError = "No extension specified for {GetAssocation} method." : Exit Sub
		Dim lSize, hKey, lRet, lDisp As Integer
		Dim aKey, aDat As String
		
		Call GetRootFile()
		
		aKey = aRootFile
		Call SetKey(aKey, aDesc)
		
		aKey = aRootFile & "\DefaultIcon"
		Call SetKey(aKey, aIcon)
		
		aKey = aRootFile & "\shell\open\command"
		Call SetKey(aKey, aOpen)
		
		'aKey$ = aRootFile & "\shell\print\command"
		'Call SetKey(aKey$, aPrint)
	End Sub
	
	Private Function GetKey(ByRef aKey As String) As String
		Dim lSize, lRet, hKey As Integer
		Dim aDat As String
		lRet = RegOpenKeyEx(HKEY_CLASSES_ROOT, aKey, 0, KEY_ALL_ACCESS, hKey)
		If hKey Then
			aDat = Space(255) : lSize = 255
			lRet = RegQueryValueEx(hKey, "", 0, REG_SZ, aDat, lSize)
			lRet = RegCloseKey(hKey)
			lSize = InStr(aDat, Chr(0))
			If lSize Then
				aDat = Left(aDat, lSize - 1)
			Else
				aDat = ""
			End If
		End If
		GetKey = aDat
	End Function
	
	Public Sub GetAssociation()
		aError = ""
		If aExt = "" Then aError = "No extension specified for {GetAssocation} method." : Exit Sub
		Dim lRet, hKey, lSize As Integer
		Dim aKey As String
		aDesc = ""
		aIcon = ""
		aOpen = ""
		aPrint = ""
		
		Call GetRootFile()
		
		aKey = aRootFile
		aDesc = GetKey(aKey)
		
		aKey = aRootFile & "\DefaultIcon"
		aIcon = GetKey(aKey)
		
		aKey = aRootFile & "\shell\open\command"
		aOpen = GetKey(aKey)
		
		aKey = aRootFile & "\shell\print\command"
		aPrint = GetKey(aKey)
	End Sub
	
	Public ReadOnly Property LastError() As String
		Get
			LastError = aError
		End Get
	End Property
	
	
	Public Property Description() As String
		Get
			Description = aDesc
		End Get
		Set(ByVal Value As String)
			aDesc = Value
		End Set
	End Property
	
	
	Public Property Extension() As String
		Get
			Extension = aExt
		End Get
		Set(ByVal Value As String)
			aExt = LCase(Value)
			If Len(aExt) Then
				If Left(aExt, 1) <> "." Then aExt = "." & aExt
			End If
		End Set
	End Property
	
	
	Public Property DefaultIcon() As String
		Get
			DefaultIcon = aIcon
		End Get
		Set(ByVal Value As String)
			aIcon = Value
		End Set
	End Property
	
	
	Public Property OpenCommand() As String
		Get
			OpenCommand = aOpen
		End Get
		Set(ByVal Value As String)
			aOpen = Value
		End Set
	End Property
	
	
	Public Property PrintCommand() As String
		Get
			PrintCommand = aPrint
		End Get
		Set(ByVal Value As String)
			aPrint = Value
		End Set
	End Property
End Class