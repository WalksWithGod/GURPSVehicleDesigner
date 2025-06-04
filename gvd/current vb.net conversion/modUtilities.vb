Option Strict Off
Option Explicit On
Module modUtilities
	
	'//////the below is for launching the default text viewer and default web browsers
	Private Declare Function FindExecutable Lib "shell32.dll"  Alias "FindExecutableA"(ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Integer
	
	Private Declare Function ShellExecute Lib "shell32.dll"  Alias "ShellExecuteA"(ByVal hWnd As Integer, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Integer) As Integer
	
	Private Declare Function GetDesktopWindow Lib "user32" () As Integer ' used for launching associated viewers
	
	Function Maximum(ByVal x As Object, ByVal y As Object) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Maximum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If x > y Then
			Maximum = x
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Maximum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Maximum = y
		End If
	End Function
	
	Function Minimum(ByVal x As Object, ByVal y As Object) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Minimum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If x < y Then
			Minimum = x
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Minimum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Minimum = y
		End If
	End Function
	
	Sub InfoPrint(ByVal Code As Short, ByVal Message As String)
		Dim frmDesigner As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.txtInfo. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.txtInfo.Text = Message & vbNewLine & Left(frmDesigner.txtInfo.Text, 2000)
	End Sub
	
	Function ExtractPathFromFile(ByRef s As String) As String
		
		Dim i As Integer
		Dim sRet As String
		
		For i = Len(s) To 1 Step -1
			If Mid(s, i, 1) = "\" Then
				sRet = Left(s, i)
				Exit For
			End If
		Next 
		
		ExtractPathFromFile = sRet
		
	End Function
	
	Function ExtractFileNameFromPath(ByRef s As String) As String
		Dim i As Integer
		Dim j As Integer
		
		' get the actual filename from the filepath
		For i = Len(s) To 1 Step -1
			j = j + 1
			If Mid(s, i, 1) = "\" Then
				ExtractFileNameFromPath = Right(s, j - 1)
				Exit For
			End If
		Next 
		
	End Function
	
	
	Function StartDoc(ByRef DocName As String) As Integer
		'this function launches the text file using the associated viewer
		Dim Scr_hDC As Integer
		
		Scr_hDC = GetDesktopWindow()
		StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", SW_SHOWNORMAL)
	End Function
	
	Function FindDefaultProgram(ByVal sType As String) As String
		
		Dim FileName As Object
		Dim Dummy As String
		Dim BrowserExec As New VB6.FixedLengthString(255)
		Dim retval As Integer
		Dim FileNumber As Short
		
		' First, create a known, temporary file
		
		BrowserExec.Value = Space(255)
		
		If sType = "Text" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object FileName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileName = My.Application.Info.DirectoryPath & "\tmp00000001.txt"
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object FileName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileName = My.Application.Info.DirectoryPath & "\tmp00000001.HTM"
		End If
		
		FileNumber = FreeFile ' Get unused file number
		'UPGRADE_WARNING: Couldn't resolve default property of object FileName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(FileNumber, FileName, OpenMode.Output) ' Create temp HTML file
		WriteLine(FileNumber, "<HTML> <\HTML>") ' Output text
		FileClose(FileNumber) ' Close file
		
		' Then find the application associated with it
		'UPGRADE_WARNING: Couldn't resolve default property of object FileName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		retval = FindExecutable(FileName, Dummy, BrowserExec.Value)
		BrowserExec.Value = Trim(BrowserExec.Value)
		
		'delete the temp file
		If sType = "Text" Then
			Kill(My.Application.Info.DirectoryPath & "\tmp00000001.txt")
		Else
			Kill(My.Application.Info.DirectoryPath & "\tmp00000001.HTM")
		End If
		
		FindDefaultProgram = BrowserExec.Value
	End Function
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''
	' 02/16/02 MPJ No longer need this function since it was only used in splash screen
	' which is now just a simpler splash without animation
	'Sub Pause(ByVal Interval As Single)
	'Dim Start As Single
	'
	'    Start = Timer   ' Set start time.
	'    Do While Timer < Start + Interval
	'        DoEvents    ' Yield to other processes.
	'    Loop
	'
	'End Sub
	'''''''''''''''''''''''''''''''''''''''''''''''''''
End Module