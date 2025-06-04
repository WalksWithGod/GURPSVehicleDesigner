Option Strict Off
Option Explicit On
Module modFileUtils
	
	' old code..
	'Function IsValidName(ByRef sName As String) As Long
	'    Dim lngRet As Long
	'    Dim sCharacter As String
	'    Dim i As Long
	'
	'    lngRet = True
	'
	'    If sName <> "" Then
	'        ' check for reserved characters
	'        For i = 1 To Len(sName)
	'            sCharacter = Mid(sName, i, 1)
	'            If (sCharacter = "|") Or (sCharacter = "@") Or (sCharacter = "[") Or (sCharacter = "]") Then
	'                InfoPrint 2, "'" & sCharacter & "' is a reserved character."
	'                lngRet = False
	'                Exit For
	'            End If
	'        Next
	'    Else
	'        lngRet = False
	'    End If
	'    IsValidName = lngRet
	'End Function
	
	Public Function IsValidFilename(ByVal sFilename As String) As Boolean
		Dim i As Short
		Dim j As Short
		Dim sTemp As String
		Dim strErrorChars As String 'Illegal characters in a filename / Directory
		Dim nMaxLength As Short
		
		IsValidFilename = False 'Default to false
		
		If sFilename = "" Then Exit Function
		If sFilename = "." Then Exit Function
		
		nMaxLength = 255 'Windows 2000 appears to be limited to 255 characters
		sTemp = sFilename
		
		strErrorChars = "\/:*?<>|" & Chr(34) & vbTab
		
		i = InStr(1, sTemp, ":", CompareMethod.Text)
		If i = 2 Then
			
			'If the filename contains a : the : must be preceded by a letter and followed by a \
			If Len(sTemp) < 4 Then Exit Function 'eg c:\
			
			If Not Mid(sTemp, 1, 1) Like "[A-Za-z]" Then Exit Function
			If Not Mid(sTemp, 3, 1) = "\" Then Exit Function
			
			sTemp = Mid(sTemp, 4)
			
		ElseIf i > 2 Then 
			Exit Function
			
		ElseIf i <> 0 Then 
			Exit Function
			
		End If
		
		'Check  any directory names for validity
		j = InStr(sTemp, "\")
		While j > 0
			If j > nMaxLength Then Exit Function 'Guard against a directory being to long
			'Ensure the directory name doesnt contain an invalid character
			For i = 1 To j - 1
				If InStr(1, strErrorChars, Mid(sTemp, i, 1), CompareMethod.Text) Then Exit Function
			Next i
			
			sTemp = Mid(sTemp, j + 1)
			j = InStr(sTemp, "\")
		End While
		
		'Finally check that the filename is valid
		j = Len(sTemp)
		
		If j > nMaxLength Then Exit Function
		
		'Ensure filename doesnt contain invalid characters
		For i = 1 To j
			If InStr(1, strErrorChars, Mid(sTemp, i, 1), CompareMethod.Text) Then Exit Function
		Next i
		
		IsValidFilename = True
		
	End Function
	
	Public Sub assertFileNames()
		System.Diagnostics.Debug.Assert(Not IsValidFilename("1:"), "")
		System.Diagnostics.Debug.Assert(Not IsValidFilename("C:"), "")
		System.Diagnostics.Debug.Assert(Not IsValidFilename("2:\filename.txt"), "")
		System.Diagnostics.Debug.Assert(IsValidFilename("C:\filename.txt"), "")
		System.Diagnostics.Debug.Assert(IsValidFilename("filename.txt"), "")
		System.Diagnostics.Debug.Assert(IsValidFilename("C:\filename\1111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111.txt"), "")
		System.Diagnostics.Debug.Assert(Not IsValidFilename("C:\filename\filename*.txt"), "")
		System.Diagnostics.Debug.Assert(Not IsValidFilename("C:filename.txt"), "")
		System.Diagnostics.Debug.Assert(Not IsValidFilename("c:\"), "")
		System.Diagnostics.Debug.Assert(Not IsValidFilename("q*"), "")
		System.Diagnostics.Debug.Assert(Not IsValidFilename(""), "")
		System.Diagnostics.Debug.Assert(Not IsValidFilename("."), "")
	End Sub
End Module