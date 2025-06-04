Attribute VB_Name = "modFileUtils"
Option Explicit

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
vbwProfiler.vbwProcIn 481
Dim i As Integer
Dim j As Integer
Dim sTemp As String
Dim strErrorChars As String   'Illegal characters in a filename / Directory
Dim nMaxLength As Integer

vbwProfiler.vbwExecuteLine 10183
  IsValidFilename = False     'Default to false

vbwProfiler.vbwExecuteLine 10184
  If sFilename = "" Then
vbwProfiler.vbwProcOut 481
vbwProfiler.vbwExecuteLine 10185
       Exit Function
  End If
vbwProfiler.vbwExecuteLine 10186 'B
vbwProfiler.vbwExecuteLine 10187
  If sFilename = "." Then
vbwProfiler.vbwProcOut 481
vbwProfiler.vbwExecuteLine 10188
       Exit Function
  End If
vbwProfiler.vbwExecuteLine 10189 'B

vbwProfiler.vbwExecuteLine 10190
  nMaxLength = 255            'Windows 2000 appears to be limited to 255 characters
vbwProfiler.vbwExecuteLine 10191
  sTemp = sFilename

vbwProfiler.vbwExecuteLine 10192
  strErrorChars = "\/:*?<>|" & Chr(34) & vbTab

vbwProfiler.vbwExecuteLine 10193
  i = InStr(1, sTemp, ":", vbTextCompare)
vbwProfiler.vbwExecuteLine 10194
  If i = 2 Then

    'If the filename contains a : the : must be preceded by a letter and followed by a \
vbwProfiler.vbwExecuteLine 10195
    If Len(sTemp) < 4 Then 'eg c:\
vbwProfiler.vbwProcOut 481
vbwProfiler.vbwExecuteLine 10196
         Exit Function
    End If
vbwProfiler.vbwExecuteLine 10197 'B

vbwProfiler.vbwExecuteLine 10198
    If Not Mid(sTemp, 1, 1) Like "[A-Za-z]" Then
vbwProfiler.vbwProcOut 481
vbwProfiler.vbwExecuteLine 10199
         Exit Function
    End If
vbwProfiler.vbwExecuteLine 10200 'B
vbwProfiler.vbwExecuteLine 10201
    If Not Mid(sTemp, 3, 1) = "\" Then
vbwProfiler.vbwProcOut 481
vbwProfiler.vbwExecuteLine 10202
         Exit Function
    End If
vbwProfiler.vbwExecuteLine 10203 'B

vbwProfiler.vbwExecuteLine 10204
    sTemp = Mid(sTemp, 4)

'vbwLine 10205:  ElseIf i > 2 Then
  ElseIf vbwProfiler.vbwExecuteLine(10205) Or i > 2 Then
vbwProfiler.vbwProcOut 481
vbwProfiler.vbwExecuteLine 10206
    Exit Function

vbwProfiler.vbwProcOut 481
'vbwLine 10207:  ElseIf i <> 0 Then Exit Function
  ElseIf vbwProfiler.vbwExecuteLine(10207) Or i <> 0 Then Exit Function

  End If
vbwProfiler.vbwExecuteLine 10208 'B

  'Check  any directory names for validity
vbwProfiler.vbwExecuteLine 10209
  j = InStr(sTemp, "\")
'vbwLine 10210:  While j > 0
  While vbwProfiler.vbwExecuteLine(10210) Or j > 0
vbwProfiler.vbwExecuteLine 10211
    If j > nMaxLength Then 'Guard against a directory being to long
vbwProfiler.vbwProcOut 481
vbwProfiler.vbwExecuteLine 10212
         Exit Function
    End If
vbwProfiler.vbwExecuteLine 10213 'B
    'Ensure the directory name doesnt contain an invalid character
vbwProfiler.vbwExecuteLine 10214
    For i = 1 To j - 1
vbwProfiler.vbwExecuteLine 10215
      If InStr(1, strErrorChars, Mid(sTemp, i, 1), vbTextCompare) Then
vbwProfiler.vbwProcOut 481
vbwProfiler.vbwExecuteLine 10216
           Exit Function
      End If
vbwProfiler.vbwExecuteLine 10217 'B
vbwProfiler.vbwExecuteLine 10218
    Next i

vbwProfiler.vbwExecuteLine 10219
    sTemp = Mid(sTemp, j + 1)
vbwProfiler.vbwExecuteLine 10220
    j = InStr(sTemp, "\")
vbwProfiler.vbwExecuteLine 10221
  Wend

  'Finally check that the filename is valid
vbwProfiler.vbwExecuteLine 10222
  j = Len(sTemp)

vbwProfiler.vbwExecuteLine 10223
  If j > nMaxLength Then
vbwProfiler.vbwProcOut 481
vbwProfiler.vbwExecuteLine 10224
       Exit Function
  End If
vbwProfiler.vbwExecuteLine 10225 'B

  'Ensure filename doesnt contain invalid characters
vbwProfiler.vbwExecuteLine 10226
  For i = 1 To j
vbwProfiler.vbwExecuteLine 10227
    If InStr(1, strErrorChars, Mid(sTemp, i, 1), vbTextCompare) Then
vbwProfiler.vbwProcOut 481
vbwProfiler.vbwExecuteLine 10228
         Exit Function
    End If
vbwProfiler.vbwExecuteLine 10229 'B
vbwProfiler.vbwExecuteLine 10230
  Next i

vbwProfiler.vbwExecuteLine 10231
  IsValidFilename = True

vbwProfiler.vbwProcOut 481
vbwProfiler.vbwExecuteLine 10232
End Function

Public Sub assertFileNames()
vbwProfiler.vbwProcIn 482
vbwProfiler.vbwExecuteLine 10233
  Debug.Assert Not IsValidFilename("1:")
vbwProfiler.vbwExecuteLine 10234
  Debug.Assert Not IsValidFilename("C:")
vbwProfiler.vbwExecuteLine 10235
  Debug.Assert Not IsValidFilename("2:\filename.txt")
vbwProfiler.vbwExecuteLine 10236
  Debug.Assert IsValidFilename("C:\filename.txt")
vbwProfiler.vbwExecuteLine 10237
  Debug.Assert IsValidFilename("filename.txt")
vbwProfiler.vbwExecuteLine 10238
  Debug.Assert IsValidFilename("C:\filename\1111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111.txt")
vbwProfiler.vbwExecuteLine 10239
  Debug.Assert Not IsValidFilename("C:\filename\filename*.txt")
vbwProfiler.vbwExecuteLine 10240
  Debug.Assert Not IsValidFilename("C:filename.txt")
vbwProfiler.vbwExecuteLine 10241
  Debug.Assert Not IsValidFilename("c:\")
vbwProfiler.vbwExecuteLine 10242
  Debug.Assert Not IsValidFilename("q*")
vbwProfiler.vbwExecuteLine 10243
  Debug.Assert Not IsValidFilename("")
vbwProfiler.vbwExecuteLine 10244
  Debug.Assert Not IsValidFilename(".")
vbwProfiler.vbwProcOut 482
vbwProfiler.vbwExecuteLine 10245
End Sub

