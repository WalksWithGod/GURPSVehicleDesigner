VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Private Const MAX_SECTION_SIZE = 2000
Dim m_sFileName As String

'----------------------------------------------------------------------------
'NOTE:  This class requires the project its in to reference the "Microsoft Scripting Runtime" (scrrun.dll)
'------------------------------------------------------------------------------


' TODO:  This class needs a flag for class initialization e.g. m_bLibraryInitialied
'        which should be checked when any external call is made.
Public Property Get FileName() As String
    FileName = m_sFileName
End Property

Public Property Let FileName(ByVal vNewValue As String)
    '//here the user sets the name of the INI file
    '//we are working with
    m_sFileName = vNewValue
    
    '//first check to make sure this file exists.
    '//if it doesnt, create it
    Dim oFile As FileSystemObject
    Set oFile = New FileSystemObject
    
    If Not oFile.FileExists(m_sFileName) Then
        '//todo: first needs to build the path or error ensues
        
        oFile.CreateTextFile m_sFileName
    End If
    Set oFile = Nothing
End Property

Public Sub DeleteSection(ByVal sSectionName As String)
    '//this accepts a section name and deletes it and all keys in it
    Dim iFree As Long
    Dim sArray() As String
    Dim s As String
    Dim i As Long
    Dim j As Long
    
    iFree = FreeFile
    ReDim sArray(0)
    i = 0
    
    Open m_sFileName For Input As #iFree
    
    Do While Not EOF(iFree)
        Line Input #iFree, s
        If s <> "" Then
            ReDim Preserve sArray(i)
            sArray(i) = s
            i = i + 1
        End If
    Loop
    Close #iFree
    
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
    
    Open m_sFileName For Output As #iFree
    
    For i = 0 To UBound(sArray)
        If sArray(i) <> "" Then
            Print #iFree, sArray(i)
        End If
    Next
    Close #iFree
End Sub

Public Sub WriteSection(ByVal sSectionName As String, ParamArray vList())
    '//this writes an entire section of data
    '//the user provides the section name and the list of keys
    '//NOTE: This does not write in any Key Values...
    
    Dim sArray() As String
    Dim s As String
    Dim i As Long
    
    For i = 0 To UBound(vList)
        sArray = Split(vList(i), "=")
        s = s & sArray(0) & Chr(10) & sArray(1) & Chr(10) & Chr(10)
    Next
    
    Call WritePrivateProfileSection(sSectionName, s, m_sFileName)
    
End Sub

Public Function GetNextIndexedSection(ByVal sSectionName As String) As String
    '//this function creates indexed section names.  It finds the first available section
    '//and creates it and returns the name of the section it created
    Dim i, j, k, l As Long
    Dim iFree As Long
    Dim sName As String
    Dim s As String
    Dim sArray() As String
    
    iFree = FreeFile
    ReDim sArray(0)
    i = 0
    j = 0
    k = 0
    l = 0
    
    Open m_sFileName For Input As #iFree
    '//get a list of all sections within this file
    Do While Not EOF(iFree)
        Line Input #iFree, s
        If Left(s, 1) = "[" Then
            ReDim Preserve sArray(i)
            sArray(i) = Mid(s, 2, Len(s) - 2)
            i = i + 1
        End If
    Loop
    Close #iFree
    
    '//now find a free section
    sName = sSectionName & k
    For i = 0 To UBound(sArray)
        For j = 0 To UBound(sArray)
            If sName = sArray(j) Then
                l = l + 1
                Exit For
            End If
        Next
        If l = k Then Exit For
        k = k + 1
        sName = sSectionName & k
    Next
    '//create the current section
    GetNextIndexedSection = sName
End Function

Public Sub WriteString(ByVal sSectionName As String, ByVal sKeyName As String, ByVal s As String)
    '//accepts a string and writes it to the ini file
     Call WritePrivateProfileString(sSectionName, sKeyName, s, m_sFileName)
End Sub

Public Sub WriteInteger(ByVal sSectionName As String, ByVal sKeyName As String, ByVal l As Long)
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

Public Function ReadString(ByVal sSectionName As String, sKeyName As String) As String
    '//accepts the Section and Keyname and returns the value of the key
    Dim sReturnString As String
    sReturnString = Space(255)
    ReadString = Left(sReturnString, GetPrivateProfileString(sSectionName, sKeyName, " ", sReturnString, 255, m_sFileName))
End Function

Public Function ReadInteger(ByVal sSectionName As String, sKeyName As String) As Long
    '//accepts the Section and Keyname and returns the integer value of the key
    ReadInteger = GetPrivateProfileInt(sSectionName, sKeyName, 0, m_sFileName)
End Function

Public Function ReadSection(ByVal sSectionName As String) As String()
    '//this function accepts the name of a section and then retreives
    '//all of the keynames
    
    Dim i, j As Long
    Dim s As String
    Dim sRet() As String
    Dim lngNullOffset As Long
    Dim lngKeyValueOffset As Long
    
    On Error GoTo errorhandler
    
    s = String(MAX_SECTION_SIZE, 0) '//allocate space in our string
    j = 0
    
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
            
    ReadSection = sRet
    Exit Function
errorhandler:
    Exit Function
End Function

Public Function RetreiveSectionNames() As String()
    '//this function opens the INI and returns the list of
    '//all the section names in an string array
    Dim iFileNum As Long
    iFileNum = FreeFile
    Dim s As String
    Dim sRet() As String
    Dim i As Long
    
    Open m_sFileName For Input As #iFileNum
    
    Do While Not EOF(iFileNum)
        Line Input #iFileNum, s
        If Left(s, 1) = "[" Then
            ReDim Preserve sRet(i)
            s = Mid(s, 2, Len(s) - 2)
            sRet(i) = s
            i = i + 1
        End If
    Loop
    Close iFileNum
    RetreiveSectionNames = sRet
End Function
