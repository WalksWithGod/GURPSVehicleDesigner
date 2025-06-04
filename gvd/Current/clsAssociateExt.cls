VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAssociateExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
  (ByVal hKey As Long, ByVal lpSubKey As String) As Long
  
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ = 1

Dim aDesc As String
Dim aError As String
Dim aExt As String
Dim aIcon As String
Dim aOpen As String
Dim aPrint As String
Dim aRootFile As String

Public Function DeleteGVDAssociation()

   Dim e As Long
   On Error Resume Next
   
    e = RegDeleteKey(HKEY_CLASSES_ROOT, ".veh")
    e = RegDeleteKey(HKEY_CLASSES_ROOT, "vehfile")
    e = RegDeleteKey(HKEY_CLASSES_ROOT, "veh_auto_file")
    
    
End Function

Public Function DrawAssociatedIcon(h As Long, ByVal x As Long, ByVal y As Long) As Boolean
    DrawAssociatedIcon = False
    If Len(aIcon) Then
        Dim aIconFile As String, k As Long, hIcon As Long
        ' %SystemRoot%\system32\shell32.dll,-152
        k = InStr(aIcon, ",")
        If k Then
            aIconFile = Left$(aIcon, k - 1)
            k = Val(Mid$(aIcon, k + 1))
            hIcon = ExtractIcon(0, aIconFile$, k)
            If hIcon Then
                x = x \ Screen.TwipsPerPixelX
                y = y \ Screen.TwipsPerPixelY
                DrawIcon h, x, y, hIcon
                DeleteObject hIcon
                DrawAssociatedIcon = True
            End If
        End If
    End If
End Function

Private Sub GetRootFile()
    If aExt = "" Then Exit Sub
    
    Dim aKey$, lRet As Long, hKey As Long, lSize As Long
    aRootFile = ""
    aKey$ = aExt
    lRet = RegOpenKeyEx(HKEY_CLASSES_ROOT, aKey$, 0, KEY_ALL_ACCESS, hKey)
    If hKey Then
        aRootFile = Space$(255): lSize = 255
        lRet = RegQueryValueEx(hKey, "", 0&, REG_SZ, ByVal aRootFile, lSize)
        lRet = RegCloseKey(hKey)
        lSize = InStr(aRootFile, Chr$(0))
        If lSize Then
            aRootFile = Left$(aRootFile, lSize - 1)
        Else
            aRootFile = ""
        End If
    End If
    If aRootFile = "" Then
        aRootFile = Mid$(aExt, 2) & "file"
        SetKey aExt$, aRootFile
    End If
End Sub

Private Sub SetKey(aKey$, aDat$)
    Dim lRet As Long, hKey As Long, lDisp As Long
    lRet = RegCreateKeyEx(HKEY_CLASSES_ROOT, aKey$, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, hKey, lDisp)
    If hKey Then
        If Len(aDat$) Then
            lRet = RegSetValueEx(hKey, "", 0&, REG_SZ, ByVal aDat$, Len(aDat$))
        Else
            lRet = RegDeleteValue(hKey, "")
        End If
        lRet = RegCloseKey(hKey)
    End If
End Sub

Public Sub SetAssociation()
    aError = ""
    If aExt = "" Then aError = "No extension specified for {GetAssocation} method.": Exit Sub
    Dim hKey As Long, lRet As Long, aKey$, aDat$, lSize As Long, lDisp As Long
    
    Call GetRootFile
    
    aKey$ = aRootFile
    Call SetKey(aKey$, aDesc)
    
    aKey$ = aRootFile & "\DefaultIcon"
    Call SetKey(aKey$, aIcon)
    
    aKey$ = aRootFile & "\shell\open\command"
    Call SetKey(aKey$, aOpen)

    'aKey$ = aRootFile & "\shell\print\command"
    'Call SetKey(aKey$, aPrint)
End Sub

Private Function GetKey(aKey$) As String
    Dim lRet As Long, lSize As Long, hKey As Long, aDat$
    lRet = RegOpenKeyEx(HKEY_CLASSES_ROOT, aKey$, 0, KEY_ALL_ACCESS, hKey)
    If hKey Then
        aDat$ = Space$(255): lSize = 255
        lRet = RegQueryValueEx(hKey, "", 0&, REG_SZ, ByVal aDat$, lSize)
        lRet = RegCloseKey(hKey)
        lSize = InStr(aDat$, Chr$(0))
        If lSize Then
            aDat$ = Left$(aDat$, lSize - 1)
        Else
            aDat$ = ""
        End If
    End If
    GetKey = aDat$
End Function

Public Sub GetAssociation()
    aError = ""
    If aExt = "" Then aError = "No extension specified for {GetAssocation} method.": Exit Sub
    Dim hKey As Long, lRet As Long, aKey$, lSize As Long
    aDesc = ""
    aIcon = ""
    aOpen = ""
    aPrint = ""
    
    Call GetRootFile
    
    aKey$ = aRootFile
    aDesc = GetKey(aKey$)
    
    aKey$ = aRootFile & "\DefaultIcon"
    aIcon = GetKey(aKey$)
    
    aKey$ = aRootFile & "\shell\open\command"
    aOpen = GetKey(aKey$)

    aKey$ = aRootFile & "\shell\print\command"
    aPrint = GetKey(aKey$)
End Sub

Public Property Get LastError() As String
    LastError = aError
End Property

Public Property Get Description() As String
    Description = aDesc
End Property

Public Property Let Description(ByVal Param As String)
    aDesc = Param
End Property

Public Property Get Extension() As String
    Extension = aExt
End Property

Public Property Let Extension(ByVal Param As String)
    aExt = LCase$(Param)
    If Len(aExt) Then
        If Left$(aExt, 1) <> "." Then aExt = "." & aExt
    End If
End Property

Public Property Get DefaultIcon() As String
    DefaultIcon = aIcon
End Property

Public Property Let DefaultIcon(ByVal Param As String)
    aIcon = Param
End Property

Public Property Get OpenCommand() As String
    OpenCommand = aOpen
End Property

Public Property Let OpenCommand(ByVal Param As String)
    aOpen = Param
End Property

Public Property Get PrintCommand() As String
    PrintCommand = aPrint
End Property

Public Property Let PrintCommand(ByVal Param As String)
    aPrint = Param
End Property

