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
vbwProfiler.vbwProcIn 146
vbwProfiler.vbwExecuteLine 3528
    FileName = m_sFileName
vbwProfiler.vbwProcOut 146
vbwProfiler.vbwExecuteLine 3529
End Property

Public Property Let FileName(ByVal vNewValue As String)
    '//here the user sets the name of the INI file
    '//we are working with
vbwProfiler.vbwProcIn 147
vbwProfiler.vbwExecuteLine 3530
    m_sFileName = vNewValue

    '//first check to make sure this file exists.
    '//if it doesnt, create it
    Dim oFile As FileSystemObject
vbwProfiler.vbwExecuteLine 3531
    Set oFile = New FileSystemObject

vbwProfiler.vbwExecuteLine 3532
    If Not oFile.FileExists(m_sFileName) Then
        '//todo: first needs to build the path or error ensues

vbwProfiler.vbwExecuteLine 3533
        oFile.CreateTextFile m_sFileName
    End If
vbwProfiler.vbwExecuteLine 3534 'B
vbwProfiler.vbwExecuteLine 3535
    Set oFile = Nothing
vbwProfiler.vbwProcOut 147
vbwProfiler.vbwExecuteLine 3536
End Property

Public Sub DeleteSection(ByVal sSectionName As String)
    '//this accepts a section name and deletes it and all keys in it
vbwProfiler.vbwProcIn 148
    Dim iFree As Long
    Dim sArray() As String
    Dim s As String
    Dim i As Long
    Dim j As Long

vbwProfiler.vbwExecuteLine 3537
    iFree = FreeFile
vbwProfiler.vbwExecuteLine 3538
    ReDim sArray(0)
vbwProfiler.vbwExecuteLine 3539
    i = 0

vbwProfiler.vbwExecuteLine 3540
    Open m_sFileName For Input As #iFree

'vbwLine 3541:    Do While Not EOF(iFree)
    Do While vbwProfiler.vbwExecuteLine(3541) Or Not EOF(iFree)
vbwProfiler.vbwExecuteLine 3542
        Line Input #iFree, s
vbwProfiler.vbwExecuteLine 3543
        If s <> "" Then
vbwProfiler.vbwExecuteLine 3544
            ReDim Preserve sArray(i)
vbwProfiler.vbwExecuteLine 3545
            sArray(i) = s
vbwProfiler.vbwExecuteLine 3546
            i = i + 1
        End If
vbwProfiler.vbwExecuteLine 3547 'B
vbwProfiler.vbwExecuteLine 3548
    Loop
vbwProfiler.vbwExecuteLine 3549
    Close #iFree

    '//now find the section in our array and delete it
vbwProfiler.vbwExecuteLine 3550
    For i = 0 To UBound(sArray)
vbwProfiler.vbwExecuteLine 3551
        If sArray(i) = "[" & sSectionName & "]" Then
vbwProfiler.vbwExecuteLine 3552
            sArray(i) = ""
vbwProfiler.vbwExecuteLine 3553
            For j = i + 1 To UBound(sArray)
vbwProfiler.vbwExecuteLine 3554
                If Left(sArray(j), 1) <> "[" Then
vbwProfiler.vbwExecuteLine 3555
                    sArray(j) = ""
                Else
vbwProfiler.vbwExecuteLine 3556 'B
vbwProfiler.vbwExecuteLine 3557
                    Exit For
                End If
vbwProfiler.vbwExecuteLine 3558 'B
vbwProfiler.vbwExecuteLine 3559
            Next
        End If
vbwProfiler.vbwExecuteLine 3560 'B
vbwProfiler.vbwExecuteLine 3561
    Next

    '//now write a new file
vbwProfiler.vbwExecuteLine 3562
    iFree = FreeFile

vbwProfiler.vbwExecuteLine 3563
    Open m_sFileName For Output As #iFree

vbwProfiler.vbwExecuteLine 3564
    For i = 0 To UBound(sArray)
vbwProfiler.vbwExecuteLine 3565
        If sArray(i) <> "" Then
vbwProfiler.vbwExecuteLine 3566
            Print #iFree, sArray(i)
        End If
vbwProfiler.vbwExecuteLine 3567 'B
vbwProfiler.vbwExecuteLine 3568
    Next
vbwProfiler.vbwExecuteLine 3569
    Close #iFree
vbwProfiler.vbwProcOut 148
vbwProfiler.vbwExecuteLine 3570
End Sub

Public Sub WriteSection(ByVal sSectionName As String, ParamArray vList())
    '//this writes an entire section of data
    '//the user provides the section name and the list of keys
    '//NOTE: This does not write in any Key Values...
vbwProfiler.vbwProcIn 149

    Dim sArray() As String
    Dim s As String
    Dim i As Long

vbwProfiler.vbwExecuteLine 3571
    For i = 0 To UBound(vList)
vbwProfiler.vbwExecuteLine 3572
        sArray = Split(vList(i), "=")
vbwProfiler.vbwExecuteLine 3573
        s = s & sArray(0) & Chr(10) & sArray(1) & Chr(10) & Chr(10)
vbwProfiler.vbwExecuteLine 3574
    Next

vbwProfiler.vbwExecuteLine 3575
    Call WritePrivateProfileSection(sSectionName, s, m_sFileName)

vbwProfiler.vbwProcOut 149
vbwProfiler.vbwExecuteLine 3576
End Sub

Public Function GetNextIndexedSection(ByVal sSectionName As String) As String
    '//this function creates indexed section names.  It finds the first available section
    '//and creates it and returns the name of the section it created
vbwProfiler.vbwProcIn 150
    Dim i, j, k, l As Long
    Dim iFree As Long
    Dim sName As String
    Dim s As String
    Dim sArray() As String

vbwProfiler.vbwExecuteLine 3577
    iFree = FreeFile
vbwProfiler.vbwExecuteLine 3578
    ReDim sArray(0)
vbwProfiler.vbwExecuteLine 3579
    i = 0
vbwProfiler.vbwExecuteLine 3580
    j = 0
vbwProfiler.vbwExecuteLine 3581
    k = 0
vbwProfiler.vbwExecuteLine 3582
    l = 0

vbwProfiler.vbwExecuteLine 3583
    Open m_sFileName For Input As #iFree
    '//get a list of all sections within this file
'vbwLine 3584:    Do While Not EOF(iFree)
    Do While vbwProfiler.vbwExecuteLine(3584) Or Not EOF(iFree)
vbwProfiler.vbwExecuteLine 3585
        Line Input #iFree, s
vbwProfiler.vbwExecuteLine 3586
        If Left(s, 1) = "[" Then
vbwProfiler.vbwExecuteLine 3587
            ReDim Preserve sArray(i)
vbwProfiler.vbwExecuteLine 3588
            sArray(i) = Mid(s, 2, Len(s) - 2)
vbwProfiler.vbwExecuteLine 3589
            i = i + 1
        End If
vbwProfiler.vbwExecuteLine 3590 'B
vbwProfiler.vbwExecuteLine 3591
    Loop
vbwProfiler.vbwExecuteLine 3592
    Close #iFree

    '//now find a free section
vbwProfiler.vbwExecuteLine 3593
    sName = sSectionName & k
vbwProfiler.vbwExecuteLine 3594
    For i = 0 To UBound(sArray)
vbwProfiler.vbwExecuteLine 3595
        For j = 0 To UBound(sArray)
vbwProfiler.vbwExecuteLine 3596
            If sName = sArray(j) Then
vbwProfiler.vbwExecuteLine 3597
                l = l + 1
vbwProfiler.vbwExecuteLine 3598
                Exit For
            End If
vbwProfiler.vbwExecuteLine 3599 'B
vbwProfiler.vbwExecuteLine 3600
        Next
vbwProfiler.vbwExecuteLine 3601
        If l = k Then
vbwProfiler.vbwExecuteLine 3602
             Exit For
        End If
vbwProfiler.vbwExecuteLine 3603 'B
vbwProfiler.vbwExecuteLine 3604
        k = k + 1
vbwProfiler.vbwExecuteLine 3605
        sName = sSectionName & k
vbwProfiler.vbwExecuteLine 3606
    Next
    '//create the current section
vbwProfiler.vbwExecuteLine 3607
    GetNextIndexedSection = sName
vbwProfiler.vbwProcOut 150
vbwProfiler.vbwExecuteLine 3608
End Function

Public Sub WriteString(ByVal sSectionName As String, ByVal sKeyName As String, ByVal s As String)
    '//accepts a string and writes it to the ini file
vbwProfiler.vbwProcIn 151
vbwProfiler.vbwExecuteLine 3609
     Call WritePrivateProfileString(sSectionName, sKeyName, s, m_sFileName)
vbwProfiler.vbwProcOut 151
vbwProfiler.vbwExecuteLine 3610
End Sub

Public Sub WriteInteger(ByVal sSectionName As String, ByVal sKeyName As String, ByVal l As Long)
     '//accepts a string and writes it to the ini file
vbwProfiler.vbwProcIn 152
     Dim s As String

vbwProfiler.vbwExecuteLine 3611
     s = Str(l)
vbwProfiler.vbwExecuteLine 3612
     Call WritePrivateProfileString(sSectionName, sKeyName, s, m_sFileName)
vbwProfiler.vbwProcOut 152
vbwProfiler.vbwExecuteLine 3613
End Sub

Public Sub WriteBoolean(ByVal sSectionName As String, ByVal sKeyName As String, ByVal b As Boolean)
    '//this sub accepts a boolean and writes it as a string "TRUE" or "FALSE" to the
    '//INI file
vbwProfiler.vbwProcIn 153
    Dim s As String
vbwProfiler.vbwExecuteLine 3614
    If b = True Then
vbwProfiler.vbwExecuteLine 3615
        s = "TRUE"
    Else
vbwProfiler.vbwExecuteLine 3616 'B
vbwProfiler.vbwExecuteLine 3617
        s = "FALSE"
    End If
vbwProfiler.vbwExecuteLine 3618 'B
vbwProfiler.vbwExecuteLine 3619
    Call WritePrivateProfileString(sSectionName, sKeyName, s, m_sFileName)
vbwProfiler.vbwProcOut 153
vbwProfiler.vbwExecuteLine 3620
End Sub

Public Function ReadString(ByVal sSectionName As String, sKeyName As String) As String
    '//accepts the Section and Keyname and returns the value of the key
vbwProfiler.vbwProcIn 154
    Dim sReturnString As String
vbwProfiler.vbwExecuteLine 3621
    sReturnString = Space(255)
vbwProfiler.vbwExecuteLine 3622
    ReadString = Left(sReturnString, GetPrivateProfileString(sSectionName, sKeyName, " ", sReturnString, 255, m_sFileName))
vbwProfiler.vbwProcOut 154
vbwProfiler.vbwExecuteLine 3623
End Function

Public Function ReadInteger(ByVal sSectionName As String, sKeyName As String) As Long
    '//accepts the Section and Keyname and returns the integer value of the key
vbwProfiler.vbwProcIn 155
vbwProfiler.vbwExecuteLine 3624
    ReadInteger = GetPrivateProfileInt(sSectionName, sKeyName, 0, m_sFileName)
vbwProfiler.vbwProcOut 155
vbwProfiler.vbwExecuteLine 3625
End Function

Public Function ReadSection(ByVal sSectionName As String) As String()
    '//this function accepts the name of a section and then retreives
    '//all of the keynames
vbwProfiler.vbwProcIn 156

    Dim i, j As Long
    Dim s As String
    Dim sRet() As String
    Dim lngNullOffset As Long
    Dim lngKeyValueOffset As Long

vbwProfiler.vbwExecuteLine 3626
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 3627
    s = String(MAX_SECTION_SIZE, 0) '//allocate space in our string
vbwProfiler.vbwExecuteLine 3628
    j = 0

vbwProfiler.vbwExecuteLine 3629
    i = GetPrivateProfileSection(sSectionName, s, MAX_SECTION_SIZE, m_sFileName)

    '//load the sections data into the array
vbwProfiler.vbwExecuteLine 3630
    Do
vbwProfiler.vbwExecuteLine 3631
        lngNullOffset = InStr(s, Chr(0))
vbwProfiler.vbwExecuteLine 3632
        If lngNullOffset > 1 Then
vbwProfiler.vbwExecuteLine 3633
            ReDim Preserve sRet(j)

            ' only save the keyname, not the value
vbwProfiler.vbwExecuteLine 3634
            lngKeyValueOffset = InStr(s, "=")
vbwProfiler.vbwExecuteLine 3635
            sRet(j) = Mid(s, 1, lngKeyValueOffset - 1)
            'sRet(j) = Mid(s, 1, lngNullOffset - 1)
vbwProfiler.vbwExecuteLine 3636
            s = Mid(s, lngNullOffset + 1)
vbwProfiler.vbwExecuteLine 3637
            j = j + 1
        End If
vbwProfiler.vbwExecuteLine 3638 'B
vbwProfiler.vbwExecuteLine 3639
    Loop While lngNullOffset > 1

vbwProfiler.vbwExecuteLine 3640
    ReadSection = sRet
vbwProfiler.vbwProcOut 156
vbwProfiler.vbwExecuteLine 3641
    Exit Function
errorhandler:
vbwProfiler.vbwProcOut 156
vbwProfiler.vbwExecuteLine 3642
    Exit Function
vbwProfiler.vbwProcOut 156
vbwProfiler.vbwExecuteLine 3643
End Function

Public Function RetreiveSectionNames() As String()
    '//this function opens the INI and returns the list of
    '//all the section names in an string array
vbwProfiler.vbwProcIn 157
    Dim iFileNum As Long
vbwProfiler.vbwExecuteLine 3644
    iFileNum = FreeFile
    Dim s As String
    Dim sRet() As String
    Dim i As Long

vbwProfiler.vbwExecuteLine 3645
    Open m_sFileName For Input As #iFileNum

'vbwLine 3646:    Do While Not EOF(iFileNum)
    Do While vbwProfiler.vbwExecuteLine(3646) Or Not EOF(iFileNum)
vbwProfiler.vbwExecuteLine 3647
        Line Input #iFileNum, s
vbwProfiler.vbwExecuteLine 3648
        If Left(s, 1) = "[" Then
vbwProfiler.vbwExecuteLine 3649
            ReDim Preserve sRet(i)
vbwProfiler.vbwExecuteLine 3650
            s = Mid(s, 2, Len(s) - 2)
vbwProfiler.vbwExecuteLine 3651
            sRet(i) = s
vbwProfiler.vbwExecuteLine 3652
            i = i + 1
        End If
vbwProfiler.vbwExecuteLine 3653 'B
vbwProfiler.vbwExecuteLine 3654
    Loop
vbwProfiler.vbwExecuteLine 3655
    Close iFileNum
vbwProfiler.vbwExecuteLine 3656
    RetreiveSectionNames = sRet
vbwProfiler.vbwProcOut 157
vbwProfiler.vbwExecuteLine 3657
End Function

