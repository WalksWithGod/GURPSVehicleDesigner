VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecentFileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_oMenu() As Menu

Public Property Set Menu(v As Menu)
vbwProfiler.vbwProcIn 435
vbwProfiler.vbwExecuteLine 9763
    ReDim m_oMenu(0)

vbwProfiler.vbwExecuteLine 9764
    Set m_oMenu(0) = v

    'test

vbwProfiler.vbwProcOut 435
vbwProfiler.vbwExecuteLine 9765
End Property

Function AddRecentFile(sFullPath As String) As Long
vbwProfiler.vbwProcIn 436

    '//add this file to our number 1 spot.  If it is already at the number 1 spot
    '//move it down a notch and same with all below it except for 5th which gets bumped
    '//if the spot we are moving down to is the same as any of the ones above it, delete it
    Dim i, j As Long
    Dim sCurrentFile As String
    Dim sCurrentPath As String
    Dim sPrevFile As String
    Dim sPrevPath As String

vbwProfiler.vbwExecuteLine 9766
    If Not m_oMenu(0) Is Nothing Then

vbwProfiler.vbwExecuteLine 9767
        sCurrentFile = ExtractFileNameFromPath(sFullPath)
vbwProfiler.vbwExecuteLine 9768
        sCurrentPath = ExtractPathFromFile(sFullPath)

vbwProfiler.vbwExecuteLine 9769
        On Error Resume Next



        '//if its already on the recent file list, we wont add it again
vbwProfiler.vbwExecuteLine 9770
        If Not AlreadyInList(sFullPath) Then
vbwProfiler.vbwExecuteLine 9771
            For i = 5 To 2 Step -1
vbwProfiler.vbwExecuteLine 9772
                m_oMenu(i).Caption = m_oMenu(i - 1).Caption
vbwProfiler.vbwExecuteLine 9773
                Debug.Print "AddRecentFile: " & "Caption " & i & " = " & m_oMenu(i).Caption
vbwProfiler.vbwExecuteLine 9774
                Debug.Print "AddRecentFile: " & "Caption " & i - 1 & " = " & m_oMenu(i - 1).Caption
vbwProfiler.vbwExecuteLine 9775
                m_oMenu(i).Visible = True
                'mnuRecent(i).Caption = "&" & i & " " & Left(mnuRecent(i - 1).Caption, Len(mnuRecent(i - 1).Caption) - 3)
vbwProfiler.vbwExecuteLine 9776
            Next
vbwProfiler.vbwExecuteLine 9777
            m_oMenu(1).Caption = sFullPath
            'mnuRecent(1).Tag = sFilePath
vbwProfiler.vbwExecuteLine 9778
            m_oMenu(1).Visible = True

            '//hide all the ones with "" for captions
vbwProfiler.vbwExecuteLine 9779
            For i = 1 To 5
vbwProfiler.vbwExecuteLine 9780
                If m_oMenu(i).Caption = "" Then
vbwProfiler.vbwExecuteLine 9781
                     m_oMenu(i).Visible = False
                End If
vbwProfiler.vbwExecuteLine 9782 'B
vbwProfiler.vbwExecuteLine 9783
            Next

            '//make sure our seperator is visible
vbwProfiler.vbwExecuteLine 9784
            If m_oMenu(1).Caption <> "" Then
vbwProfiler.vbwExecuteLine 9785
                 m_oMenu(0).Visible = True
            End If
vbwProfiler.vbwExecuteLine 9786 'B
        End If
vbwProfiler.vbwExecuteLine 9787 'B
    End If
vbwProfiler.vbwExecuteLine 9788 'B
vbwProfiler.vbwProcOut 436
vbwProfiler.vbwExecuteLine 9789
End Function

Sub DeleteRecentFile(sFilePath As String)
vbwProfiler.vbwProcIn 437
    Dim i As Long

vbwProfiler.vbwExecuteLine 9790
    For i = 1 To 5
vbwProfiler.vbwExecuteLine 9791
        If m_oMenu(i).Caption = sFilePath Then
vbwProfiler.vbwExecuteLine 9792
            m_oMenu(i).Caption = ""
vbwProfiler.vbwExecuteLine 9793
            m_oMenu(i).Visible = False
        End If
vbwProfiler.vbwExecuteLine 9794 'B
vbwProfiler.vbwExecuteLine 9795
    Next

vbwProfiler.vbwProcOut 437
vbwProfiler.vbwExecuteLine 9796
End Sub
Private Function AlreadyInList(sFilePath) As Boolean
vbwProfiler.vbwProcIn 438
    Dim i As Long

vbwProfiler.vbwExecuteLine 9797
    For i = 5 To 1 Step -1
vbwProfiler.vbwExecuteLine 9798
        If m_oMenu(i).Caption = sFilePath Then
vbwProfiler.vbwExecuteLine 9799
            AlreadyInList = True
vbwProfiler.vbwProcOut 438
vbwProfiler.vbwExecuteLine 9800
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 9801 'B
vbwProfiler.vbwExecuteLine 9802
    Next

vbwProfiler.vbwExecuteLine 9803
    AlreadyInList = False
vbwProfiler.vbwProcOut 438
vbwProfiler.vbwExecuteLine 9804
End Function

Private Sub Class_Terminate()
vbwProfiler.vbwProcIn 439
vbwProfiler.vbwExecuteLine 9805
    Set m_oMenu(0) = Nothing
vbwProfiler.vbwProcOut 439
vbwProfiler.vbwExecuteLine 9806
End Sub

