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
    ReDim m_oMenu(0)
    
    Set m_oMenu(0) = v
    
    'test
   
End Property

Function AddRecentFile(sFullPath As String) As Long

    '//add this file to our number 1 spot.  If it is already at the number 1 spot
    '//move it down a notch and same with all below it except for 5th which gets bumped
    '//if the spot we are moving down to is the same as any of the ones above it, delete it
    Dim i, j As Long
    Dim sCurrentFile As String
    Dim sCurrentPath As String
    Dim sPrevFile As String
    Dim sPrevPath As String
    
    If Not m_oMenu(0) Is Nothing Then
    
        sCurrentFile = ExtractFileNameFromPath(sFullPath)
        sCurrentPath = ExtractPathFromFile(sFullPath)
        
        On Error Resume Next
        
        
        
        '//if its already on the recent file list, we wont add it again
        If Not AlreadyInList(sFullPath) Then
            For i = 5 To 2 Step -1
                m_oMenu(i).Caption = m_oMenu(i - 1).Caption
                Debug.Print "AddRecentFile: " & "Caption " & i & " = " & m_oMenu(i).Caption
                Debug.Print "AddRecentFile: " & "Caption " & i - 1 & " = " & m_oMenu(i - 1).Caption
                m_oMenu(i).Visible = True
                'mnuRecent(i).Caption = "&" & i & " " & Left(mnuRecent(i - 1).Caption, Len(mnuRecent(i - 1).Caption) - 3)
            Next
            m_oMenu(1).Caption = sFullPath
            'mnuRecent(1).Tag = sFilePath
            m_oMenu(1).Visible = True
            
            '//hide all the ones with "" for captions
            For i = 1 To 5
                If m_oMenu(i).Caption = "" Then m_oMenu(i).Visible = False
            Next
            
            '//make sure our seperator is visible
            If m_oMenu(1).Caption <> "" Then m_oMenu(0).Visible = True
        End If
    End If
End Function

Sub DeleteRecentFile(sFilePath As String)
    Dim i As Long
    
    For i = 1 To 5
        If m_oMenu(i).Caption = sFilePath Then
            m_oMenu(i).Caption = ""
            m_oMenu(i).Visible = False
        End If
    Next
    
End Sub
Private Function AlreadyInList(sFilePath) As Boolean
    Dim i As Long
    
    For i = 5 To 1 Step -1
        If m_oMenu(i).Caption = sFilePath Then
            AlreadyInList = True
            Exit Function
        End If
    Next
    
    AlreadyInList = False
End Function

Private Sub Class_Terminate()
    Set m_oMenu(0) = Nothing
End Sub
