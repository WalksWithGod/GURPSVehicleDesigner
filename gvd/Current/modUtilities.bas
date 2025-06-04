Attribute VB_Name = "modUtilities"
Option Explicit

'//////the below is for launching the default text viewer and default web browsers
Private Declare Function FindExecutable Lib "shell32.dll" Alias _
         "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
         String, ByVal lpResult As String) As Long
         
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
      "ShellExecuteA" (ByVal hWnd As Long, ByVal lpszOp As _
      String, ByVal lpszFile As String, ByVal lpszParams As String, _
      ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
      
Private Declare Function GetDesktopWindow Lib "user32" () As Long ' used for launching associated viewers
      
Function Maximum(ByVal x, ByVal y)
    If x > y Then Maximum = x Else Maximum = y
End Function

Function Minimum(ByVal x, ByVal y)
    If x < y Then Minimum = x Else Minimum = y
End Function

Sub InfoPrint(ByVal Code As Integer, ByVal Message As String)
   frmDesigner.txtInfo.Text = Message & vbNewLine & Left(frmDesigner.txtInfo.Text, 2000)
End Sub

Function ExtractPathFromFile(s As String) As String
    
    Dim i As Long
    Dim sRet As String
    
    For i = Len(s) To 1 Step -1
        If Mid(s, i, 1) = "\" Then
            sRet = Left(s, i)
            Exit For
        End If
    Next
    
    ExtractPathFromFile = sRet
    
End Function

Function ExtractFileNameFromPath(s As String) As String
    Dim i As Long
    Dim j As Long
    
    ' get the actual filename from the filepath
        For i = Len(s) To 1 Step -1
            j = j + 1
            If Mid(s, i, 1) = "\" Then
                ExtractFileNameFromPath = Right(s, j - 1)
                Exit For
            End If
        Next
        
End Function


Function StartDoc(ByRef DocName As String) As Long
'this function launches the text file using the associated viewer
    Dim Scr_hDC As Long
          
    Scr_hDC = GetDesktopWindow()
    StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", SW_SHOWNORMAL)
End Function

Function FindDefaultProgram(ByVal sType As String) As String

    Dim FileName, Dummy As String
    Dim BrowserExec As String * 255
    Dim retval As Long
    Dim FileNumber As Integer
        
      ' First, create a known, temporary file
      
      BrowserExec = Space(255)
      
      If sType = "Text" Then
          FileName = App.Path + "\tmp00000001.txt"
      Else
          FileName = App.Path + "\tmp00000001.HTM"
      End If
      
      FileNumber = FreeFile                    ' Get unused file number
      Open FileName For Output As #FileNumber  ' Create temp HTML file
          Write #FileNumber, "<HTML> <\HTML>"  ' Output text
      Close #FileNumber                        ' Close file
      
      ' Then find the application associated with it
      retval = FindExecutable(FileName, Dummy, BrowserExec)
      BrowserExec = Trim(BrowserExec)
      
      'delete the temp file
      If sType = "Text" Then
          Kill App.Path + "\tmp00000001.txt"
      Else
          Kill App.Path + "\tmp00000001.HTM"
      End If
      
FindDefaultProgram = BrowserExec
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


