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
vbwProfiler.vbwProcIn 237
vbwProfiler.vbwExecuteLine 4585
    If x > y Then
vbwProfiler.vbwExecuteLine 4586
         Maximum = x
    Else
vbwProfiler.vbwExecuteLine 4587 'B
vbwProfiler.vbwExecuteLine 4588
         Maximum = y
    End If
vbwProfiler.vbwExecuteLine 4589 'B
vbwProfiler.vbwProcOut 237
vbwProfiler.vbwExecuteLine 4590
End Function

Function Minimum(ByVal x, ByVal y)
vbwProfiler.vbwProcIn 238
vbwProfiler.vbwExecuteLine 4591
    If x < y Then
vbwProfiler.vbwExecuteLine 4592
         Minimum = x
    Else
vbwProfiler.vbwExecuteLine 4593 'B
vbwProfiler.vbwExecuteLine 4594
         Minimum = y
    End If
vbwProfiler.vbwExecuteLine 4595 'B
vbwProfiler.vbwProcOut 238
vbwProfiler.vbwExecuteLine 4596
End Function

Sub InfoPrint(ByVal Code As Integer, ByVal Message As String)
vbwProfiler.vbwProcIn 239
vbwProfiler.vbwExecuteLine 4597
   frmDesigner.txtInfo.Text = Message & vbNewLine & Left(frmDesigner.txtInfo.Text, 2000)
vbwProfiler.vbwProcOut 239
vbwProfiler.vbwExecuteLine 4598
End Sub

Function ExtractPathFromFile(s As String) As String
vbwProfiler.vbwProcIn 240

    Dim i As Long
    Dim sRet As String

vbwProfiler.vbwExecuteLine 4599
    For i = Len(s) To 1 Step -1
vbwProfiler.vbwExecuteLine 4600
        If Mid(s, i, 1) = "\" Then
vbwProfiler.vbwExecuteLine 4601
            sRet = Left(s, i)
vbwProfiler.vbwExecuteLine 4602
            Exit For
        End If
vbwProfiler.vbwExecuteLine 4603 'B
vbwProfiler.vbwExecuteLine 4604
    Next

vbwProfiler.vbwExecuteLine 4605
    ExtractPathFromFile = sRet

vbwProfiler.vbwProcOut 240
vbwProfiler.vbwExecuteLine 4606
End Function

Function ExtractFileNameFromPath(s As String) As String
vbwProfiler.vbwProcIn 241
    Dim i As Long
    Dim j As Long

    ' get the actual filename from the filepath
vbwProfiler.vbwExecuteLine 4607
        For i = Len(s) To 1 Step -1
vbwProfiler.vbwExecuteLine 4608
            j = j + 1
vbwProfiler.vbwExecuteLine 4609
            If Mid(s, i, 1) = "\" Then
vbwProfiler.vbwExecuteLine 4610
                ExtractFileNameFromPath = Right(s, j - 1)
vbwProfiler.vbwExecuteLine 4611
                Exit For
            End If
vbwProfiler.vbwExecuteLine 4612 'B
vbwProfiler.vbwExecuteLine 4613
        Next

vbwProfiler.vbwProcOut 241
vbwProfiler.vbwExecuteLine 4614
End Function


Function StartDoc(ByRef DocName As String) As Long
'this function launches the text file using the associated viewer
vbwProfiler.vbwProcIn 242
    Dim Scr_hDC As Long

vbwProfiler.vbwExecuteLine 4615
    Scr_hDC = GetDesktopWindow()
vbwProfiler.vbwExecuteLine 4616
    StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", SW_SHOWNORMAL)
vbwProfiler.vbwProcOut 242
vbwProfiler.vbwExecuteLine 4617
End Function

Function FindDefaultProgram(ByVal sType As String) As String
vbwProfiler.vbwProcIn 243

    Dim FileName, Dummy As String
    Dim BrowserExec As String * 255
    Dim retval As Long
    Dim FileNumber As Integer

      ' First, create a known, temporary file

vbwProfiler.vbwExecuteLine 4618
      BrowserExec = Space(255)

vbwProfiler.vbwExecuteLine 4619
      If sType = "Text" Then
vbwProfiler.vbwExecuteLine 4620
          FileName = App.Path + "\tmp00000001.txt"
      Else
vbwProfiler.vbwExecuteLine 4621 'B
vbwProfiler.vbwExecuteLine 4622
          FileName = App.Path + "\tmp00000001.HTM"
      End If
vbwProfiler.vbwExecuteLine 4623 'B

vbwProfiler.vbwExecuteLine 4624
      FileNumber = FreeFile                    ' Get unused file number
vbwProfiler.vbwExecuteLine 4625
      Open FileName For Output As #FileNumber  ' Create temp HTML file
vbwProfiler.vbwExecuteLine 4626
          Write #FileNumber, "<HTML> <\HTML>"  ' Output text
vbwProfiler.vbwExecuteLine 4627
      Close #FileNumber                        ' Close file

      ' Then find the application associated with it
vbwProfiler.vbwExecuteLine 4628
      retval = FindExecutable(FileName, Dummy, BrowserExec)
vbwProfiler.vbwExecuteLine 4629
      BrowserExec = Trim(BrowserExec)

      'delete the temp file
vbwProfiler.vbwExecuteLine 4630
      If sType = "Text" Then
vbwProfiler.vbwExecuteLine 4631
          Kill App.Path + "\tmp00000001.txt"
      Else
vbwProfiler.vbwExecuteLine 4632 'B
vbwProfiler.vbwExecuteLine 4633
          Kill App.Path + "\tmp00000001.HTM"
      End If
vbwProfiler.vbwExecuteLine 4634 'B

vbwProfiler.vbwExecuteLine 4635
FindDefaultProgram = BrowserExec
vbwProfiler.vbwProcOut 243
vbwProfiler.vbwExecuteLine 4636
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



