VERSION 5.00
Begin VB.Form frmConfigure 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configure"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAssociateExt 
      Caption         =   "Associate .Veh extensions with GVD"
      Height          =   270
      Left            =   4620
      TabIndex        =   14
      Top             =   2190
      Width           =   3210
   End
   Begin VB.Frame Frame4 
      Caption         =   "Text Formating"
      Height          =   675
      Left            =   4575
      TabIndex        =   11
      Top             =   1230
      Width           =   3345
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmConfigure.frx":0000
         Left            =   2370
         List            =   "frmConfigure.frx":0013
         TabIndex        =   12
         Text            =   "2"
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Number of Decimal Places"
         Height          =   225
         Left            =   330
         TabIndex        =   13
         Top             =   360
         Width           =   2265
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Default Publish Email Address"
      Height          =   765
      Left            =   4575
      TabIndex        =   10
      Top             =   270
      Width           =   3330
      Begin VB.TextBox txtPublishEmailAddress 
         Height          =   285
         Left            =   210
         TabIndex        =   4
         Top             =   360
         Width           =   2970
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Default HTML Viewer"
      Height          =   1635
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   4305
      Begin VB.CommandButton cmdDefaultBrowserPath 
         Caption         =   "Command3"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   1170
         Width           =   3795
      End
      Begin VB.CheckBox chkUseDefaultWebBrowser 
         Caption         =   "Use associated viewer for .HTM/HTML extensions"
         Height          =   225
         Left            =   240
         TabIndex        =   0
         Top             =   390
         Width           =   3915
      End
      Begin VB.Label Label2 
         Caption         =   "Location of HTML viewer:"
         Height          =   225
         Left            =   210
         TabIndex        =   9
         Top             =   870
         Width           =   4035
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Done"
      Height          =   405
      Left            =   7140
      TabIndex        =   5
      Top             =   3690
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "Default Text Viewer"
      Height          =   1605
      Left            =   120
      TabIndex        =   6
      Top             =   2100
      Width           =   4320
      Begin VB.CheckBox chkUseDefaultTextViewer 
         Caption         =   "Use associated viewer for .TXT extensions"
         Height          =   225
         Left            =   300
         TabIndex        =   2
         Top             =   390
         Width           =   3555
      End
      Begin VB.CommandButton cmdDefaultTextViewerPath 
         Caption         =   "Command1"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   3585
      End
      Begin VB.Label Label1 
         Caption         =   "Location of Vehicle text file viewer:"
         Height          =   225
         Left            =   270
         TabIndex        =   7
         Top             =   810
         Width           =   2925
      End
   End
End
Attribute VB_Name = "frmConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkAssociateExt_Click()
    On Error Resume Next
    Dim oAssociater As clsAssociateExt
    
    Set oAssociater = New clsAssociateExt
    
    If chkAssociateExt.value <> 0 Then 'true
    
        'set the association
        With oAssociater
            .Extension = ".veh"
            .DefaultIcon = "shell32.dll,72"
            .Description = "GVD vehicle file"
            .OpenCommand = """" & App.Path & "\GVD.exe""" & " %1"
            Debug.Print "chkAssociateExt_Click: " & .OpenCommand
            '.PrintCommand = Text1(4).Text
            .SetAssociation
        End With
    Else
       
        ' if its set, remove it
        oAssociater.DeleteGVDAssociation
        
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''
'02/16/02 MPJ  No longer using Animation or SOund in opening splash sequence so these two functions are obsolete
'Private Sub chkQuickStart_Click()
'    If chkQuickStart.value = 0 Then 'true
'
'        Settings.bQuickStart = 0
'    Else
'        Settings.bQuickStart = 1
'    End If
'End Sub
'
'Private Sub chkSound_Click()
'
'    If chkSound.value = 0 Then  'TRUE then
'        Settings.bSoundOff = 0
'    Else
'        Settings.bSoundOff = 1
'    End If
'
'End Sub
''''''''''''''''''''''''''''''''''''''''''''''


Private Sub chkUseDefaultTextViewer_Click()
If chkUseDefaultTextViewer.value = 1 Then
    Settings.bUseDefaultTextViewer = 1
    cmdDefaultTextViewerPath.Enabled = False
    cmdDefaultTextViewerPath.TAG = FindDefaultProgram("Text")
    cmdDefaultTextViewerPath.Caption = Abbreviated(cmdDefaultTextViewerPath.TAG)

    Settings.TextViewerPath = cmdDefaultTextViewerPath.TAG
Else
    Settings.bUseDefaultTextViewer = 0
    cmdDefaultTextViewerPath.Enabled = True
End If
End Sub

Private Sub chkUseDefaultWebBrowser_Click()
If chkUseDefaultWebBrowser.value = 1 Then
    Settings.bUseDefaultWebBrowser = 1
    cmdDefaultBrowserPath.Enabled = False
    cmdDefaultBrowserPath.TAG = FindDefaultProgram("Browser")
    cmdDefaultBrowserPath.Caption = Abbreviated(cmdDefaultBrowserPath.TAG)
    Settings.HTMLBrowserPath = cmdDefaultBrowserPath.TAG
Else
    Settings.bUseDefaultWebBrowser = 0
    cmdDefaultBrowserPath.Enabled = True
End If
End Sub



Private Sub cmdDefaultBrowserPath_Click()
Dim sPath As String

sPath = GetPath("HTMLBrowser")

If sPath <> "" Then
    cmdDefaultBrowserPath.TAG = sPath
    cmdDefaultBrowserPath.Caption = Abbreviated(sPath)
    Settings.HTMLBrowserPath = sPath
End If
End Sub

Private Sub cmdDefaultTextViewerPath_Click()
Dim sPath As String

sPath = GetPath("TextViewer")

If sPath <> "" Then
    cmdDefaultTextViewerPath.TAG = sPath
    cmdDefaultTextViewerPath.Caption = Abbreviated(sPath)
    Settings.TextViewerPath = sPath
End If
End Sub

Private Sub Combo1_click()
    Settings.DecimalPlaces = Val(Combo1)
    Select Case Settings.DecimalPlaces
        Case 0
            Settings.FormatString = "#,##0"
        Case 1
            Settings.FormatString = "#,##0.0"
        Case 2
            Settings.FormatString = "standard"
        Case 3
            Settings.FormatString = "#,##0.000"
        Case 4
            Settings.FormatString = "#,##0.0000"
    End Select
    m_oCurrentVeh.FormatString = Settings.FormatString
    
    
End Sub

Private Sub Command2_Click()

On Error GoTo errorhandler
Unload Me

Exit Sub
errorhandler:
    DoEvents
    DoEvents
    Resume
End Sub


Private Function GetPath(sType As String) As String

'code for user to set the path of the exe
Dim Cancel As Boolean ' detects whether the user clicks cancel at the Open dialog
Dim sPath As String
Dim oCDLG As clsCmdlg

'lets change directories to the top of C
ChDir App.Path
Cancel = False  ' initialize the cancel button variable for Common dialog

With oCDLG
    .InitialDir = App.Path
    .Filter = "Executeable (*.EXE)|*.EXE|All files (*.*)|*.*"
    .CancelError = True
    .DefaultFilename = ""
    '.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    .MultiSelect = False
End With

Cancel = oCDLG.ShowOpen

If Not Cancel Then
    sPath = oCDLG.cFileName(0)
    GetPath = sPath
Else
    Cancel = True
    GetPath = ""
End If
End Function

Function Abbreviated(sPath As String) As String
If Len(sPath) < 24 Then
    Abbreviated = sPath
Else
    Abbreviated = Left(sPath, 24) & "\...\"
End If
End Function


Private Sub Form_Activate()
    On Error Resume Next
'//load in our settings
    txtPublishEmailAddress = Settings.PublishEmailAddress

    Combo1.Text = Val(Settings.DecimalPlaces)
   
   
    'chkQuickStart.value = Abs(Settings.bQuickStart) 'Disabled 02/16/02 MPJ (obsolete)
    'chkSound.value = Abs(Settings.bSoundOff)        'Disabled 02/16/02 MPJ (obsolete)
    chkAssociateExt.value = Abs(Settings.bAssociateExt)
    
    chkUseDefaultWebBrowser.value = Abs(Settings.bUseDefaultWebBrowser)
    If chkUseDefaultWebBrowser.value = 1 Then
        With cmdDefaultBrowserPath
            .Enabled = False
            .TAG = FindDefaultProgram("Browser")
            .Caption = Abbreviated(.TAG)
            Settings.HTMLBrowserPath = .TAG
        End With
    Else
        Settings.bUseDefaultWebBrowser = 0
        With cmdDefaultBrowserPath
            .Enabled = True
            .TAG = Settings.HTMLBrowserPath
            .Caption = Abbreviated(.TAG)
        End With
    End If
    
    chkUseDefaultTextViewer.value = Abs(Settings.bUseDefaultTextViewer)
    If chkUseDefaultTextViewer.value = 1 Then
        With cmdDefaultTextViewerPath
            .Enabled = False
            .TAG = FindDefaultProgram("Text")
            .Caption = Abbreviated(.TAG)
            Settings.TextViewerPath = .TAG
        End With
    Else
        Settings.bUseDefaultTextViewer = 0
        With cmdDefaultTextViewerPath
            .Enabled = True
            .TAG = Settings.TextViewerPath
            .Caption = Abbreviated(.TAG)
        End With
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
'the max length of the labels should be around 40 characters?

End Sub

Private Sub Form_Unload(Cancel As Integer)
'//need to save the settings
With Settings
    .bUseDefaultTextViewer = Abs(chkUseDefaultTextViewer)
    .TextViewerPath = cmdDefaultTextViewerPath.TAG
    .bUseDefaultWebBrowser = Abs(chkUseDefaultWebBrowser)
    .HTMLBrowserPath = cmdDefaultBrowserPath.TAG
    .PublishEmailAddress = txtPublishEmailAddress
    '.bQuickStart = Abs(chkQuickStart)  'Disabled 02/16/02 MPJ (obsolete)
    '.bSoundOff = Abs(chkSound)         'Disabled 02/16/02 MPJ (obsolete)
    .bAssociateExt = Abs(chkAssociateExt)
End With

End Sub

Private Sub txtPublishEmailAddress_KeyPress(KeyAscii As Integer)
    If Not IsValidKeyCode(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub
