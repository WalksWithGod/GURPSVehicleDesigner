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
vbwProfiler.vbwProcIn 86
vbwProfiler.vbwExecuteLine 2510
    On Error Resume Next
    Dim oAssociater As clsAssociateExt

vbwProfiler.vbwExecuteLine 2511
    Set oAssociater = New clsAssociateExt

vbwProfiler.vbwExecuteLine 2512
    If chkAssociateExt.value <> 0 Then 'true

        'set the association
vbwProfiler.vbwExecuteLine 2513
        With oAssociater
vbwProfiler.vbwExecuteLine 2514
            .Extension = ".veh"
vbwProfiler.vbwExecuteLine 2515
            .DefaultIcon = "shell32.dll,72"
vbwProfiler.vbwExecuteLine 2516
            .Description = "GVD vehicle file"
vbwProfiler.vbwExecuteLine 2517
            .OpenCommand = """" & App.Path & "\GVD.exe""" & " %1"
vbwProfiler.vbwExecuteLine 2518
            Debug.Print "chkAssociateExt_Click: " & .OpenCommand
            '.PrintCommand = Text1(4).Text
vbwProfiler.vbwExecuteLine 2519
            .SetAssociation
vbwProfiler.vbwExecuteLine 2520
        End With
    Else
vbwProfiler.vbwExecuteLine 2521 'B

        ' if its set, remove it
vbwProfiler.vbwExecuteLine 2522
        oAssociater.DeleteGVDAssociation

    End If
vbwProfiler.vbwExecuteLine 2523 'B
vbwProfiler.vbwProcOut 86
vbwProfiler.vbwExecuteLine 2524
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
vbwProfiler.vbwProcIn 87
vbwProfiler.vbwExecuteLine 2525
If chkUseDefaultTextViewer.value = 1 Then
vbwProfiler.vbwExecuteLine 2526
    Settings.bUseDefaultTextViewer = 1
vbwProfiler.vbwExecuteLine 2527
    cmdDefaultTextViewerPath.Enabled = False
vbwProfiler.vbwExecuteLine 2528
    cmdDefaultTextViewerPath.TAG = FindDefaultProgram("Text")
vbwProfiler.vbwExecuteLine 2529
    cmdDefaultTextViewerPath.Caption = Abbreviated(cmdDefaultTextViewerPath.TAG)

vbwProfiler.vbwExecuteLine 2530
    Settings.TextViewerPath = cmdDefaultTextViewerPath.TAG
Else
vbwProfiler.vbwExecuteLine 2531 'B
vbwProfiler.vbwExecuteLine 2532
    Settings.bUseDefaultTextViewer = 0
vbwProfiler.vbwExecuteLine 2533
    cmdDefaultTextViewerPath.Enabled = True
End If
vbwProfiler.vbwExecuteLine 2534 'B
vbwProfiler.vbwProcOut 87
vbwProfiler.vbwExecuteLine 2535
End Sub

Private Sub chkUseDefaultWebBrowser_Click()
vbwProfiler.vbwProcIn 88
vbwProfiler.vbwExecuteLine 2536
If chkUseDefaultWebBrowser.value = 1 Then
vbwProfiler.vbwExecuteLine 2537
    Settings.bUseDefaultWebBrowser = 1
vbwProfiler.vbwExecuteLine 2538
    cmdDefaultBrowserPath.Enabled = False
vbwProfiler.vbwExecuteLine 2539
    cmdDefaultBrowserPath.TAG = FindDefaultProgram("Browser")
vbwProfiler.vbwExecuteLine 2540
    cmdDefaultBrowserPath.Caption = Abbreviated(cmdDefaultBrowserPath.TAG)
vbwProfiler.vbwExecuteLine 2541
    Settings.HTMLBrowserPath = cmdDefaultBrowserPath.TAG
Else
vbwProfiler.vbwExecuteLine 2542 'B
vbwProfiler.vbwExecuteLine 2543
    Settings.bUseDefaultWebBrowser = 0
vbwProfiler.vbwExecuteLine 2544
    cmdDefaultBrowserPath.Enabled = True
End If
vbwProfiler.vbwExecuteLine 2545 'B
vbwProfiler.vbwProcOut 88
vbwProfiler.vbwExecuteLine 2546
End Sub



Private Sub cmdDefaultBrowserPath_Click()
vbwProfiler.vbwProcIn 89
Dim sPath As String

vbwProfiler.vbwExecuteLine 2547
sPath = GetPath("HTMLBrowser")

vbwProfiler.vbwExecuteLine 2548
If sPath <> "" Then
vbwProfiler.vbwExecuteLine 2549
    cmdDefaultBrowserPath.TAG = sPath
vbwProfiler.vbwExecuteLine 2550
    cmdDefaultBrowserPath.Caption = Abbreviated(sPath)
vbwProfiler.vbwExecuteLine 2551
    Settings.HTMLBrowserPath = sPath
End If
vbwProfiler.vbwExecuteLine 2552 'B
vbwProfiler.vbwProcOut 89
vbwProfiler.vbwExecuteLine 2553
End Sub

Private Sub cmdDefaultTextViewerPath_Click()
vbwProfiler.vbwProcIn 90
Dim sPath As String

vbwProfiler.vbwExecuteLine 2554
sPath = GetPath("TextViewer")

vbwProfiler.vbwExecuteLine 2555
If sPath <> "" Then
vbwProfiler.vbwExecuteLine 2556
    cmdDefaultTextViewerPath.TAG = sPath
vbwProfiler.vbwExecuteLine 2557
    cmdDefaultTextViewerPath.Caption = Abbreviated(sPath)
vbwProfiler.vbwExecuteLine 2558
    Settings.TextViewerPath = sPath
End If
vbwProfiler.vbwExecuteLine 2559 'B
vbwProfiler.vbwProcOut 90
vbwProfiler.vbwExecuteLine 2560
End Sub

Private Sub Combo1_click()
vbwProfiler.vbwProcIn 91
vbwProfiler.vbwExecuteLine 2561
    Settings.DecimalPlaces = Val(Combo1)
vbwProfiler.vbwExecuteLine 2562
    Select Case Settings.DecimalPlaces
'vbwLine 2563:        Case 0
        Case IIf(vbwProfiler.vbwExecuteLine(2563), VBWPROFILER_EMPTY, _
        0)
vbwProfiler.vbwExecuteLine 2564
            Settings.FormatString = "#,##0"
'vbwLine 2565:        Case 1
        Case IIf(vbwProfiler.vbwExecuteLine(2565), VBWPROFILER_EMPTY, _
        1)
vbwProfiler.vbwExecuteLine 2566
            Settings.FormatString = "#,##0.0"
'vbwLine 2567:        Case 2
        Case IIf(vbwProfiler.vbwExecuteLine(2567), VBWPROFILER_EMPTY, _
        2)
vbwProfiler.vbwExecuteLine 2568
            Settings.FormatString = "standard"
'vbwLine 2569:        Case 3
        Case IIf(vbwProfiler.vbwExecuteLine(2569), VBWPROFILER_EMPTY, _
        3)
vbwProfiler.vbwExecuteLine 2570
            Settings.FormatString = "#,##0.000"
'vbwLine 2571:        Case 4
        Case IIf(vbwProfiler.vbwExecuteLine(2571), VBWPROFILER_EMPTY, _
        4)
vbwProfiler.vbwExecuteLine 2572
            Settings.FormatString = "#,##0.0000"
    End Select
vbwProfiler.vbwExecuteLine 2573 'B
vbwProfiler.vbwExecuteLine 2574
    m_oCurrentVeh.FormatString = Settings.FormatString


vbwProfiler.vbwProcOut 91
vbwProfiler.vbwExecuteLine 2575
End Sub

Private Sub Command2_Click()
vbwProfiler.vbwProcIn 92

vbwProfiler.vbwExecuteLine 2576
On Error GoTo errorhandler
vbwProfiler.vbwExecuteLine 2577
Unload Me

vbwProfiler.vbwProcOut 92
vbwProfiler.vbwExecuteLine 2578
Exit Sub
errorhandler:
vbwProfiler.vbwExecuteLine 2579
    DoEvents
vbwProfiler.vbwExecuteLine 2580
    DoEvents
vbwProfiler.vbwExecuteLine 2581
    Resume
vbwProfiler.vbwProcOut 92
vbwProfiler.vbwExecuteLine 2582
End Sub


Private Function GetPath(sType As String) As String
vbwProfiler.vbwProcIn 93

'code for user to set the path of the exe
Dim Cancel As Boolean ' detects whether the user clicks cancel at the Open dialog
Dim sPath As String
Dim oCDLG As clsCmdlg

'lets change directories to the top of C
vbwProfiler.vbwExecuteLine 2583
ChDir App.Path
vbwProfiler.vbwExecuteLine 2584
Cancel = False  ' initialize the cancel button variable for Common dialog

vbwProfiler.vbwExecuteLine 2585
With oCDLG
vbwProfiler.vbwExecuteLine 2586
    .InitialDir = App.Path
vbwProfiler.vbwExecuteLine 2587
    .Filter = "Executeable (*.EXE)|*.EXE|All files (*.*)|*.*"
vbwProfiler.vbwExecuteLine 2588
    .CancelError = True
vbwProfiler.vbwExecuteLine 2589
    .DefaultFilename = ""
    '.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
vbwProfiler.vbwExecuteLine 2590
    .MultiSelect = False
vbwProfiler.vbwExecuteLine 2591
End With

vbwProfiler.vbwExecuteLine 2592
Cancel = oCDLG.ShowOpen

vbwProfiler.vbwExecuteLine 2593
If Not Cancel Then
vbwProfiler.vbwExecuteLine 2594
    sPath = oCDLG.cFileName(0)
vbwProfiler.vbwExecuteLine 2595
    GetPath = sPath
Else
vbwProfiler.vbwExecuteLine 2596 'B
vbwProfiler.vbwExecuteLine 2597
    Cancel = True
vbwProfiler.vbwExecuteLine 2598
    GetPath = ""
End If
vbwProfiler.vbwExecuteLine 2599 'B
vbwProfiler.vbwProcOut 93
vbwProfiler.vbwExecuteLine 2600
End Function

Function Abbreviated(sPath As String) As String
vbwProfiler.vbwProcIn 94
vbwProfiler.vbwExecuteLine 2601
If Len(sPath) < 24 Then
vbwProfiler.vbwExecuteLine 2602
    Abbreviated = sPath
Else
vbwProfiler.vbwExecuteLine 2603 'B
vbwProfiler.vbwExecuteLine 2604
    Abbreviated = Left(sPath, 24) & "\...\"
End If
vbwProfiler.vbwExecuteLine 2605 'B
vbwProfiler.vbwProcOut 94
vbwProfiler.vbwExecuteLine 2606
End Function


Private Sub Form_Activate()
vbwProfiler.vbwProcIn 95
vbwProfiler.vbwExecuteLine 2607
    On Error Resume Next
'//load in our settings
vbwProfiler.vbwExecuteLine 2608
    txtPublishEmailAddress = Settings.PublishEmailAddress

vbwProfiler.vbwExecuteLine 2609
    Combo1.Text = Val(Settings.DecimalPlaces)


    'chkQuickStart.value = Abs(Settings.bQuickStart) 'Disabled 02/16/02 MPJ (obsolete)
    'chkSound.value = Abs(Settings.bSoundOff)        'Disabled 02/16/02 MPJ (obsolete)
vbwProfiler.vbwExecuteLine 2610
    chkAssociateExt.value = Abs(Settings.bAssociateExt)

vbwProfiler.vbwExecuteLine 2611
    chkUseDefaultWebBrowser.value = Abs(Settings.bUseDefaultWebBrowser)
vbwProfiler.vbwExecuteLine 2612
    If chkUseDefaultWebBrowser.value = 1 Then
vbwProfiler.vbwExecuteLine 2613
        With cmdDefaultBrowserPath
vbwProfiler.vbwExecuteLine 2614
            .Enabled = False
vbwProfiler.vbwExecuteLine 2615
            .TAG = FindDefaultProgram("Browser")
vbwProfiler.vbwExecuteLine 2616
            .Caption = Abbreviated(.TAG)
vbwProfiler.vbwExecuteLine 2617
            Settings.HTMLBrowserPath = .TAG
vbwProfiler.vbwExecuteLine 2618
        End With
    Else
vbwProfiler.vbwExecuteLine 2619 'B
vbwProfiler.vbwExecuteLine 2620
        Settings.bUseDefaultWebBrowser = 0
vbwProfiler.vbwExecuteLine 2621
        With cmdDefaultBrowserPath
vbwProfiler.vbwExecuteLine 2622
            .Enabled = True
vbwProfiler.vbwExecuteLine 2623
            .TAG = Settings.HTMLBrowserPath
vbwProfiler.vbwExecuteLine 2624
            .Caption = Abbreviated(.TAG)
vbwProfiler.vbwExecuteLine 2625
        End With
    End If
vbwProfiler.vbwExecuteLine 2626 'B

vbwProfiler.vbwExecuteLine 2627
    chkUseDefaultTextViewer.value = Abs(Settings.bUseDefaultTextViewer)
vbwProfiler.vbwExecuteLine 2628
    If chkUseDefaultTextViewer.value = 1 Then
vbwProfiler.vbwExecuteLine 2629
        With cmdDefaultTextViewerPath
vbwProfiler.vbwExecuteLine 2630
            .Enabled = False
vbwProfiler.vbwExecuteLine 2631
            .TAG = FindDefaultProgram("Text")
vbwProfiler.vbwExecuteLine 2632
            .Caption = Abbreviated(.TAG)
vbwProfiler.vbwExecuteLine 2633
            Settings.TextViewerPath = .TAG
vbwProfiler.vbwExecuteLine 2634
        End With
    Else
vbwProfiler.vbwExecuteLine 2635 'B
vbwProfiler.vbwExecuteLine 2636
        Settings.bUseDefaultTextViewer = 0
vbwProfiler.vbwExecuteLine 2637
        With cmdDefaultTextViewerPath
vbwProfiler.vbwExecuteLine 2638
            .Enabled = True
vbwProfiler.vbwExecuteLine 2639
            .TAG = Settings.TextViewerPath
vbwProfiler.vbwExecuteLine 2640
            .Caption = Abbreviated(.TAG)
vbwProfiler.vbwExecuteLine 2641
        End With
    End If
vbwProfiler.vbwExecuteLine 2642 'B
vbwProfiler.vbwProcOut 95
vbwProfiler.vbwExecuteLine 2643
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
vbwProfiler.vbwProcIn 96
vbwProfiler.vbwExecuteLine 2644
    Select Case KeyCode
'vbwLine 2645:        Case vbKeyEscape
        Case IIf(vbwProfiler.vbwExecuteLine(2645), VBWPROFILER_EMPTY, _
        vbKeyEscape)
vbwProfiler.vbwExecuteLine 2646
            Unload Me
    End Select
vbwProfiler.vbwExecuteLine 2647 'B
vbwProfiler.vbwProcOut 96
vbwProfiler.vbwExecuteLine 2648
End Sub

Private Sub Form_Load()
'the max length of the labels should be around 40 characters?
vbwProfiler.vbwProcIn 97

vbwProfiler.vbwProcOut 97
vbwProfiler.vbwExecuteLine 2649
End Sub

Private Sub Form_Unload(Cancel As Integer)
'//need to save the settings
vbwProfiler.vbwProcIn 98
vbwProfiler.vbwExecuteLine 2650
With Settings
vbwProfiler.vbwExecuteLine 2651
    .bUseDefaultTextViewer = Abs(chkUseDefaultTextViewer)
vbwProfiler.vbwExecuteLine 2652
    .TextViewerPath = cmdDefaultTextViewerPath.TAG
vbwProfiler.vbwExecuteLine 2653
    .bUseDefaultWebBrowser = Abs(chkUseDefaultWebBrowser)
vbwProfiler.vbwExecuteLine 2654
    .HTMLBrowserPath = cmdDefaultBrowserPath.TAG
vbwProfiler.vbwExecuteLine 2655
    .PublishEmailAddress = txtPublishEmailAddress
    '.bQuickStart = Abs(chkQuickStart)  'Disabled 02/16/02 MPJ (obsolete)
    '.bSoundOff = Abs(chkSound)         'Disabled 02/16/02 MPJ (obsolete)
vbwProfiler.vbwExecuteLine 2656
    .bAssociateExt = Abs(chkAssociateExt)
vbwProfiler.vbwExecuteLine 2657
End With

vbwProfiler.vbwProcOut 98
vbwProfiler.vbwExecuteLine 2658
End Sub

Private Sub txtPublishEmailAddress_KeyPress(KeyAscii As Integer)
vbwProfiler.vbwProcIn 99
vbwProfiler.vbwExecuteLine 2659
    If Not IsValidKeyCode(KeyAscii) Then
vbwProfiler.vbwExecuteLine 2660
        KeyAscii = 0
    End If
vbwProfiler.vbwExecuteLine 2661 'B
vbwProfiler.vbwProcOut 99
vbwProfiler.vbwExecuteLine 2662
End Sub

