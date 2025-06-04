VERSION 5.00
Begin VB.Form frmCredits 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Credits"
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Registration Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   5385
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.makosoft.com/gvd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   45
         TabIndex        =   10
         Top             =   1635
         Width           =   5295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmCredits.frx":0000
         ForeColor       =   &H80000008&
         Height          =   1245
         Left            =   240
         TabIndex        =   9
         Top             =   330
         Width           =   4905
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Enter Registration Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   90
      TabIndex        =   5
      Top             =   2220
      Width           =   5385
      Begin VB.TextBox txtRegID 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   3
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   540
         Width           =   345
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register Now!"
         Height          =   315
         Left            =   450
         TabIndex        =   3
         Top             =   900
         Width           =   1545
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Register Later"
         Height          =   315
         Left            =   2940
         TabIndex        =   4
         Top             =   900
         Width           =   1635
      End
      Begin VB.TextBox txtRegName 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   540
         Width           =   2265
      End
      Begin VB.TextBox txtRegNum 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3270
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   540
         Width           =   1905
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   3120
         X2              =   3240
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Label lblRegName 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name:"
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   210
         TabIndex        =   7
         Top             =   300
         Width           =   825
      End
      Begin VB.Label lblRegNum 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration #:"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2700
         TabIndex        =   6
         Top             =   300
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bFlag As Boolean
Private tempbyte() As Byte

Private Sub Form_Activate()
vbwProfiler.vbwProcIn 100
    Dim i As Long
    Dim s As String
vbwProfiler.vbwExecuteLine 2663
    On Error GoTo errorhandler
    Dim iFake As Long
    Dim bFakeFlag As Boolean
    '//smoke screen
vbwProfiler.vbwExecuteLine 2664
    For iFake = 1 To 13
vbwProfiler.vbwExecuteLine 2665
        bFakeFlag = True
vbwProfiler.vbwExecuteLine 2666
    Next
vbwProfiler.vbwExecuteLine 2667
    If bFakeFlag = False Then
    End If
vbwProfiler.vbwExecuteLine 2668 'B
vbwProfiler.vbwExecuteLine 2669
    tempbyte = ChopCheck2
vbwProfiler.vbwExecuteLine 2670
    For iFake = 13 To 1
vbwProfiler.vbwExecuteLine 2671
        bFakeFlag = False
vbwProfiler.vbwExecuteLine 2672
    Next
vbwProfiler.vbwExecuteLine 2673
    If bFakeFlag = True Then
    End If
vbwProfiler.vbwExecuteLine 2674 'B

vbwProfiler.vbwExecuteLine 2675
    If Me.TAG = "register" Then
vbwProfiler.vbwExecuteLine 2676
        Me.Caption = "Register"
vbwProfiler.vbwExecuteLine 2677
        Frame1.Visible = True
vbwProfiler.vbwExecuteLine 2678
        Frame2.Visible = True

vbwProfiler.vbwExecuteLine 2679
        Call SetBFlag

vbwProfiler.vbwExecuteLine 2680
        If bFlag Then
vbwProfiler.vbwExecuteLine 2681
            txtRegNum = "*******************"
vbwProfiler.vbwExecuteLine 2682
            txtRegID = "***"
vbwProfiler.vbwExecuteLine 2683
            For i = 1 To UBound(gsRegName)
vbwProfiler.vbwExecuteLine 2684
                s = s & Chr(gsRegName(i))
vbwProfiler.vbwExecuteLine 2685
            Next
vbwProfiler.vbwExecuteLine 2686
            txtRegName = s
vbwProfiler.vbwExecuteLine 2687
            cmdRegister.Visible = False
vbwProfiler.vbwExecuteLine 2688
            cmdCancel.Caption = "Done"
vbwProfiler.vbwExecuteLine 2689
            cmdCancel.Left = cmdCancel.Left + 950
vbwProfiler.vbwExecuteLine 2690
            cmdCancel.Width = cmdCancel.Width - 400
vbwProfiler.vbwExecuteLine 2691
            Label1.Caption = "Thank you for registering your copy of GURPS Vehicle Designer.  For information regarding program updates, please visit the below web site."
        Else 'its false and we should clear all
vbwProfiler.vbwExecuteLine 2692 'B
vbwProfiler.vbwExecuteLine 2693
            ReDim gsRegName(1)
vbwProfiler.vbwExecuteLine 2694
            ReDim gsRegNum(1)
vbwProfiler.vbwExecuteLine 2695
            gsRegID = Empty
        End If
vbwProfiler.vbwExecuteLine 2696 'B
    End If
vbwProfiler.vbwExecuteLine 2697 'B
vbwProfiler.vbwProcOut 100
vbwProfiler.vbwExecuteLine 2698
    Exit Sub
errorhandler:
vbwProfiler.vbwExecuteLine 2699
    bFlag = False
vbwProfiler.vbwExecuteLine 2700
    ReDim gsRegName(1)
vbwProfiler.vbwExecuteLine 2701
    ReDim gsRegNum(1)
vbwProfiler.vbwExecuteLine 2702
    gsRegID = Empty
vbwProfiler.vbwProcOut 100
vbwProfiler.vbwExecuteLine 2703
End Sub


Private Sub cmdCancel_Click()
vbwProfiler.vbwProcIn 101
vbwProfiler.vbwExecuteLine 2704
    Call SetBFlag
vbwProfiler.vbwExecuteLine 2705
    LocalSetRegisteredToolbarButtonStates
vbwProfiler.vbwExecuteLine 2706
    Unload Me
vbwProfiler.vbwProcOut 101
vbwProfiler.vbwExecuteLine 2707
End Sub

Private Sub cmdRegister_Click()
    'check the reg number and the number to the registry
vbwProfiler.vbwProcIn 102
    Dim i As Long

vbwProfiler.vbwExecuteLine 2708
    gsRegID = Val(txtRegID) 'this is the "userID" of the reg'd user
vbwProfiler.vbwExecuteLine 2709
    If Len(txtRegName) = 0 Or Len(txtRegNum) = 0 Or Len(txtRegID) <> 3 Then
vbwProfiler.vbwExecuteLine 2710
        ReDim blankbyte(1) 'make sure any previous reg nubmers are cleared
vbwProfiler.vbwExecuteLine 2711
        gsRegID = Empty
vbwProfiler.vbwExecuteLine 2712
        ReDim gsRegName(1)
vbwProfiler.vbwExecuteLine 2713
        ReDim gsRegNum(1)
vbwProfiler.vbwExecuteLine 2714
        Unload Me
vbwProfiler.vbwProcOut 102
vbwProfiler.vbwExecuteLine 2715
        Exit Sub
    End If
vbwProfiler.vbwExecuteLine 2716 'B

vbwProfiler.vbwExecuteLine 2717
    For i = 1 To Len(txtRegName)
vbwProfiler.vbwExecuteLine 2718
        ReDim Preserve gsRegName(i)

vbwProfiler.vbwExecuteLine 2719
        gsRegName(i) = Asc(Mid(txtRegName.Text, i, 1))
vbwProfiler.vbwExecuteLine 2720
    Next
vbwProfiler.vbwExecuteLine 2721
    For i = 1 To Len(txtRegNum)
vbwProfiler.vbwExecuteLine 2722
        ReDim Preserve gsRegNum(i)
vbwProfiler.vbwExecuteLine 2723
        gsRegNum(i) = Asc(Mid(txtRegNum.Text, i, 1))
vbwProfiler.vbwExecuteLine 2724
    Next
vbwProfiler.vbwExecuteLine 2725
    tempbyte = ChopCheck2
vbwProfiler.vbwExecuteLine 2726
    Call SetBFlag
vbwProfiler.vbwExecuteLine 2727
    LocalSetRegisteredToolbarButtonStates
vbwProfiler.vbwExecuteLine 2728
    Unload Me
vbwProfiler.vbwProcOut 102
vbwProfiler.vbwExecuteLine 2729
End Sub

Private Sub SetBFlag()
    'if we have a reg'd app then display asterisk for pass
    'regID but show the users name
    '//reg check
vbwProfiler.vbwProcIn 103
    Dim i As Long
    #If DEBUG_MODE = False Then
        If (IsEmpty(tempbyte) = False) And (IsEmpty(gsRegNum) = False) And (UBound(tempbyte) - LBound(tempbyte) = UBound(gsRegNum) - LBound(gsRegNum)) Then
            For i = 1 To UBound(gsRegNum)  '//dont not check the last 4  numbers of the numbers in this verison of the CHopCheck.  So if they try to produce a serial gen based off tracing this code, it wont work on the other
                If tempbyte(i) = gsRegNum(i) Then
                    bFlag = True
                Else
                    bFlag = False
                    Exit For
                End If
            Next
        Else
            bFlag = False
        End If
    #Else
vbwProfiler.vbwExecuteLine 2730
        bFlag = True
    #End If
vbwProfiler.vbwProcOut 103
vbwProfiler.vbwExecuteLine 2731
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
vbwProfiler.vbwProcIn 104
vbwProfiler.vbwExecuteLine 2732
    Select Case KeyCode
'vbwLine 2733:        Case vbKeyEscape
        Case IIf(vbwProfiler.vbwExecuteLine(2733), VBWPROFILER_EMPTY, _
        vbKeyEscape)
vbwProfiler.vbwExecuteLine 2734
            cmdCancel_Click
    End Select
vbwProfiler.vbwExecuteLine 2735 'B
vbwProfiler.vbwProcOut 104
vbwProfiler.vbwExecuteLine 2736
End Sub

Private Sub Label19_Click()
vbwProfiler.vbwProcIn 105
vbwProfiler.vbwExecuteLine 2737
    cmdCancel_Click
vbwProfiler.vbwProcOut 105
vbwProfiler.vbwExecuteLine 2738
End Sub

Private Sub LocalSetRegisteredToolbarButtonStates()
vbwProfiler.vbwProcIn 106
vbwProfiler.vbwExecuteLine 2739
    With frmDesigner
vbwProfiler.vbwExecuteLine 2740
        If bFlag Then
            'Set the button states for the Toolbar1
            'mnuPerformance.Enabled = True
vbwProfiler.vbwExecuteLine 2741
            .mnuSave.Enabled = True
vbwProfiler.vbwExecuteLine 2742
            .mnuSaveAs.Enabled = True
vbwProfiler.vbwExecuteLine 2743
            .mnuExport.Enabled = True
vbwProfiler.vbwExecuteLine 2744
            .mnuPrint.Enabled = True
vbwProfiler.vbwExecuteLine 2745
            .mnuPublish.Enabled = True
vbwProfiler.vbwExecuteLine 2746
            .Toolbar1.Buttons.Item(3).Enabled = True ' save button
vbwProfiler.vbwExecuteLine 2747
            .Toolbar1.Buttons.Item(5).Enabled = True ' print preview
vbwProfiler.vbwExecuteLine 2748
            .Toolbar1.Buttons.Item(9).Enabled = True ' publish
        Else
vbwProfiler.vbwExecuteLine 2749 'B
vbwProfiler.vbwExecuteLine 2750
            .mnuSave.Enabled = False
vbwProfiler.vbwExecuteLine 2751
            .mnuSaveAs.Enabled = False
vbwProfiler.vbwExecuteLine 2752
            .mnuExport.Enabled = False
vbwProfiler.vbwExecuteLine 2753
            .mnuPrint.Enabled = False
vbwProfiler.vbwExecuteLine 2754
            .mnuPublish.Enabled = False
vbwProfiler.vbwExecuteLine 2755
            .Toolbar1.Buttons.Item(3).Enabled = False ' save button
vbwProfiler.vbwExecuteLine 2756
            .Toolbar1.Buttons.Item(5).Enabled = False ' print preview
vbwProfiler.vbwExecuteLine 2757
            .Toolbar1.Buttons.Item(9).Enabled = False ' publish
        End If
vbwProfiler.vbwExecuteLine 2758 'B
vbwProfiler.vbwExecuteLine 2759
    End With
vbwProfiler.vbwProcOut 106
vbwProfiler.vbwExecuteLine 2760
End Sub


