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
    Dim i As Long
    Dim s As String
    On Error GoTo errorhandler
    Dim iFake As Long
    Dim bFakeFlag As Boolean
    '//smoke screen
    For iFake = 1 To 13
        bFakeFlag = True
    Next
    If bFakeFlag = False Then
    End If
    tempbyte = ChopCheck2
    For iFake = 13 To 1
        bFakeFlag = False
    Next
    If bFakeFlag = True Then
    End If

    If Me.TAG = "register" Then
        Me.Caption = "Register"
        Frame1.Visible = True
        Frame2.Visible = True
       
        Call SetBFlag
        
        If bFlag Then
            txtRegNum = "*******************"
            txtRegID = "***"
            For i = 1 To UBound(gsRegName)
                s = s & Chr(gsRegName(i))
            Next
            txtRegName = s
            cmdRegister.Visible = False
            cmdCancel.Caption = "Done"
            cmdCancel.Left = cmdCancel.Left + 950
            cmdCancel.Width = cmdCancel.Width - 400
            Label1.Caption = "Thank you for registering your copy of GURPS Vehicle Designer.  For information regarding program updates, please visit the below web site."
        Else 'its false and we should clear all
            ReDim gsRegName(1)
            ReDim gsRegNum(1)
            gsRegID = Empty
        End If
    End If
    Exit Sub
errorhandler:
    bFlag = False
    ReDim gsRegName(1)
    ReDim gsRegNum(1)
    gsRegID = Empty
End Sub


Private Sub cmdCancel_Click()
    Call SetBFlag
    LocalSetRegisteredToolbarButtonStates
    Unload Me
End Sub

Private Sub cmdRegister_Click()
    'check the reg number and the number to the registry
    Dim i As Long
    
    gsRegID = Val(txtRegID) 'this is the "userID" of the reg'd user
    If Len(txtRegName) = 0 Or Len(txtRegNum) = 0 Or Len(txtRegID) <> 3 Then
        ReDim blankbyte(1) 'make sure any previous reg nubmers are cleared
        gsRegID = Empty
        ReDim gsRegName(1)
        ReDim gsRegNum(1)
        Unload Me
        Exit Sub
    End If
    
    For i = 1 To Len(txtRegName)
        ReDim Preserve gsRegName(i)
        
        gsRegName(i) = Asc(Mid(txtRegName.Text, i, 1))
    Next
    For i = 1 To Len(txtRegNum)
        ReDim Preserve gsRegNum(i)
        gsRegNum(i) = Asc(Mid(txtRegNum.Text, i, 1))
    Next
    tempbyte = ChopCheck2
    Call SetBFlag
    LocalSetRegisteredToolbarButtonStates
    Unload Me
End Sub

Private Sub SetBFlag()
    'if we have a reg'd app then display asterisk for pass
    'regID but show the users name
    '//reg check
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
        bFlag = True
    #End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdCancel_Click
    End Select
End Sub

Private Sub Label19_Click()
    cmdCancel_Click
End Sub

Private Sub LocalSetRegisteredToolbarButtonStates()
    With frmDesigner
        If bFlag Then
            'Set the button states for the Toolbar1
            'mnuPerformance.Enabled = True
            .mnuSave.Enabled = True
            .mnuSaveAs.Enabled = True
            .mnuExport.Enabled = True
            .mnuPrint.Enabled = True
            .mnuPublish.Enabled = True
            .Toolbar1.Buttons.Item(3).Enabled = True ' save button
            .Toolbar1.Buttons.Item(5).Enabled = True ' print preview
            .Toolbar1.Buttons.Item(9).Enabled = True ' publish
        Else
            .mnuSave.Enabled = False
            .mnuSaveAs.Enabled = False
            .mnuExport.Enabled = False
            .mnuPrint.Enabled = False
            .mnuPublish.Enabled = False
            .Toolbar1.Buttons.Item(3).Enabled = False ' save button
            .Toolbar1.Buttons.Item(5).Enabled = False ' print preview
            .Toolbar1.Buttons.Item(9).Enabled = False ' publish
        End If
    End With
End Sub
