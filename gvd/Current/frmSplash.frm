VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GURPS Vehicle Designer"
   ClientHeight    =   4185
   ClientLeft      =   2340
   ClientTop       =   1890
   ClientWidth     =   8370
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2888.561
   ScaleMode       =   0  'User
   ScaleWidth      =   7859.863
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   1710
      Left            =   120
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   1158.85
      ScaleMode       =   0  'User
      ScaleWidth      =   1369.55
      TabIndex        =   0
      Top             =   120
      Width           =   2010
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   210
      TabIndex        =   2
      Top             =   2970
      Width           =   8055
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   1
         X1              =   150
         X2              =   7740
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "GURPS® is a trademark of Steve Jackson Games Incorporated, used by permission. All Rights Reserved."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   210
         Width           =   7785
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSplash.frx":20AF
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   180
         TabIndex        =   18
         Top             =   600
         Width           =   7725
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.makosoft.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   6390
      TabIndex        =   22
      Top             =   660
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "© 1999-2002  Mako Software"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   2370
      TabIndex        =   21
      Top             =   660
      Width           =   2115
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Art © 1998  Anthony Affrunti"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   20
      Top             =   900
      Width           =   2355
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dwayne R Corlett"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   5190
      TabIndex        =   17
      Top             =   2130
      Width           =   1245
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Evyn MacDude"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5190
      TabIndex        =   16
      Top             =   2460
      Width           =   1110
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ryan Tomas"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6600
      TabIndex        =   15
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anthony Affrunti"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5190
      TabIndex        =   14
      Top             =   1800
      Width           =   1125
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Troy Guffey"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6600
      TabIndex        =   13
      Top             =   1500
      Width           =   825
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Peter Giroux"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5190
      TabIndex        =   12
      Top             =   1500
      Width           =   870
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Testers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5970
      TabIndex        =   11
      Top             =   1260
      Width           =   645
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nick Ashton"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   6600
      TabIndex        =   10
      Top             =   2130
      Width           =   870
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alan Atkinson"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   6600
      TabIndex        =   9
      Top             =   2460
      Width           =   975
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anthony Affrunti"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2520
      TabIndex        =   8
      Top             =   2550
      Width           =   1125
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Artwork"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2400
      TabIndex        =   7
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programming"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2370
      TabIndex        =   6
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Michael Joseph"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   2490
      TabIndex        =   5
      Top             =   1500
      Width           =   1110
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jeff Wilson"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
      Width           =   780
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code Contributions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2370
      TabIndex        =   3
      Top             =   1800
      Width           =   1620
   End
   Begin VB.Label lblTitle 
      Caption         =   "GURPS Vehicle Designer (version)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   2340
      TabIndex        =   1
      Top             =   120
      Width           =   6315
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Me.Caption = "GURPS Vehicle Designer"
    lblTitle.Caption = "GURPS Vehicle Designer " & App.Major & "." & App.Minor & "." & App.Revision
End Sub



