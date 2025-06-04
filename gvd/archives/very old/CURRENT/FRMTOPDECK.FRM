VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmTopDeck 
   Caption         =   "Top Deck"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2985
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Disabled"
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox Check10 
         Caption         =   "Landing Pad Option"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Angled Flight Deck"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   840
         Width           =   375
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown2 
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin VB.Label Label7 
         Caption         =   "% Covered Deck"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "% Flight Deck"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2730
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5159
            Text            =   "Cost:"
            TextSave        =   "Cost:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5159
            Text            =   "Weight:"
            TextSave        =   "Weight:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MouseIcon       =   "frmTopDeck.frx":0000
   End
End
Attribute VB_Name = "frmTopDeck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
