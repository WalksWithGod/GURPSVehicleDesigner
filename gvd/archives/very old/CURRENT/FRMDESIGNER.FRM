VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmDesigner 
   BackColor       =   &H80000004&
   Caption         =   "Vehicle Designer "
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDesigner.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      MouseIcon       =   "frmDesigner.frx":0442
   End
   Begin ComctlLib.Toolbar Toolbar2 
      Align           =   3  'Align Left
      Height          =   6510
      Left            =   0
      TabIndex        =   3
      Top             =   420
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   11483
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   23
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Arrow Pointer"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Skids subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Wheels subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Tracks subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Leg subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Wings subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Rotars subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Hydrofoil subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Hovercraft subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Superstructure subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Open Mount subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Masts subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Pod subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Turret subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Pop Turret subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Gasbag Envelope subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Arm subassembly"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Armor"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Hardpoint"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Equiptment Pod"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Solar Panel"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Side Car"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Cargo Space"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
      EndProperty
      MouseIcon       =   "frmDesigner.frx":045E
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   4200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   327681
      End
      Begin ComctlLib.ImageList ImageList2 
         Left            =   120
         Top             =   5520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327680
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   10
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":047A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":058C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":069E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":07B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":08C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":09D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":0F16
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":1458
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":1772
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":1A8C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   105
         Top             =   4665
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327680
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   28
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":1DA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":20C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":23DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":26F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":2A0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":2D28
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":3042
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":335C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":3676
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":3990
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":3CAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":3FC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":42DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":45F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":4912
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":4C2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":4F46
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":5260
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":557A
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":5894
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":5BAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":5EC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":61E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":64FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":6816
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":6B30
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":6E4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDesigner.frx":7164
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picSplitter 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   3960
      ScaleHeight     =   2295
      ScaleWidth      =   45
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   40
   End
   Begin VB.PictureBox picTitles 
      BorderStyle     =   0  'None
      Height          =   490
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   8955
      TabIndex        =   14
      Top             =   350
      Width           =   8955
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmDesigner.frx":747E
         Left            =   5640
         List            =   "frmDesigner.frx":7488
         TabIndex        =   18
         Top             =   120
         Width           =   2940
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmDesigner.frx":74A0
         Left            =   3120
         List            =   "frmDesigner.frx":74B6
         TabIndex        =   17
         Top             =   120
         Width           =   2385
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3000
         TabIndex        =   19
         Top             =   75
         Width           =   5925
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "       -- No Vehicle Loaded --"
         Height          =   435
         Left            =   85
         TabIndex        =   15
         Top             =   75
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   10305
      TabIndex        =   7
      Top             =   3840
      Width           =   10335
      Begin VB.CommandButton Command5 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5655
         TabIndex        =   13
         Top             =   45
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Two Page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3300
         TabIndex        =   12
         Top             =   45
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Prev Page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2130
         TabIndex        =   11
         Top             =   45
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Next Page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1065
         TabIndex        =   10
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   9
         Top             =   45
         Width           =   855
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmDesigner.frx":7539
         Left            =   4485
         List            =   "frmDesigner.frx":754F
         TabIndex        =   8
         Text            =   "100%"
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7320
      ScaleHeight     =   1515
      ScaleWidth      =   2115
      TabIndex        =   6
      Top             =   5280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6360
      ScaleHeight     =   1875
      ScaleWidth      =   2595
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   6105
      Left            =   3735
      TabIndex        =   1
      Top             =   840
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   10769
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327680
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmDesigner.frx":7574
      NumItems        =   0
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   6105
      Left            =   795
      TabIndex        =   0
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   10769
      _Version        =   327680
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmDesigner.frx":7590
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   6930
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11298
            Text            =   "Status"
            TextSave        =   "Status"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "5/31/98"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "5:35 PM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmDesigner.frx":78AA
   End
   Begin VB.Image imgSplitter 
      Height          =   2595
      Left            =   3660
      MousePointer    =   9  'Size W E
      Top             =   1020
      Width           =   30
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Preview"
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "Print Setup"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuPerformance 
         Caption         =   "Performance"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuSettings 
         Caption         =   "Vehicle Settings"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Vehicle Options"
      End
      Begin VB.Menu mnuProfiles 
         Caption         =   "Performance Profiles"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mnIndex As Integer  ' Holds the index of a Node
Dim mbIndrag As Boolean ' Flag that signals a Drag Drop operation.
Dim moDragNode As Object ' Item that is being dragged.
Dim sVehicleFile As String ' path and fileName of the file user opens/saves
Dim bLoadedFlag As Boolean  ' tracks whether a .veh file is loaded in the program
Dim bSavedFlag As Boolean ' tracks whether the current .veh has been saved
Dim arrProperties() As String
Dim mbMoving As Boolean
Const sglSplitLimit = 500
Private Sub Combo1_Click()
' Load the frmDesigner.Listview1 with new headers and items from the selected table
Select Case Combo1.Text
    '///////////////////////////////////////////////////////////
    ' WEAPON and WEAPON ACCESSORIES
    Case "Artillery and Guns"
        TableName = "Artillery and Guns"
        LoadListView "Artillery and Guns"
    Case "Beam Weapons"
        TableName = "Beam Weapons"
        LoadListView "Beam Weapons"
    Case "Launchers"
        TableName = "Launchers"
        LoadListView "Launchers"
    Case "Torpedos and Missiles"
        TableName = "Torpedos and Missiles"
        LoadListView "Torpedos and Missiles"
    Case "Weapon Accessories"
        TableName = "Weapon Accessories"
        LoadListView "Weapon Accessories"
    '////////////////////////////////////////////////////////////
    ' PROPULSION SYSTEMS
     Case "Drivetrains"
        TableName = "Drivetrains"
        LoadListView "Drivetrains"
     Case "Jet Engines"
        TableName = "Jet Engines"
        LoadListView "Jet Engines"
     Case "Rockets and Spacedrives"
        TableName = "Rocket Engines and Spacedrives"
        LoadListView "Rocket Engines and Spacedrives"
     Case "Reactionless Thrusters"
        TableName = "Reactionless Thrusters"
        LoadListView "Reactionless Thrusters"
     Case "Stardrives"
        TableName = "Stardrives"
        LoadListView "Stardrives"
     Case "Orion Engine"
        TableName = "Orion Engine"
        LoadListView "Orion Engine"
     Case "Mag-Lev Lifters"
        TableName = "Mag-Lev Lifters"
        LoadListView "Mag-Lev Lifters"
     Case "Sails"
        TableName = "Sails"
        LoadListView "Sails"
     Case "Aquatic Propulsion"
        TableName = "Aquatic Propulsion"
        LoadListView "Aquatic Propulsion"
     Case "Harnesses"
        TableName = "Harness"
        LoadListView "Harness"
     Case "Rowing Positions"
        TableName = "Rowing Positions"
        LoadListView "Rowing Positions"
    Case "Airscrews"
        TableName = "Airscrew Drivetrains"
        LoadListView "Airscrew Drivetrains"
     Case "Contragrav Generators"
        TableName = "Contragravity"
        LoadListView "Contragravity"
    Case "Lifting Gas"
        TableName = "Lifting Gas"
        LoadListView "Lifting Gas"
     Case "Specials"
        TableName = "Special Drives"
        LoadListView "Special Drives"
    '////////////////////////////////////////////////////////////////
    'POWER AND FUEL SYSTEMS
    '////////////////////////////////////////////////////////////////
    Case "Exotic Power Plants"
        TableName = "Exotic Power Plants"
        LoadListView "Exotic Power Plants"
    Case "Gas and MHD Turbines"
        TableName = "Gas and MHD Turbines"
        LoadListView "Gas and MHD Turbines"
    Case "Combustion Engines"
        TableName = "Internal Combustion Engines"
        LoadListView "Internal Combustion Engines"
    Case "Reactors"
        TableName = "Reactors"
        LoadListView "Reactors"
    Case "Steam Engines"
        TableName = "Steam Engines"
        LoadListView "Steam Engines"
    Case "Energy Bank"
        TableName = "Energy Bank"
        LoadListView "Energy Bank"
    Case "Other"
        TableName = "Other Power and Fuel"
        LoadListView "Other Power and Fuel"
    Case "Fuel Storage"
        TableName = "Fuel Storage"
        LoadListView "Fuel Storage"
    Case "Fuel"
        TableName = "Fuel"
        LoadListView "Fuel"
    'MANNED VEHICLE COMPONENTS
    '////////////////////////////////////////////////////////////////
    Case "Maneuver Controls"
        TableName = "Maneuver Controls"
        LoadListView "Maneuver Controls"
    Case "Crew Stations"
        TableName = "Crew Stations"
        LoadListView "Crew Stations"
    Case "Accommodations"
        TableName = "Accommodations"
        LoadListView "Accommodations"
    Case "Environmental Systems"
        TableName = "Environmental Systems"
        LoadListView "Environmental Systems"
    Case "Safety Systems"
        TableName = "Safety Systems"
        LoadListView "Safety Systems"
    'INSTRUMENTS AND ELECTRONICS
    '////////////////////////////////////////////////////////////////
    Case "Communicators"
        TableName = "Communicators"
        LoadListView "Communicators"
    Case "Searchlights"
        TableName = "Searchlights"
        LoadListView "Searchlights"
    Case "Extendable Sensor Periscopes"
        TableName = "Extendable Sensor Periscopes"
        LoadListView "Extendable Sensor Periscopes"
    Case "Sensors"
        TableName = "Sensors"
        LoadListView "Sensors"
    Case "Navigation Systems"
        TableName = "Navigation Systems"
        LoadListView "Navigation Systems"
    Case "Targeting Systems"
        TableName = "Targeting Systems"
        LoadListView "Targeting Systems"
    Case "Electronic Countermeasures"
        TableName = "Electronic Countermeasures"
        LoadListView "Electronic Countermeasures"
    Case "Computers, Terminals & Neural Interfaces"
        TableName = "Computers, Terminals & Neural Interfaces"
        LoadListView "Computers, Terminals & Neural Interfaces"
    Case "Software"
        TableName = "Software"
        LoadListView "Software"
    Case "Neural Interface Systems"
        TableName = "Neural Interface Systems"
        LoadListView "Neural Interface Systems"
    'MISCELLANEOU EQUIPTMENT
    '////////////////////////////////////////////////////////////////
    Case "Arm Motors"
        TableName = "Arm Motors"
        LoadListView "Arm Motors"
    Case "Fire Extinguishers"
        TableName = "Fire Extinguishers"
        LoadListView "Fire Extinguishers"
    Case "Bilge Pumps"
        TableName = "Bilge Pumps"
        LoadListView "Bilge Pumps"
    Case "Lab and Workshops"
        TableName = "Labs and Workshops"
        LoadListView "Labs and Workshops"
    Case "Heavy Equiptment"
        TableName = "Heavy Equiptment"
        LoadListView "Heavy Equiptment"
    Case "Emergency & Medical Equiptment"
        TableName = "Emergency & Medical Equiptment"
        LoadListView "Emergency & Medical Equiptment"
    Case "Entertainment Facilities"
        TableName = "Entertainment Facilities"
        LoadListView "Entertainment Facilities"
    Case "Vehicle Access"
        TableName = "Vehicle Access"
        LoadListView "Vehicle Access"
    Case "Teleport Projectors"
        TableName = "Teleport Projectors"
        LoadListView "Teleport Projectors"
    Case "Security & Dirty Tricks"
        TableName = "Security & Dirty Tricks"
        LoadListView "Security & Dirty Tricks"
    Case "Vehicle Storage"
        TableName = "Vehicle Storage"
        LoadListView "Vehicle Storage"
    Case "Landing Aids"
        TableName = "Landing Aids"
        LoadListView "Landing Aids"
    Case "Fuel Accessories"
        TableName = "Fuel Accessories"
        LoadListView "Fuel Accessories"
    Case "Nuclear Dampers"
        TableName = "Nuclear Dampers"
        LoadListView "Nuclear Dampers"
    Case "Reality Stabilizers"
        TableName = "Reality Stabilizers"
        LoadListView "Reality Stabilizers"
    Case "Modules & Sockets"
        TableName = "Modules & Sockets"
        LoadListView "Modules & Sockets"
End Select
End Sub

Private Sub Combo2_Click()
'Clear the Combo1 items
Combo1.Clear
' Load the frmDesigner.Listview1 with new headers and items from the selected table
Select Case Combo2.Text
    Case "Weaponry"
        With Combo1
            .AddItem "Beam Weapons"
            .AddItem "Artillery and Guns"
            .AddItem "Launchers"
            .AddItem "Torpedos and Missiles"
            .AddItem "Weapon Accessories"
        End With
        Combo1 = "Artillery and Guns"
        TableName = "Artillery and Guns"
        LoadListView "Artillery and Guns"
    Case "Propulsion & Lift"
        With Combo1
            .AddItem "Drivetrains"
            .AddItem "Airscrews"
            .AddItem "Jet Engines"
            .AddItem "Rockets and Spacedrives"
            .AddItem "Reactionless Thrusters"
            .AddItem "Stardrives"
            .AddItem "Orion Engine"
            .AddItem "Mag-Lev Lifters"
            .AddItem "Contragrav Generators"
            .AddItem "Sails"
            .AddItem "Aquatic Propulsion"
            .AddItem "Harnesses"
            .AddItem "Rowing Positions"
            .AddItem "Lifting Gas"
            .AddItem "Specials"
        End With
        Combo1 = "Drivetrains"
        TableName = "Drivetrains"
        LoadListView "Drivetrains"
    Case "Manned Vehicle Components"
        With Combo1
            .AddItem "Maneuver Controls"
            .AddItem "Crew Stations"
            .AddItem "Accommodations"
            .AddItem "Environmental Systems"
            .AddItem "Safety Systems"
        End With
        Combo1 = "Accommodations"
        TableName = "Accommodations"
        LoadListView "Accommodations"
    Case "Power and Fuel"
        With Combo1
            .AddItem "Exotic Power Plants"
            .AddItem "Gas and MHD Turbines"
            .AddItem "Combustion Engines"
            .AddItem "Reactors"
            .AddItem "Steam Engines"
            .AddItem "Energy Bank"
            .AddItem "Other"
            .AddItem "Fuel Storage"
            .AddItem "Fuel"
        End With
        Combo1 = "Combustion Engines"
        TableName = "Internal Combustion Engines"
        LoadListView "Internal Combustion Engines"
    Case "Instruments and Electronics"
        With Combo1
            .AddItem "Communicators"
            .AddItem "Searchlights"
            .AddItem "Extendable Sensor Periscopes"
            .AddItem "Sensors"
            .AddItem "Navigation Systems"
            .AddItem "Targeting Systems"
            .AddItem "Electronic Countermeasures"
            .AddItem "Computers, Terminals & Neural Interfaces"
            .AddItem "Software"
        End With
        Combo1 = "Communicators"
        TableName = "Communicators"
        LoadListView "Communicators"
    Case "Miscellaneous Equiptment"
        With Combo1
            .AddItem "Arm Motors"
            .AddItem "Fire Extinguishers"
            .AddItem "Bilge Pumps"
            .AddItem "Lab and Workshops"
            .AddItem "Heavy Equiptment"
            .AddItem "Emergency & Medical Equiptment"
            .AddItem "Entertainment Facilities"
            .AddItem "Vehicle Access"
            .AddItem "Teleport Projectors"
            .AddItem "Security & Dirty Tricks"
            .AddItem "Vehicle Storage"
            .AddItem "Landing Aids"
            .AddItem "Fuel Accessories"
            .AddItem "Nuclear Dampers"
            .AddItem "Reality Stabilizers"
            .AddItem "Modules & Sockets"
        End With
        Combo1 = "Heavy Equiptment"
        TableName = "Heavy Equiptment"
        LoadListView "Heavy Equiptment"
End Select
End Sub

Private Sub Command1_Click()
PrintDevice 0 ' print to printer
End Sub

Private Sub Command5_Click()

Picture3.Visible = False
Picture1.Visible = False
Picture2.Visible = False

TreeView1.Visible = True
ListView1.Visible = True
Toolbar1.Visible = True
Toolbar2.Visible = True
Combo1.Visible = True
Combo2.Visible = True
Label1.Visible = True
Label2.Visible = True
mnuFile.Visible = True
mnuView.Visible = True
mnuHelp.Visible = True

End Sub

Private Sub Form_Load()
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False  ' set the print preview toolbar to invisible

Load frmProperties
Load frmPerformance

ReadINI ' call procedure to load the default user settings and app settings

Combo1_Click ' this will set the defaults for the listview

' set the default states of buttons and menu items
mnuProperties.Enabled = False
mnuPerformance.Enabled = False
mnuSave.Enabled = False
mnuSaveAs.Enabled = False
Toolbar1.Buttons.Item(3).Enabled = False ' save button
Toolbar1.Buttons.Item(5).Enabled = False ' print button
Toolbar1.Buttons.Item(6).Enabled = False ' print preview

' Set the default toolbar2 button to its mixed state
Toolbar2.Buttons.Item(1).Value = tbrPressed
' Associate ListView1 and TreeView1 with the ImageList
Set ListView1.SmallIcons = frmDesigner.ImageList1
Set TreeView1.ImageList = frmDesigner.ImageList1

bLoadedFlag = False ' initialize the loaded flag (whether a .veh is currently loaded)
bSavedFlag = False ' initialize the saved flag

End Sub



 Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' If there is already a Vehicle loaded
If bLoadedFlag = True Then
    Select Case MsgBox("Save changes to current vehicle?", vbYesNoCancel + 64, "Save Confirmation")
        Case vbYes
            SaveClick  ' call the general save sub
        Case vbCancel
            Cancel = True ' cancel out of the unload event
          
        Case vbNo
    End Select
Else
End If
End Sub

Private Sub Form_Resize()
'need to fix this code up real good
On Error Resume Next
    If Me.Width < 9000 Then Me.Width = 9000
    If Me.Height < 5500 Then Me.Height = 5500
    SizeControls imgSplitter.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
WriteINI ' call the procedure to save the users settings to the "designer.ini"
Set colVehicle = Nothing ' Destroy the collection and all objects in it
UnloadAllForms Me.Name ' call procedure to unload all open forms if any
'Also, here is the unload code if you're calling it from other Subs:
'UnloadAllForms ""
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
' sort the report view according to the header clicked by the user
ListView1.SortKey = ColumnHeader.Index - 1
ListView1.Sorted = True
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iCount As Integer
Dim i As Integer
    Set moDragNode = ListView1.HitTest(X, Y) ' Set the item being dragged.
    Set ListView1.SelectedItem = moDragNode
    Set TreeView1.DropHighlight = Nothing ' make sure there is no drop highlight already set
    If ListView1.SelectedItem Is Nothing Then 'make sure there is an item selected
        Set TreeView1.SelectedItem = Nothing ' this is a hack so that we dont get false drag icon
        Exit Sub
    Else
        ' Fill an array with all the data of the selected item
        iCount = ListView1.ColumnHeaders.Count ' test code
        ReDim arrProperties(iCount)
        For i = 1 To iCount - 1
            arrProperties(i) = ListView1.SelectedItem.SubItems(i)
        Next
    End If
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then ' Signal a Drag operation.
    If ListView1.SelectedItem Is Nothing Then ' make sure there is an item selected
        Exit Sub
    Else
        GetDataTypeandIcons ListView1.SelectedItem, TableName
        ListView1.DragIcon = ImageList1.ListImages(myImage).Picture
        mbIndrag = True ' Set the flag to true.
        ListView1.Drag vbBeginDrag ' Drag operation.
    End If
End If
End Sub

Private Sub mnuHardpoints_Click()
frmHardpoints.Show
End Sub



Private Sub mnuNew_Click()
Dim NewKey As String

' If there is already a Vehicle loaded
If bLoadedFlag = True Then
    Select Case MsgBox("Save changes to current vehicle?", vbYesNoCancel + 64, "Save Confirmation")
        Case vbYes
            SaveClick  ' call the general save sub
        Case vbCancel
            Exit Sub
        Case vbNo
    End Select
Else
End If
frmSettings.Caption = "New Vehicle Settings"
frmSettings.Show vbModal, frmDesigner

End Sub


Private Sub mnuOpen_Click()
' Code to display the Open commond dialog and then handle the loading
' of a file if one is selected
Dim Cancel As Boolean ' detects whether the user clicks cancel at the Open dialog
On Error GoTo ErrorHandler
' If there is already a Vehicle loaded
If bLoadedFlag = True Then
    Select Case MsgBox("Save changes to current vehicle?", vbYesNoCancel + 64, "Save Confirmation")
        Case vbYes
            SaveClick  ' call the general save sub
        Case vbCancel
            Exit Sub
        Case vbNo
    End Select
End If
Cancel = False  ' initialize the cancel button variable for Common dialog
CommonDialog1.Filter = "Vehicle files (*.veh)|*.veh|All files (*.*)|*.*"
CommonDialog1.CancelError = True
CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
'CommonDialog1.InitDir = udtSettings.InitialDir ' this is saved in the INI file
CommonDialog1.ShowOpen
If Not Cancel Then
    ' The Open button was pressed.  Add the code to load the file here
    sVehicleFile = CommonDialog1.filename
    LoadRecords sVehicleFile ' load from the file to fill collection and tree
    frmDesigner.Caption = CommonDialog1.FileTitle & " - Vehicle Designer"
    bLoadedFlag = True  ' there is a vehicle file loaded
    bSavedFlag = True   ' the vehicle has been previously saved
    ' select the first node and set the drophighlight
    TreeView1.Nodes(1).Selected = True 'Select the first Node
    ' Update the properties window
    frmProperties.Caption = "Properties - " & TreeView1.SelectedItem.Text
    Properties_Show TreeView1.SelectedItem.Key
    ' set the states of buttons and menu items
    mnuProperties.Enabled = True
    mnuPerformance.Enabled = True
    mnuSave = True
    mnuSaveAs = True
    Toolbar1.Buttons.Item(3).Enabled = True ' save button
    Toolbar1.Buttons.Item(5).Enabled = True ' print button
    Toolbar1.Buttons.Item(6).Enabled = True ' print preview
End If
Exit Sub

' If the user Cancels out instead of electing to Open a file
ErrorHandler:
If Err.Number = cdlCancel Then
    Cancel = True
    Resume Next
End If
End Sub

Private Sub mnuOptions_Click()
frmOptions.Show vbModal, frmDesigner
End Sub

Private Sub mnuPerformance_Click()
  If mnuPerformance.Checked Then
        ' uncheck it and make the Performance invisible
        mnuPerformance.Checked = False
        frmPerformance.Visible = False
        'uncheck it
    Else
        ' check it and make the Properties visible
        mnuPerformance.Checked = True
        frmPerformance.Visible = True
    End If
End Sub


Private Sub mnuPreview_Click()
Picture3.Top = 0
Picture3.Left = 0
Picture3.Width = frmDesigner.Width

Picture1.Top = Picture3.Height + 250
Picture1.Width = frmDesigner.Width
Picture1.Height = frmDesigner.Height - 2000

TreeView1.Visible = False
ListView1.Visible = False
Toolbar1.Visible = False
Toolbar2.Visible = False
Combo1.Visible = False
Combo2.Visible = False
Label1.Visible = False
Label2.Visible = False
mnuFile.Visible = False
mnuView.Visible = False
mnuHelp.Visible = False

Picture3.Visible = True
'Picture1.Visible = True
Picture2.Visible = True

PrintDevice 1 ' Print to preview window
End Sub

Private Sub PrintDevice(Device As Integer)

Dim X As Integer

      ' Setup (Could be done at design time or in form load)
      ' Make printing stick:
      Picture1.AutoRedraw = True
      ' Add a palette for 256 colors:
      Picture1.Picture = LoadPicture(App.Path + "\PASTEL.DIB")
      ' Set up hidden picture:
      Picture2.AutoRedraw = False
      Picture2.ScaleMode = 3 'Pixels
      Picture2.Visible = False
      Picture2.AutoSize = True
      Picture2.Picture = LoadPicture(App.Path + "\PRINTER.WMF")

      ' This print job can go to the printer or the picture box:
      If Device = 0 Then PrinterFlag = True
      PrintStartDoc Picture1, PrinterFlag, 8.5, 11
      
      ' Position the picture box on the form
      Picture1.Left = (frmDesigner.Width - Picture1.Width) / 2

      ' All the subs use inches:
      PrintBox 1, 1, 6.5, 9
      'PrintLine 1.1, 2, 7.4, 2
      PrintPicture Picture2, 1.1, 1.1, 0.8, 0.8
      'PrintFilledBox 2.1, 1.2, 5.2, 0.7, RGB(200, 200, 200)
      PrintFontName "Courier New"
      PrintCurrentX 2.3
      PrintCurrentY 1.3
      PrintFontSize 16
      PrintPrint "Gurps"
      PrintFontName "Courier New"
      PrintFontSize 12
      PrintCurrentX 1.5
      PrintCurrentY 5
      PrintPrint "Vehicle:"
      PrintCurrentX 1.5
      PrintCurrentY 6.5
      PrintPrint "Components:"
      PrintCurrentX 1.5
      PrintCurrentY 8
      PrintPrint "Statistics:"
      PrintCurrentX 1.5
      PrintCurrentY 9.5
      PrintPrint "Performance:"
      PrintEndDoc
End Sub

Private Sub mnuPrint_Click()
PrintDevice 0 ' print to printer

End Sub

Private Sub mnuProfiles_Click()
frmProfiles.Show vbModal, frmDesigner
End Sub

Private Sub mnuProperties_Click()
  If mnuProperties.Checked Then
        ' uncheck it and make the Properties invisible
        mnuProperties.Checked = False
        frmProperties.Visible = False
    Else
        ' check it and make the Properties visible
        mnuProperties.Checked = True
        frmProperties.Visible = True
    End If
End Sub



Private Sub mnuSave_Click()
SaveClick
End Sub



Private Sub mnuSaveAs_Click()
' Code to load the SaveAs common dialog and to handle the saving
' of the file if the user does want to save the file

Dim Cancel As Boolean

On Error GoTo ErrorHandler
Cancel = False
CommonDialog1.DefaultExt = ".veh"
CommonDialog1.Filter = "Vehicle files (*.veh)|*.veh|All files (*.*)|*.*"
CommonDialog1.CancelError = True
CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
'CommonDialog1.InitDir = udtSettings.InitialDir

CommonDialog1.ShowSave
If Not Cancel Then
    ' A fileName was selected. Add the code to save the file here
    sVehicleFile = CommonDialog1.filename
    WriteRecord sVehicleFile ' save the vehicle to file
    frmDesigner.Caption = CommonDialog1.FileTitle & " - Vehicle Designer"
    bSavedFlag = True ' the vehicle has been saved
    MsgBox "saved"
End If

' If the user Cancels out instead of electing to save a file
ErrorHandler:
MsgBox Error
If Err.Number = cdlCancel Then
    Cancel = True
    Resume Next
End If
End Sub

Private Sub mnuSettings_Click()
frmSettings.Show vbModal, frmDesigner
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
    Case 1 ' New Vehicle
        mnuNew_Click
    Case 2 ' Open
        mnuOpen_Click
    Case 3 ' Save
        SaveClick
    Case 5 ' Print
        PrintDevice 0 ' print to printer
    Case 6 ' Print Preview
        mnuPreview_Click ' print to preview window
End Select
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
Dim i As Integer
    If Button.Value = tbrPressed Then
        Exit Sub
    Else
        For i = 1 To Toolbar2.Buttons.Count
            Toolbar2.Buttons(i).Value = tbrUnpressed
        Next
            Button.Value = tbrPressed
    End If
End Sub


Private Sub TreeView1_DblClick()
TreeView1.Sorted = True

End Sub

Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
' Perhaps i can prohibit dragging from within the tree.  This would make
' recalculation of stats simpler since i only have to recheck when new items
' are added as opposed to when items are moved from one place to another (rearranged)
Dim NewKey As String  ' The key taken from the GetNextKey procedure

If Source Is imgSplitter Then
    SizeControls X
End If

' If user didn't move mouse or released it over an invalid area.
If TreeView1.DropHighlight Is Nothing Then
        mbIndrag = False
        Exit Sub
    Else
    ' If user has dragged an item from the Listview, add that node
    If Source.Name = "ListView1" Then
        NewKey = GetNextKey ' get a new unique key. This key will be shared by the object collection key and its node key in the tree
        AddNewChildNode ListView1.SelectedItem, myImage, mySelectedImage, NewKey 'add the node to the tree
        TreeView1.Nodes.Item(NewKey).EnsureVisible ' expand the tree branch
        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
            AddObject Datatype, NewKey, arrProperties    ' add the new object to the collection
        Else
        ' If the object never got created, delete the node
            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
            TreeView1.Nodes.Remove (NewKey)
            Exit Sub
        End If
        ' Make the newly dragged over item the selected node
        TreeView1.Nodes.Item(NewKey).Selected = True
        TreeView1.DropHighlight = TreeView1.Nodes(NewKey)
        ' Update the properties window
        frmProperties.Caption = "Properties - " & TreeView1.Nodes.Item(NewKey).Text
        Properties_Show NewKey
     ' Else if the user has dragged from within the Treeview, update that node
     Else
        ' Set dragged node's parent property to the target node.
         On Error GoTo checkerror ' To prevent circular errors.
         Set moDragNode.Parent = TreeView1.DropHighlight
         'update the parent node properties of the object in the collection
         With colVehicle.Item(moDragNode.Key)
               .Parent = TreeView1.Nodes.Item(moDragNode.Key).Parent.Key ' root node doesnt have a parent right?
         End With
         ' Release the DropHighlight reference.
         Set TreeView1.DropHighlight = Nothing
         mbIndrag = False
         TreeView1.Nodes.Item(moDragNode.Key).EnsureVisible ' expand the tree branch
         Exit Sub ' Exit if no errors occured.
    End If
    
 End If
're-calc the vehicles statistics
CalcStats
' re-calc vehicle performance figures only
If frmPerformance.Combo1.Text <> "" Then colVehicle.Item(frmPerformance.Combo1.Text).CalcPerformance
' update the Performance window
Performance_Update
checkerror:
    ' Define constants to represent Visual Basic errors code.
    Const CircularError = 35614
    If Err.Number = CircularError Then
        Dim msg As String
        msg = "A node can't be made a child of its own children."
' Display the message box with an exclamation mark icon
        ' and with OK and Cancel buttons.
        If MsgBox(msg, vbExclamation & vbOKCancel) = vbOK Then
            ' Release the DropHighlight reference.
            mbIndrag = False
            Set TreeView1.DropHighlight = Nothing
            Exit Sub
        End If
    End If
End Sub

Private Sub TreeView1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If mbIndrag = True Then
     ' Set DropHighlight to the mouse's coordinates.
       Set TreeView1.DropHighlight = TreeView1.HitTest(X, Y)
      
     End If
End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
' add code for the delete key here
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NewKey As String
Dim i As Integer ' counter
Dim iCurrentButton As Integer ' holds value of current toolbar2button

' On RMK return the arrow pointer. RMK cancels sub assembly selection from toolbar2
If Button = vbRightButton Then
    If Toolbar2.Buttons(1).Value = tbrUnpressed Then
        For i = 1 To Toolbar2.Buttons.Count
            Toolbar2.Buttons(i).Value = tbrUnpressed
        Next
        Toolbar2.Buttons(1).Value = tbrPressed
    End If
End If
Set TreeView1.DropHighlight = TreeView1.HitTest(X, Y)
    'Make sure we are over a Node
    If Not TreeView1.DropHighlight Is Nothing Then
        'Set the Node we are on to be the selected Node
        'if we don't do this it will not be the selected node
        'until we finish clicking on the Node
        TreeView1.SelectedItem = TreeView1.HitTest(X, Y)
            ' on right mouse click, show the Popup menu
            If Button = vbRightButton Then
                ' only show pop up when mouse is arrow pointer
                If TreeView1.MousePointer = ccDefault Then
                    PopupMenu frmPopup.mnuPopup
                Else
                End If
            Else
                If Toolbar2.Buttons(1).Value = tbrPressed Then  ' if already arrow pointer
                Else
                    NewKey = GetNextKey ' get a new unique key.
                    For i = 1 To Toolbar2.Buttons.Count
                        If Toolbar2.Buttons(i).Value = tbrPressed Then
                            iCurrentButton = i
                            Exit For
                        End If
                    Next
                    Select Case iCurrentButton
                    Case 1 'Mouse Arrow Pointer
                        MsgBox "does this get called?"
                        
                    Case 2 'Skid
                        AddNewChildNode "Skids", Skid, Skid, NewKey  'add the node to the tree
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objSkid = New clsSkid
                            colVehicle.Add objSkid, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 3 ' Wheel
                        AddNewChildNode "Wheels", Wheel, Wheel, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objWheel = New clsWheel
                            colVehicle.Add objWheel, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 4 ' Track
                        AddNewChildNode "Tracks", Track, Track, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objTrack = New clsTrack
                            colVehicle.Add objTrack, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 5 ' Leg
                        AddNewChildNode "Leg", Leg, Leg, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objLeg = New clsLeg
                            colVehicle.Add objLeg, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 6 ' Wing
                        AddNewChildNode "Wings", Wing, Wing, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objWing = New clsWing
                            colVehicle.Add objWing, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 7 ' Rotar
                        AddNewChildNode "Rotars", Rotar, Rotar, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objRotar = New clsRotar
                            colVehicle.Add objRotar, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 8 ' Hydrofoil
                        AddNewChildNode "Hydrofoil", Hydrofoil, Hydrofoil, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objHydrofoil = New clsHydrofoil
                            colVehicle.Add objHydrofoil, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 9 ' Hovercraft
                        AddNewChildNode "Hovercraft", Hovercraft, Hovercraft, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objHovercraft = New clsHovercraft
                            colVehicle.Add objHovercraft, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 10 ' Superstructure
                        AddNewChildNode "SuperStructure", Superstructure, Superstructure, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objSuperStructure = New clsSuperStructure
                            colVehicle.Add objSuperStructure, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 11 ' OpenMount
                        AddNewChildNode "OpenMount", OpenMount, OpenMount, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objOpenMount = New clsOpenMount
                            colVehicle.Add objOpenMount, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 12 ' Mast
                        AddNewChildNode "Mast", Mast, Mast, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objMast = New clsMast
                            colVehicle.Add objMast, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 13 ' Pod
                        AddNewChildNode "Pod", Pod, Pod, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objPod = New clsPod
                            colVehicle.Add objPod, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 14 ' Turret
                        AddNewChildNode "Turret", Turret, Turret, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objTurret = New clsTurret
                            colVehicle.Add objTurret, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 15 ' Popturret
                        AddNewChildNode "PopTurret", Popturret, Popturret, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objPopturret = New clsPopTurret
                            colVehicle.Add objPopturret, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 16 ' Gasbag
                       AddNewChildNode "Gasbag", Gasbag, Gasbag, NewKey
                       If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objGasbag = New clsGasbag
                            colVehicle.Add objGasbag, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                       Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 17 ' Arm
                        AddNewChildNode "Arm", Arm, Arm, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objArm = New clsArm
                            colVehicle.Add objArm, NewKey  ' add the new object to the collection
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    Case 18 ' Armor
                        AddNewChildNode "Armor", Armor, Armor, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objArmor = New clsArmor
                            colVehicle.Add objArmor, NewKey ' add the new object
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                        
                        
                    Case 23 ' Cargo
                    
                        AddNewChildNode "Cargo", Cargo, Cargo, NewKey
                        If CheckValidLocation(NewKey, TreeView1.Nodes.Item(NewKey).Parent) Then
                            Set objCargo = New clsCargo
                            colVehicle.Add objCargo, NewKey ' add the new object
                            colVehicle(NewKey).Parent = TreeView1.SelectedItem.Key
                        Else
                        ' If the object never got created, delete the node
                            MsgBox "Item cannot be placed here.  Refer to Vehicles 2nd Ed. for more information"
                            TreeView1.Nodes.Remove (NewKey)
                            Exit Sub
                        End If
                    
                    End Select
                    're-calc the vehicles statistics
                    CalcStats
                    ' re-calc vehicle performance figures only
                    If frmPerformance.Combo1.Text <> "" Then colVehicle.Item(frmPerformance.Combo1.Text).CalcPerformance
                    ' update the Performance window
                    Performance_Update
                    TreeView1.Nodes.Item(NewKey).EnsureVisible ' expand the tree branch
                    ' Make the newly added item the selected node
                    TreeView1.Nodes.Item(NewKey).Selected = True
                    TreeView1.DropHighlight = TreeView1.Nodes(NewKey)
                    ' Update the properties window
                    frmProperties.Caption = "Properties - " & TreeView1.Nodes.Item(NewKey).Text
                    Properties_Show NewKey
                        With colVehicle.Item(NewKey)
                            .Key = frmDesigner.TreeView1.Nodes.Item(NewKey).Key
                           '.Parent = frmDesigner.TreeView1.Nodes.Item(NewKey).Parent.Key ' root node doesnt have a parent right?
                            .Text = frmDesigner.TreeView1.Nodes.Item(NewKey).Text
                            .Image = frmDesigner.TreeView1.Nodes.Item(NewKey).Image
                            .SelectedImage = frmDesigner.TreeView1.Nodes.Item(NewKey).SelectedImage
                        End With
                End If
                Set moDragNode = TreeView1.SelectedItem ' Set the item being dragged.
            End If
    Else
        Set TreeView1.SelectedItem = Nothing
    End If
    
    Set TreeView1.DropHighlight = Nothing
    ' Update the properties window but make sure that a node is selected
    If TreeView1.SelectedItem Is Nothing Then
    Else
        ' hide frmProperties' textbox and combobox if they are still visible
        frmProperties.Combo1.Visible = False
        frmProperties.Text1.Visible = False
        ' Show the properties for the newly selected item
        frmProperties.Caption = "Properties - " & TreeView1.SelectedItem.Text
        Properties_Show TreeView1.SelectedItem.Key
    End If
End Sub


Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer ' counter
Dim iCurrentButton As Integer ' current toolbar button
' if the arrow pointer is selected
If Toolbar2.Buttons(1).Value = tbrPressed Then
    TreeView1.MousePointer = ccDefault
    
Else
  For i = 1 To Toolbar2.Buttons.Count
    If Toolbar2.Buttons(i).Value = tbrPressed Then
        iCurrentButton = i
        Exit For
    End If
  Next
    TreeView1.MouseIcon = ImageList1.ListImages(iCurrentButton).Picture
    TreeView1.MousePointer = ccCustom
End If

    If Button = vbLeftButton Then ' Signal a Drag operation.
        If TreeView1.SelectedItem Is Nothing Then ' prevents false drag icon
            Exit Sub
        Else
            mbIndrag = True ' Set the flag to true.
            TreeView1.DragIcon = ImageList1.ListImages(25).Picture
            TreeView1.Drag vbBeginDrag ' Drag operation.
        End If
    End If
End Sub

Private Sub SaveClick()
    ' If the file has already been saved with a valid fileName
    If bSavedFlag = True Then
        WriteRecord sVehicleFile
        ' Display the save icon in the status bar
    Else
        mnuSaveAs_Click
    End If
End Sub


'///////////////////////////////////////////////////////
'The below series of subs is for the splitter bar
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


Sub SizeControls(X As Single)
    On Error Resume Next
    

    'set the width
    If X < 3500 Then X = 3500
    If X > (Me.Width - 5785) Then X = Me.Width - 5785
    TreeView1.Width = X - Toolbar2.Width
    imgSplitter.Left = X
    ListView1.Left = X + imgSplitter.Width
    ListView1.Width = Me.Width - (TreeView1.Width + Toolbar2.Width + 140)
    Label2.Width = TreeView1.Width - 25
    Label1.Left = Label2.Width + Label2.Left + 70
    picTitles.Width = Me.Width
    Label1.Width = ListView1.Width '- 40
    Combo2.Left = Label1.Left + 85
    Combo1.Left = Combo2.Left + Combo2.Width + 160

    'set the top
    If Toolbar1.Visible Then
        TreeView1.Top = Toolbar1.Height + picTitles.Height - 70
    Else
        TreeView1.Top = picTitles.Height
    End If
    ListView1.Top = TreeView1.Top
    

    'set the height
    If sbStatusBar.Visible Then
        TreeView1.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
    Else
        TreeView1.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
    End If
    

    ListView1.Height = TreeView1.Height
    imgSplitter.Top = TreeView1.Top
    imgSplitter.Height = TreeView1.Height
End Sub
'//////////////////////////////
'End splitter bar specific code
'//////////////////////////////////
