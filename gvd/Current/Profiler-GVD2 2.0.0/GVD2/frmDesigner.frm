VERSION 5.00
Object = "{E3EEE314-E0F8-11D1-B9BA-0000F654E814}#1.0#0"; "PropList.dll"
Object = "{FF047D84-C3F1-11D2-877E-0040055C08D9}#1.0#0"; "TreeX.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDesigner 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   Caption         =   " GURPS Vehicle Designer "
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDesigner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3885
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   16
      Top             =   4950
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6945
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16431
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3570
      Top             =   5250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesigner.frx":0BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesigner.frx":1106
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesigner.frx":164A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesigner.frx":1B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesigner.frx":20D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesigner.frx":2616
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesigner.frx":2A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesigner.frx":2EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesigner.frx":3312
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesigner.frx":3766
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesigner.frx":3BBA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New Vehicle"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Vehicle"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Battlesuit - Form Fitting"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Battlesuit - Pilot in Body"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Battlesuit - Pilot in Turret"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Export"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Text"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "HTML (classic)"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Use Surface Area Table"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Publish Vehicle to Internet!"
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin VB.ComboBox cboComponents 
         Height          =   315
         ItemData        =   "frmDesigner.frx":400E
         Left            =   4050
         List            =   "frmDesigner.frx":4010
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   0
         Width           =   2235
      End
   End
   Begin TabDlg.SSTab tabSub 
      Height          =   2085
      Index           =   1
      Left            =   4050
      TabIndex        =   3
      Top             =   4860
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   3678
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   441
      MouseIcon       =   "frmDesigner.frx":4012
      TabCaption(0)   =   "Alerts"
      TabPicture(0)   =   "frmDesigner.frx":402E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtInfo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Veh Stats"
      TabPicture(1)   =   "frmDesigner.frx":404A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstStats"
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtInfo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   360
         Width           =   5295
      End
      Begin VB.ListBox lstStats 
         Columns         =   1
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   -74910
         TabIndex        =   7
         Top             =   330
         Width           =   5205
      End
   End
   Begin TabDlg.SSTab tabMain 
      CausesValidation=   0   'False
      Height          =   4365
      Index           =   0
      Left            =   4050
      TabIndex        =   2
      Top             =   390
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   7699
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   529
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      MouseIcon       =   "frmDesigner.frx":4066
      TabCaption(0)   =   "Components"
      TabPicture(0)   =   "frmDesigner.frx":4082
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "PLC1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Profiles/Links"
      TabPicture(1)   =   "frmDesigner.frx":409E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstPropulsionSystems"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Config"
      TabPicture(2)   =   "frmDesigner.frx":40BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "treeLinks"
      Tab(2).Control(1)=   "lstviewLinks"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Debug Output"
      TabPicture(3)   =   "frmDesigner.frx":40D6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "rtbOutput"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Image"
      TabPicture(4)   =   "frmDesigner.frx":40F2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picVehicleImage"
      Tab(4).ControlCount=   1
      Begin MSComctlLib.ListView lstviewLinks 
         CausesValidation=   0   'False
         Height          =   3735
         Left            =   -72540
         TabIndex        =   13
         Top             =   510
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   6588
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.ListBox lstPropulsionSystems 
         Height          =   3660
         Left            =   -74820
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   450
         Width           =   2595
      End
      Begin VB.PictureBox picVehicleImage 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -74910
         ScaleHeight     =   245
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   352
         TabIndex        =   9
         Top             =   390
         Width           =   5340
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3840
         Left            =   90
         TabIndex        =   5
         Top             =   460
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   6773
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin PROPLISTLibCtl.PropListCtl PLC1 
         Height          =   3795
         Left            =   2640
         TabIndex        =   6
         Top             =   480
         Width           =   2730
         _cx             =   4815
         _cy             =   6694
         BackColor       =   -2147483643
         ForeColor       =   0
         BorderStyle     =   1
         Appearance      =   1
         GridColor       =   12632256
         Sorted          =   0   'False
         ShowDescription =   0   'False
         ItemDisabledBackColor=   -2147483643
         ItemDisabledTextColor=   192
         SelectedItemBackColor=   -2147483635
         SelectedItemForeColor=   -2147483634
         PopupMenuEnabled=   0   'False
         TabToNextCell   =   0   'False
         DrawFocusRect   =   -1  'True
         UserCanSizeColumn=   0   'False
         LeftColumnBackColor=   -2147483643
         LeftColumnForeColor=   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbOutput 
         Height          =   3825
         Left            =   -74910
         TabIndex        =   10
         Top             =   390
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   6747
         _Version        =   393217
         ReadOnly        =   -1  'True
         TextRTF         =   $"frmDesigner.frx":410E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TreeXLibCtl.TreeX treeLinks 
         CausesValidation=   0   'False
         Height          =   3675
         Left            =   -74820
         TabIndex        =   11
         Top             =   510
         Width           =   2205
         _cx             =   1368919857
         _cy             =   1368922450
         BorderStyle     =   4
         BackColor       =   -2147483643
         ForeColor       =   0
         PicturePosition =   17
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         MousePointer    =   0
         AutoHScroll     =   -1  'True
         AutoVScroll     =   -1  'True
      End
   End
   Begin TreeXLibCtl.TreeX treeVehicle 
      Height          =   5655
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   3690
      _cx             =   1368922477
      _cy             =   1368925943
      BorderStyle     =   4
      BackColor       =   -2147483643
      ForeColor       =   0
      PicturePosition =   17
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      AutoHScroll     =   -1  'True
      AutoVScroll     =   -1  'True
   End
   Begin MSComctlLib.TabStrip tabVehicle 
      Height          =   6480
      Left            =   0
      TabIndex        =   14
      Top             =   420
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   11430
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Begin VB.Menu mnuNewVehicle 
            Caption         =   "Vehicle"
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuBattleSuitFF 
            Caption         =   "Battlesuit - Form Fitting "
         End
         Begin VB.Menu mnuBattleSuitBody 
            Caption         =   "Battlesuit - Pilot in Body"
         End
         Begin VB.Menu mnuBattleSuitTurret 
            Caption         =   "Battlesuit - Pilot in Turret"
         End
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
         Enabled         =   0   'False
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintSetup 
         Caption         =   "Print Set&up"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export"
         Enabled         =   0   'False
         Begin VB.Menu mnuTextSlim 
            Caption         =   "&Text File (slimline)"
            Shortcut        =   ^T
         End
         Begin VB.Menu mnuText 
            Caption         =   "Te&xt File (classic)"
         End
         Begin VB.Menu mnuHTML_Tables 
            Caption         =   "&HTML (Tables)"
            Shortcut        =   ^H
         End
         Begin VB.Menu mnuHTML_classic 
            Caption         =   "HT&ML (classic)"
         End
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuConfigure 
      Caption         =   "&Performance"
      Begin VB.Menu mnuPerformanceProfiles 
         Caption         =   "Add Performance Profile"
         Begin VB.Menu mnuAddPerformanceProfile 
            Caption         =   "Ground - Wheels"
            Index           =   0
         End
         Begin VB.Menu mnuAddPerformanceProfile 
            Caption         =   "Ground - Tracks"
            Index           =   1
         End
         Begin VB.Menu mnuAddPerformanceProfile 
            Caption         =   "Ground - Skids"
            Index           =   2
         End
         Begin VB.Menu mnuAddPerformanceProfile 
            Caption         =   "Ground - Legs"
            Index           =   3
         End
         Begin VB.Menu mnuAddPerformanceProfile 
            Caption         =   "Ground - Flexibody"
            Index           =   4
         End
         Begin VB.Menu mnuAddPerformanceProfile 
            Caption         =   "Air"
            Index           =   5
         End
         Begin VB.Menu mnuAddPerformanceProfile 
            Caption         =   "Hovercraft"
            Index           =   6
         End
         Begin VB.Menu mnuAddPerformanceProfile 
            Caption         =   "Mag-Lev"
            Index           =   7
         End
         Begin VB.Menu mnuAddPerformanceProfile 
            Caption         =   "Water"
            Index           =   8
         End
         Begin VB.Menu mnuAddPerformanceProfile 
            Caption         =   "Submerged"
            Index           =   9
         End
         Begin VB.Menu mnuAddPerformanceProfile 
            Caption         =   "Space"
            Index           =   10
         End
      End
   End
   Begin VB.Menu mnuPower 
      Caption         =   "&Create Config"
      Begin VB.Menu mnuPowerCreateNew 
         Caption         =   "New Power Config"
      End
      Begin VB.Menu mnuFuelCreateNew 
         Caption         =   "New Fuel Config"
      End
      Begin VB.Menu mnuWeaponCreateNew 
         Caption         =   "New Weapon Link"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuDesignCheck 
         Caption         =   "&Design Check"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuUnitConversion 
         Caption         =   "&Unit Conversion Rules"
      End
      Begin VB.Menu mnuConfigureGVD 
         Caption         =   "&Configure"
      End
      Begin VB.Menu mnuUseSurfaceAreaTable 
         Caption         =   "Use &Surface Area Table"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearMessages 
         Caption         =   "C&lear Warning Panel Messages"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPublish 
         Caption         =   "&Publish Vehicle"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "&Register"
      End
      Begin VB.Menu mnuVersion 
         Caption         =   "&Version"
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "treeVehiclePopup"
         Visible         =   0   'False
         Begin VB.Menu mnuNotes 
            Caption         =   "Notes"
         End
         Begin VB.Menu mnuSep7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRename 
            Caption         =   "Rename"
         End
         Begin VB.Menu mnuRevert 
            Caption         =   "Revert"
         End
         Begin VB.Menu mnuSeperator 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCopy 
            Caption         =   "Copy"
         End
         Begin VB.Menu mnuPaste 
            Caption         =   "Paste"
         End
         Begin VB.Menu mnuSeperator2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSaveComponent 
            Caption         =   "Save Component"
         End
         Begin VB.Menu mnuSeperator3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDelete 
            Caption         =   "Delete"
         End
      End
      Begin VB.Menu mnuLinksPopup 
         Caption         =   "treeLinksPopup"
         Visible         =   0   'False
         Begin VB.Menu mnuSupplierAddToNewGroup 
            Caption         =   "Add To New Group"
         End
         Begin VB.Menu mnuConsumerDeleteAll 
            Caption         =   "Delete All Consumers"
         End
         Begin VB.Menu mnuConsumerDelete 
            Caption         =   "Delete Consumer"
         End
      End
   End
End
Attribute VB_Name = "frmDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 ' used for treeview ajustments (might not need after switching to TreeX? TODO: Investigate this
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const GWL_STYLE = -16&
Private Const TVM_SETBKCOLOR = 4381&
Private Const TVM_GETBKCOLOR = 4383&
Private Const TVS_HASLINES = 2&
Private Const COLOR_WINDOW = 5

Private Const MAX_NODES = 1000
Private Const NEWE_VEHICLE_COMPONENT = "\components\templates\vehicle.cmp"
Private Const NEW_VEHICLE_CAPTION = "untitled - Vehicle Designer"
Private Const NEW_VEHICLE_DEFAULT_FILENAME = "untitled"
Private Const OPEN_SAVE_DIALOG_FILTER = "Vehicle files (*.veh)|*.veh|All files (*.*)|*.*"
Private Const FIRST_TASK_MESSAGE = "As your first task, set the Tech Level and Structural options of the Body.  The default Tech Level and Structure for the subassemblies or components you will add is determined by those of the Body."
Private Const DEFAULT_SPLITTER1_LEFT = 270
Private Const DEFAULT_HSPLITTER_TOP = 325

#If DEBUG_MODE Then
    Private Const APP_CAPTION = "          DEBUG MODE             DEBUG MODE             DEBUG MODE"
#Else
    Private Const APP_CAPTION = "GURPS Vehicle Designer 2.0"
#End If

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type SIZE
    cx As Long
    cy As Long
End Type

Private Type uTreeLinkItem
    hParent As Long
    hwnd As Long
    iIndex As Long
    iGroupIndex As Long
    iNewGroupIndex As Long
    sText As String
    sKey As String
End Type

Public Enum VIEW_MODE
    component_view = 1
    performance_view = 2
    links_view = 3
    output_view = 4
    image_view = 5
    'crew_view = 3
End Enum

Private m_uTreeLinksNodeData As uTreeLinkItem
Private m_oLstviewLinkDragItem As MSComctlLib.ListItem
Private m_bTreeLinksDrag As Boolean
Private m_sVehicleFile As String ' path and fileName of the file user opens/saves
Private m_bLoadedFlag As Boolean  ' tracks whether a .veh file is loaded in the program
Private m_bSavedFlag As Boolean ' tracks whether the current .veh has been saved
Private m_sClipboard As String  ' holds the last "copied" node/branch for the purposes of copy/paste

'Dim WithEvents oHSplitter As cSplitter  '07/10/02 MPJ  Dont seem to be using the events for these custom widgets so redim'ing them sans WithEvents
'Dim WithEvents oVSplitter1 As cSplitter
'Dim WithEvents oVSplitter2 As cSplitter
Private oHSplitter As cSplitter
Attribute oHSplitter.VB_VarHelpID = -1
Private oVSplitter1 As cSplitter
Attribute oVSplitter1.VB_VarHelpID = -1
Private oVSplitter2 As cSplitter
Attribute oVSplitter2.VB_VarHelpID = -1
Private m_oCmdlg As clsCmdlg
Private WithEvents m_oCBO  As clsCompList
Attribute m_oCBO.VB_VarHelpID = -1
Private m_oManager As Vehicles.cManager
Private m_oRecentFiles As clsRecentFileManager

Private Sub SetHooks()
    ' subclass the combo box so we can create our own drop down list
vbwProfiler.vbwProcIn 285
vbwProfiler.vbwExecuteLine 8462
    Call SetHook(cboComponents.hwnd, "test1", True)
    ' subclass the tabstrips so that we know when they are resized
vbwProfiler.vbwExecuteLine 8463
    Call SetHook(tabMain(0).hwnd, "test2", True)
    'Call SetHook(tabSub(1).hWnd, True)
vbwProfiler.vbwProcOut 285
vbwProfiler.vbwExecuteLine 8464
End Sub

Private Sub RemoveHooks()
    ' remove the subclassing hook from the cboComponents
vbwProfiler.vbwProcIn 286
vbwProfiler.vbwExecuteLine 8465
    Call SetHook(cboComponents.hwnd, "test1", False)
vbwProfiler.vbwExecuteLine 8466
    Call SetHook(tabMain(0).hwnd, "test2", False)
    'Call SetHook(tabSub(1).hWnd, False)
vbwProfiler.vbwProcOut 286
vbwProfiler.vbwExecuteLine 8467
End Sub

Private Sub Form_Load()
vbwProfiler.vbwProcIn 287
    Dim lngStyle As Long

    ' set our subclass hooks
vbwProfiler.vbwExecuteLine 8468
    Call SetHooks

    'set our common dialog object
    'todo: verify in .Unload() all these are being unloaded
vbwProfiler.vbwExecuteLine 8469
    Set m_oCmdlg = New clsCmdlg
vbwProfiler.vbwExecuteLine 8470
    Set m_oRecentFiles = New clsRecentFileManager
vbwProfiler.vbwExecuteLine 8471
    Set m_oRecentFiles.Menu = frmDesigner.mnuRecent(0)
vbwProfiler.vbwExecuteLine 8472
    Set m_oManager = New Vehicles.cManager
    'todo: 'set a reference to our text dsplay area
    'm_oFactory.SetMessageTextBox txtInfo
    'm_oFactory.FormatString = Settings.FormatString

vbwProfiler.vbwExecuteLine 8473
    Me.Caption = APP_CAPTION
vbwProfiler.vbwExecuteLine 8474
    Me.ScaleMode = vbTwips
vbwProfiler.vbwExecuteLine 8475
    lstviewLinks.LabelEdit = lvwManual
vbwProfiler.vbwExecuteLine 8476
    cboComponents.Text = "Components"
vbwProfiler.vbwExecuteLine 8477
    tabVehicle.Tabs.Remove (1) ' dont know why this stupid control always adds a tab after ive deleted them all at design time... oh well.

    ' Set listbox to report mode
vbwProfiler.vbwExecuteLine 8478
    With ListView1
vbwProfiler.vbwExecuteLine 8479
        .View = lvwReport
vbwProfiler.vbwExecuteLine 8480
        .ColumnHeaders.Add 1, , "Component"
        'hide the column headers
vbwProfiler.vbwExecuteLine 8481
        .HideColumnHeaders = True
vbwProfiler.vbwExecuteLine 8482
    End With
vbwProfiler.vbwExecuteLine 8483
    PLC1.Clear

    'must now set the screen display
vbwProfiler.vbwExecuteLine 8484
    With Settings
vbwProfiler.vbwExecuteLine 8485
        GVDVehiclesPath = .InitialDir
vbwProfiler.vbwExecuteLine 8486
        Me.mnuUseSurfaceAreaTable.Checked = Abs(.bUseSurfaceAreaTable)
vbwProfiler.vbwExecuteLine 8487
        Me.Toolbar1.Buttons.Item(9).value = Abs(.bUseSurfaceAreaTable)

        'check to see if the desktop resolution has changed since the last run
vbwProfiler.vbwExecuteLine 8488
        If (Screen.Width <> .DesktopX) And (Screen.Height <> .DesktopY) Then
vbwProfiler.vbwExecuteLine 8489
            MsgBox "Desktop settings have changed since last run. GVD will use default settings."
vbwProfiler.vbwExecuteLine 8490
            If .windowstate = vbMinimized Then
vbwProfiler.vbwExecuteLine 8491
                 .windowstate = vbNormal
            End If
vbwProfiler.vbwExecuteLine 8492 'B
vbwProfiler.vbwExecuteLine 8493
            Me.windowstate = .windowstate
vbwProfiler.vbwExecuteLine 8494
            MoveWindow Me.hwnd, 0, 0, 640, 480, 0
vbwProfiler.vbwExecuteLine 8495
            .Splitter1 = DEFAULT_SPLITTER1_LEFT
vbwProfiler.vbwExecuteLine 8496
            .Splitter2 = tabMain(0).Width / Screen.TwipsPerPixelX / 2
vbwProfiler.vbwExecuteLine 8497
            .HSplitter = DEFAULT_HSPLITTER_TOP
        Else
vbwProfiler.vbwExecuteLine 8498 'B
vbwProfiler.vbwExecuteLine 8499
            Me.windowstate = .windowstate
vbwProfiler.vbwExecuteLine 8500
            MoveWindow Me.hwnd, .FormLeft, .FormTop, .FormWidth, .FormHeight, 0
        End If
vbwProfiler.vbwExecuteLine 8501 'B

vbwProfiler.vbwExecuteLine 8502
        Call ConfigureMainSplitters
        ' set the main tab to display the "Components Tab"
vbwProfiler.vbwExecuteLine 8503
        Call SetViewMode(component_view)
        ' set the sub tab to start with the "Alerts Tab"
vbwProfiler.vbwExecuteLine 8504
        tabSub(1).Tab = 0

        'set our recent file list
vbwProfiler.vbwExecuteLine 8505
        m_oRecentFiles.AddRecentFile Settings.Recent2
vbwProfiler.vbwExecuteLine 8506
        m_oRecentFiles.AddRecentFile Settings.Recent2
vbwProfiler.vbwExecuteLine 8507
        m_oRecentFiles.AddRecentFile Settings.Recent3
vbwProfiler.vbwExecuteLine 8508
        m_oRecentFiles.AddRecentFile Settings.Recent4
vbwProfiler.vbwExecuteLine 8509
        m_oRecentFiles.AddRecentFile Settings.Recent5
vbwProfiler.vbwExecuteLine 8510
    End With
vbwProfiler.vbwProcOut 287
vbwProfiler.vbwExecuteLine 8511
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' If there is already a Vehicle loaded
vbwProfiler.vbwProcIn 288
vbwProfiler.vbwExecuteLine 8512
    If m_bLoadedFlag And p_bChangedFlag Then
vbwProfiler.vbwExecuteLine 8513
        Select Case MsgBox("Save changes to current vehicle?", vbYesNoCancel + 64, "Save Confirmation")
'vbwLine 8514:            Case vbYes
            Case IIf(vbwProfiler.vbwExecuteLine(8514), VBWPROFILER_EMPTY, _
        vbYes)
vbwProfiler.vbwExecuteLine 8515
                SaveVehicle  ' call the general save sub
'vbwLine 8516:            Case vbCancel
            Case IIf(vbwProfiler.vbwExecuteLine(8516), VBWPROFILER_EMPTY, _
        vbCancel)
vbwProfiler.vbwExecuteLine 8517
                Cancel = True ' cancel out of the unload event
        End Select
vbwProfiler.vbwExecuteLine 8518 'B
    End If
vbwProfiler.vbwExecuteLine 8519 'B
vbwProfiler.vbwProcOut 288
vbwProfiler.vbwExecuteLine 8520
End Sub

Private Sub Form_Unload(Cancel As Integer)
vbwProfiler.vbwProcIn 289
vbwProfiler.vbwExecuteLine 8521
    Me.Hide
vbwProfiler.vbwExecuteLine 8522
    DoEvents
vbwProfiler.vbwExecuteLine 8523
    On Error Resume Next

vbwProfiler.vbwExecuteLine 8524
    Call RemoveHooks
vbwProfiler.vbwExecuteLine 8525
    Set m_oLstviewLinkDragItem = Nothing
vbwProfiler.vbwExecuteLine 8526
    Set m_oManager = Nothing

    ' prior to calling WriteINI and prior to destroying the splitters, update the Settings UDT
vbwProfiler.vbwExecuteLine 8527
    With Settings
vbwProfiler.vbwExecuteLine 8528
        .FormLeft = Me.ScaleLeft / Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 8529
        .FormTop = Me.ScaleTop / Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 8530
        .FormWidth = Me.ScaleWidth / Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 8531
        .FormHeight = Me.ScaleHeight / Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 8532
        .HSplitter = oHSplitter.Top
vbwProfiler.vbwExecuteLine 8533
        .Splitter1 = oVSplitter1.Left
vbwProfiler.vbwExecuteLine 8534
    End With

    ' destroy the splitters
vbwProfiler.vbwExecuteLine 8535
    Set oHSplitter = Nothing
vbwProfiler.vbwExecuteLine 8536
    Set oVSplitter1 = Nothing
vbwProfiler.vbwExecuteLine 8537
    Set oVSplitter2 = Nothing

    ' destroy our dropdown
vbwProfiler.vbwExecuteLine 8538
    Set m_oCBO = Nothing
vbwProfiler.vbwExecuteLine 8539
    Set m_oCmdlg = Nothing
vbwProfiler.vbwExecuteLine 8540
    Set m_oRecentFiles = Nothing

vbwProfiler.vbwExecuteLine 8541
    WriteINI ' call the procedure to save the users settings to the "designer.ini"
vbwProfiler.vbwExecuteLine 8542
    WriteLicenseFile

vbwProfiler.vbwExecuteLine 8543
    UnloadAllForms Me.Name
vbwProfiler.vbwProcOut 289
vbwProfiler.vbwExecuteLine 8544
End Sub

'Load a Standard Vehicle, Not a battlesuit
Private Sub mnuNewVehicle_Click()
vbwProfiler.vbwProcIn 290
vbwProfiler.vbwExecuteLine 8545
    LoadNewVehicle
vbwProfiler.vbwProcOut 290
vbwProfiler.vbwExecuteLine 8546
End Sub

Public Sub LoadNewVehicle()
vbwProfiler.vbwProcIn 291
    Dim sFilePath As String
vbwProfiler.vbwExecuteLine 8547
    sFilePath = App.Path & NEWE_VEHICLE_COMPONENT
vbwProfiler.vbwExecuteLine 8548
    If LoadVehicle(sFilePath) Then
vbwProfiler.vbwExecuteLine 8549
        Call setGUID
        ' NOTE: This must be done AFTER a vehicle is loaded (either new or from command line)
        ' set the toolbar states for registered or unregistered version
vbwProfiler.vbwExecuteLine 8550
        SetRegisteredToolbarButtonStates
    End If
vbwProfiler.vbwExecuteLine 8551 'B
vbwProfiler.vbwProcOut 291
vbwProfiler.vbwExecuteLine 8552
End Sub

Private Function InitVehicleInstance() As Long
vbwProfiler.vbwProcIn 292
    Dim sKey As String
    Dim lngTempIcon As Long
    Dim lp As Long
    Dim oBase As Vehicles.cINode

vbwProfiler.vbwExecuteLine 8553
    lngTempIcon = 1

    'clear any ghost GUI stuff from previous vehicle if any
vbwProfiler.vbwExecuteLine 8554
    lstviewLinks.ListItems.Clear              '<--- Make sure we dont leave ghosts in the lists
vbwProfiler.vbwExecuteLine 8555
    treeLinks.RemoveAllItems                  '<--- Make sure we dont leave ghosts in the lists
vbwProfiler.vbwExecuteLine 8556
    lstPropulsionSystems.Clear
vbwProfiler.vbwExecuteLine 8557
    txtInfo = ""
vbwProfiler.vbwExecuteLine 8558
    p_ActiveNode.Key = "" 'todo: i should try and ditch these active n ode crap
                          ' the only reason left to use them is I was tracking
                          ' which were the last accessed "profiles" or "links" etc
                          ' so that when a user switched the tabs from say components listing
                          ' to profiles or links, it would automatically load the correct
                          ' profile display in the tabbed display.
                          ' However, i still dont need to use that lame struct anymore
                          ' since now all components are cINode and i can derive description,parent, etc
                          ' from any of them
                          ' NOTE: With ability to edit multiple vehicles, tracking the last of the various node
                          ' types is a bit more of a chore since we need to track depending on which vehicle is
                          ' selected.  We also need to delete references when vehicles are closed.

vbwProfiler.vbwExecuteLine 8559
    p_ActiveNode.CustomDescription = ""
vbwProfiler.vbwExecuteLine 8560
    p_ActiveNode.Parent = ""
vbwProfiler.vbwExecuteLine 8561
    p_ActiveNode.Datatype = 0
vbwProfiler.vbwExecuteLine 8562
    p_ActiveNode.ParentDataType = 0

vbwProfiler.vbwExecuteLine 8563
    InfoPrint 1, FIRST_TASK_MESSAGE
vbwProfiler.vbwExecuteLine 8564
    frmDesigner.Caption = NEW_VEHICLE_CAPTION
vbwProfiler.vbwExecuteLine 8565
    m_sVehicleFile = NEW_VEHICLE_DEFAULT_FILENAME
vbwProfiler.vbwExecuteLine 8566
    frmDesigner.treeVehicle.RemoveAllItems
vbwProfiler.vbwExecuteLine 8567
    Call FillListViewDefaults

    'Set the flags and states
vbwProfiler.vbwExecuteLine 8568
    p_bChangedFlag = False 'JAW 2000.05.07 of course no changes can have been made yet
vbwProfiler.vbwExecuteLine 8569
    InitVehicleInstance = lp
vbwProfiler.vbwProcOut 292
vbwProfiler.vbwExecuteLine 8570
    Exit Function
err:
vbwProfiler.vbwExecuteLine 8571
    Debug.Print "frmDesigner.InitVehicleInstance -- Error#" & err.Number & "  " & err.Description
vbwProfiler.vbwProcOut 292
vbwProfiler.vbwExecuteLine 8572
End Function

Private Function OpenNewVehicle() As Long
    ' Check if there is already a vehicle loaded and give user option to save it
vbwProfiler.vbwProcIn 293
vbwProfiler.vbwExecuteLine 8573
    If m_bLoadedFlag And p_bChangedFlag Then
vbwProfiler.vbwExecuteLine 8574
        Select Case MsgBox("Save changes to current vehicle?", vbYesNoCancel + 64, "Save Confirmation")
'vbwLine 8575:            Case vbYes
            Case IIf(vbwProfiler.vbwExecuteLine(8575), VBWPROFILER_EMPTY, _
        vbYes)
vbwProfiler.vbwExecuteLine 8576
                SaveVehicle  ' call the general save sub
'vbwLine 8577:            Case vbCancel
            Case IIf(vbwProfiler.vbwExecuteLine(8577), VBWPROFILER_EMPTY, _
        vbCancel)
vbwProfiler.vbwExecuteLine 8578
                OpenNewVehicle = False
vbwProfiler.vbwProcOut 293
vbwProfiler.vbwExecuteLine 8579
                Exit Function 'user clicks Cancel.. Exit the sub
'vbwLine 8580:            Case vbNo
            Case IIf(vbwProfiler.vbwExecuteLine(8580), VBWPROFILER_EMPTY, _
        vbNo)
        End Select
vbwProfiler.vbwExecuteLine 8581 'B
    End If
vbwProfiler.vbwExecuteLine 8582 'B
vbwProfiler.vbwExecuteLine 8583
    OpenNewVehicle = True
vbwProfiler.vbwProcOut 293
vbwProfiler.vbwExecuteLine 8584
End Function

Private Sub mnuOpen_Click()
vbwProfiler.vbwProcIn 294
    Dim i As Integer
    Dim oCDLG As clsCmdlg

    ' Code to display the Open commond dialog and then handle the loading
    ' of a file if one is selected
    Dim bSuccess As Boolean ' detects whether the user clicks cancel at the Open dialog
vbwProfiler.vbwExecuteLine 8585
    On Error GoTo errorhandler

    ' If there is already a Vehicle loaded
vbwProfiler.vbwExecuteLine 8586
    If OpenNewVehicle Then
vbwProfiler.vbwExecuteLine 8587
        Set oCDLG = New clsCmdlg
vbwProfiler.vbwExecuteLine 8588
        oCDLG.Filter = OPEN_SAVE_DIALOG_FILTER
vbwProfiler.vbwExecuteLine 8589
        oCDLG.CancelError = True

        Dim oFile As FileSystemObject
vbwProfiler.vbwExecuteLine 8590
        Set oFile = New FileSystemObject
vbwProfiler.vbwExecuteLine 8591
        If oFile.FolderExists(Settings.VehiclesOpenPath) Then
vbwProfiler.vbwExecuteLine 8592
            oCDLG.InitialDir = Settings.VehiclesOpenPath
        Else
vbwProfiler.vbwExecuteLine 8593 'B
vbwProfiler.vbwExecuteLine 8594
            oCDLG.InitialDir = App.Path
        End If
vbwProfiler.vbwExecuteLine 8595 'B
vbwProfiler.vbwExecuteLine 8596
        bSuccess = oCDLG.ShowOpen(Me.hwnd)

vbwProfiler.vbwExecuteLine 8597
        If bSuccess Then
vbwProfiler.vbwExecuteLine 8598
            p_sGUID = ""
vbwProfiler.vbwExecuteLine 8599
            If LoadVehicle(oCDLG.cFileName(1)) Then
vbwProfiler.vbwExecuteLine 8600
                Call setGUID
            End If
vbwProfiler.vbwExecuteLine 8601 'B
        End If
vbwProfiler.vbwExecuteLine 8602 'B
    End If
vbwProfiler.vbwExecuteLine 8603 'B
vbwProfiler.vbwProcOut 294
vbwProfiler.vbwExecuteLine 8604
    Exit Sub
errorhandler:
vbwProfiler.vbwExecuteLine 8605
    Debug.Print "frmDesigner.mnuOpen_Click() - Error # " & err.Number & " " & err.Description
vbwProfiler.vbwProcOut 294
vbwProfiler.vbwExecuteLine 8606
End Sub

Public Function LoadVehicle(sFilePath As String) As Boolean
vbwProfiler.vbwProcIn 295
    Dim lptr As Long
    Dim lngErrorCode As Long
    Dim index As Long
    Dim sKey As String

vbwProfiler.vbwExecuteLine 8607
    Debug.Print sFilePath

        'prepare for new vehicle
vbwProfiler.vbwExecuteLine 8608
    If OpenNewVehicle Then
vbwProfiler.vbwExecuteLine 8609
        Call InitVehicleInstance
vbwProfiler.vbwExecuteLine 8610
        lptr = m_oManager.createVehicle(sFilePath)
vbwProfiler.vbwExecuteLine 8611
        If lptr <> 0 Then
            ' add it to the tabs
vbwProfiler.vbwExecuteLine 8612
            index = tabVehicle.Tabs.Count + 1
vbwProfiler.vbwExecuteLine 8613
            sKey = KeyFromLong(lptr)
vbwProfiler.vbwExecuteLine 8614
            tabVehicle.Tabs.Add index, sKey, sKey

            ' graph it
vbwProfiler.vbwExecuteLine 8615
            GraphVehicle treeVehicle, 0, lptr

vbwProfiler.vbwExecuteLine 8616
            LoadVehicle = True
vbwProfiler.vbwExecuteLine 8617
            m_bLoadedFlag = True
vbwProfiler.vbwExecuteLine 8618
            m_bSavedFlag = False  ' the vehicle has not been saved yet
            ' Call m_oRecentFiles.DeleteRecentFile(sFilePath)  'todo: fix most recent file listing and uncomment this
vbwProfiler.vbwProcOut 295
vbwProfiler.vbwExecuteLine 8619
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 8620 'B
    End If
vbwProfiler.vbwExecuteLine 8621 'B
err:
vbwProfiler.vbwExecuteLine 8622
     m_bLoadedFlag = False
vbwProfiler.vbwExecuteLine 8623
    LoadVehicle = False
vbwProfiler.vbwExecuteLine 8624
    MsgBox TypeName(Me) & ":LoadVehicle() -- Error.  Failed to load vehicle."
vbwProfiler.vbwProcOut 295
vbwProfiler.vbwExecuteLine 8625
End Function
Private Sub mnuPaste_Click()
vbwProfiler.vbwProcIn 296
    Dim hTreeParent As Long
    Dim hObject As Long
    Dim f As Boolean

vbwProfiler.vbwExecuteLine 8626
    hTreeParent = treeVehicle.Selection
vbwProfiler.vbwExecuteLine 8627
    If hTreeParent <> 0 Then
vbwProfiler.vbwExecuteLine 8628
        hObject = treeVehicle.ItemData(hTreeParent)
vbwProfiler.vbwExecuteLine 8629
        If m_sClipboard <> "" Then
vbwProfiler.vbwExecuteLine 8630
            f = AddComponentsFromFile(m_sClipboard, hTreeParent, hObject)
        Else
vbwProfiler.vbwExecuteLine 8631 'B
vbwProfiler.vbwExecuteLine 8632
            MsgBox "Nothing to paste."
vbwProfiler.vbwProcOut 296
vbwProfiler.vbwExecuteLine 8633
            Exit Sub
        End If
vbwProfiler.vbwExecuteLine 8634 'B
    End If
vbwProfiler.vbwExecuteLine 8635 'B

vbwProfiler.vbwExecuteLine 8636
    If Not f Then
vbwProfiler.vbwExecuteLine 8637
        MsgBox "Node could not be pasted."
    End If
vbwProfiler.vbwExecuteLine 8638 'B
vbwProfiler.vbwProcOut 296
vbwProfiler.vbwExecuteLine 8639
End Sub
Private Sub mnuCopy_Click()
vbwProfiler.vbwProcIn 297
vbwProfiler.vbwExecuteLine 8640
    If Not copyNode(treeVehicle.Selection, False) Then
vbwProfiler.vbwExecuteLine 8641
        MsgBox "Node cannot be copied."
    End If
vbwProfiler.vbwExecuteLine 8642 'B
vbwProfiler.vbwProcOut 297
vbwProfiler.vbwExecuteLine 8643
End Sub
Private Sub mnuCopyBranch_Click()
vbwProfiler.vbwProcIn 298
vbwProfiler.vbwExecuteLine 8644
    If Not copyNode(treeVehicle.Selection, True) Then
vbwProfiler.vbwExecuteLine 8645
        MsgBox "Branch cannot be copied."
    End If
vbwProfiler.vbwExecuteLine 8646 'B
vbwProfiler.vbwProcOut 298
vbwProfiler.vbwExecuteLine 8647
End Sub
Private Sub mnuSaveComponent_Click()
    'frmSaveComponent.Show vbModal, Me
    'Set frmSaveComponent = Nothing
    ' todo: devise way to pick save file name AND wht about component categories?
    '       do we save components in same location as intrinsic? are intrinsic actually in a .PAK or ZIP?
    '       maybe use a toggle button to Show/Unshow user created components within the same listview as intrinsic?
vbwProfiler.vbwProcIn 299
vbwProfiler.vbwExecuteLine 8648
    m_oManager.saveNode treeVehicle.ItemData(treeVehicle.Selection), "c:\testsavexml.txt", True
vbwProfiler.vbwProcOut 299
vbwProfiler.vbwExecuteLine 8649
End Sub
Private Function copyNode(ByVal hTreeNode As Long, ByVal fRecurse As Boolean) As Boolean
vbwProfiler.vbwProcIn 300
    Dim sFileName As String
    Dim szBuf As String * 256
    Dim lRet As Long
    Dim sPath As String
vbwProfiler.vbwExecuteLine 8650
    Const BUFFER_SIZE = 512

vbwProfiler.vbwExecuteLine 8651
    sPath = Space$(BUFFER_SIZE)
vbwProfiler.vbwExecuteLine 8652
    sFileName = Space$(BUFFER_SIZE)
    ' - create a temp file, or overright existing?
vbwProfiler.vbwExecuteLine 8653
    lRet = GetTempPath(BUFFER_SIZE, sPath)
vbwProfiler.vbwExecuteLine 8654
    szBuf = sPath
vbwProfiler.vbwExecuteLine 8655
    lRet = GetTempFileName(szBuf, "tmp", 0, sFileName)
    ' - save node using temp filename
vbwProfiler.vbwExecuteLine 8656
    m_sClipboard = Trim$(sFileName)
vbwProfiler.vbwExecuteLine 8657
    If m_oManager.copyNode(treeVehicle.ItemData(hTreeNode), m_sClipboard, fRecurse) Then
vbwProfiler.vbwExecuteLine 8658
        Debug.Print "CLIPBOARD = " & m_sClipboard
vbwProfiler.vbwExecuteLine 8659
        copyNode = True
vbwProfiler.vbwProcOut 300
vbwProfiler.vbwExecuteLine 8660
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 8661 'B
err:
vbwProfiler.vbwExecuteLine 8662
    copyNode = False
vbwProfiler.vbwExecuteLine 8663
    m_sClipboard = ""
vbwProfiler.vbwProcOut 300
vbwProfiler.vbwExecuteLine 8664
End Function
Private Sub mnuDelete_Click()
vbwProfiler.vbwProcIn 301
vbwProfiler.vbwExecuteLine 8665
    Call RemoveComponent
vbwProfiler.vbwProcOut 301
vbwProfiler.vbwExecuteLine 8666
End Sub
Private Sub mnuRename_Click()
vbwProfiler.vbwProcIn 302
vbwProfiler.vbwExecuteLine 8667
    Call RenameComponent
vbwProfiler.vbwProcOut 302
vbwProfiler.vbwExecuteLine 8668
End Sub
Private Sub RenameComponent()
vbwProfiler.vbwProcIn 303
    Dim hNode As Long
    Dim sName As String
    Dim hSelected As Long

vbwProfiler.vbwExecuteLine 8669
    sName = InputBox("Enter new name", "Rename")
vbwProfiler.vbwExecuteLine 8670
    If IsValidFilename(sName) Then
vbwProfiler.vbwExecuteLine 8671
        hSelected = treeVehicle.Selection
vbwProfiler.vbwExecuteLine 8672
        hNode = treeVehicle.ItemData(hSelected)
vbwProfiler.vbwExecuteLine 8673
        If Not m_oManager.renameNode(hNode, sName) Then
vbwProfiler.vbwExecuteLine 8674
            MsgBox "You are not allowed to rename this node."
        Else
vbwProfiler.vbwExecuteLine 8675 'B
vbwProfiler.vbwExecuteLine 8676
            treeVehicle.ItemText(hSelected) = sName
        End If
vbwProfiler.vbwExecuteLine 8677 'B
    Else
vbwProfiler.vbwExecuteLine 8678 'B
vbwProfiler.vbwExecuteLine 8679
        MsgBox "Name contains invalid characters."
    End If
vbwProfiler.vbwExecuteLine 8680 'B
vbwProfiler.vbwProcOut 303
vbwProfiler.vbwExecuteLine 8681
End Sub
Private Sub mnuRevert_Click()
    ' todo: if we allow this, there are two options
    ' 1) we call a .Revert function of sorts from the node via oManager which reads its original name from the XML
    ' 2) we store the origianl read only name in the node itself and call oManager.Revert (hNode) as string
vbwProfiler.vbwProcIn 304
vbwProfiler.vbwProcOut 304
vbwProfiler.vbwExecuteLine 8682
End Sub
Private Sub RemoveComponent()
vbwProfiler.vbwProcIn 305
    Dim f As Boolean
    Dim hObject As Long
    Dim hSelected As Long

    ' obtain the selected tree node, retreive the pointer to the vehicle component it repesents
    ' and delete them both
vbwProfiler.vbwExecuteLine 8683
    hSelected = treeVehicle.Selection
vbwProfiler.vbwExecuteLine 8684
    If hSelected Then
vbwProfiler.vbwExecuteLine 8685
        hObject = treeVehicle.ItemData(hSelected)
vbwProfiler.vbwExecuteLine 8686
        f = m_oManager.DeleteNode(hObject)
vbwProfiler.vbwExecuteLine 8687
        If f Then
vbwProfiler.vbwExecuteLine 8688
            treeVehicle.RemoveItem (hSelected)
        Else
vbwProfiler.vbwExecuteLine 8689 'B
vbwProfiler.vbwExecuteLine 8690
            Debug.Print TypeName(Me) & ":RemoveComponent() -- Could not delete node"
        End If
vbwProfiler.vbwExecuteLine 8691 'B
    End If
vbwProfiler.vbwExecuteLine 8692 'B
vbwProfiler.vbwExecuteLine 8693
    p_bChangedFlag = True ' JAW 2000.05.07
vbwProfiler.vbwProcOut 305
vbwProfiler.vbwExecuteLine 8694
End Sub
Private Sub AddComponentsFromFile(ByRef sSourceKey As String, ByVal hNodeParent As Long, ByVal hTreeParent As Long)
vbwProfiler.vbwProcIn 306
    Dim sFilePath As String
    Dim hChild As Long
    Dim lngNodeCount As Long

vbwProfiler.vbwExecuteLine 8695
    lngNodeCount = 20 'todo: need function for calc'ing nodeCount since apparently there isnt a count property in the control
                      ' todo: Its still possible to add more nodes by simply dragging a saved file that has tons of children
                      ' onto the tree since we never count how many child nodes exist inthe saved file.  Investigate solutions.
vbwProfiler.vbwExecuteLine 8696
    If Not lngNodeCount = MAX_NODES Then
vbwProfiler.vbwExecuteLine 8697
        sFilePath = sSourceKey  ' from listview, the key is actually the full path
vbwProfiler.vbwExecuteLine 8698
        hChild = m_oManager.addNode(sFilePath, hNodeParent)
vbwProfiler.vbwExecuteLine 8699
        If hChild <> 0 Then
            ' add it to the tree
vbwProfiler.vbwExecuteLine 8700
            GraphVehicle treeVehicle, hTreeParent, hChild
        Else
vbwProfiler.vbwExecuteLine 8701 'B
            ' could not add the child to the parent object, possible reasons are
            ' its a leaf node, invalid location given the type, max nodes, etc
vbwProfiler.vbwExecuteLine 8702
            InfoPrint 1, "frmDesigner:AddComponentsFromFile() --  Could not add node to parent. Possible reasons are its a leaf node, invalid location given the type, max node cound reached?"
        End If
vbwProfiler.vbwExecuteLine 8703 'B
    Else
vbwProfiler.vbwExecuteLine 8704 'B
vbwProfiler.vbwExecuteLine 8705
        MsgBox "Node count reached.  You cannot have more than '" & MAX_NODES & "' nodes in the tree."
    End If
vbwProfiler.vbwExecuteLine 8706 'B
vbwProfiler.vbwProcOut 306
vbwProfiler.vbwExecuteLine 8707
End Sub

Private Function MoveExistingComponent(ByVal hSrc As Long, ByVal hDest As Long) As Boolean
vbwProfiler.vbwProcIn 307
vbwProfiler.vbwExecuteLine 8708
    MoveExistingComponent = m_oManager.moveNode(hSrc, hDest)
vbwProfiler.vbwProcOut 307
vbwProfiler.vbwExecuteLine 8709
End Function

Private Function getCurrentVehicle() As Long
' returns handle to cVehicle
vbwProfiler.vbwProcIn 308
    Dim oTabs As Tabs
    Dim i As Long

vbwProfiler.vbwExecuteLine 8710
    For i = 1 To tabVehicle.Tabs.Count
vbwProfiler.vbwExecuteLine 8711
        If tabVehicle.Tabs.Item(i).Selected Then
vbwProfiler.vbwExecuteLine 8712
            getCurrentVehicle = Val(tabVehicle.Tabs.Item(i).Key)
vbwProfiler.vbwProcOut 308
vbwProfiler.vbwExecuteLine 8713
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 8714 'B
vbwProfiler.vbwExecuteLine 8715
    Next
vbwProfiler.vbwExecuteLine 8716
    getCurrentVehicle = 0
vbwProfiler.vbwProcOut 308
vbwProfiler.vbwExecuteLine 8717
End Function
Private Function RemoveVehicle(ByVal rootHanlde As Long)
vbwProfiler.vbwProcIn 309
    Dim cTab As TabStrip.Tab
    Dim i As Long
    Dim sKey As String

vbwProfiler.vbwExecuteLine 8718
    If m_oManager.deleteVehicle(rootHandle) Then
vbwProfiler.vbwExecuteLine 8719
        sKey = KeyFromLong(rootHandle)
vbwProfiler.vbwExecuteLine 8720
        For Each cTab In tabVehicle.Tabs
vbwProfiler.vbwExecuteLine 8721
            If tabVehicle.Tabs.Item(i).Key = sKey Then
vbwProfiler.vbwExecuteLine 8722
                tabVehicle.Tabs.Remove (i)
vbwProfiler.vbwExecuteLine 8723
                Exit For
            End If
vbwProfiler.vbwExecuteLine 8724 'B
vbwProfiler.vbwExecuteLine 8725
            i = i + 1
vbwProfiler.vbwExecuteLine 8726
        Next
    Else
vbwProfiler.vbwExecuteLine 8727 'B
vbwProfiler.vbwExecuteLine 8728
        MsgBox TypeName(Me) & ":RemoveVehicle() -- Error: Could not delete vehicle"
    End If
vbwProfiler.vbwExecuteLine 8729 'B
vbwProfiler.vbwProcOut 309
vbwProfiler.vbwExecuteLine 8730
End Function

Private Sub SaveVehicle()
vbwProfiler.vbwProcIn 310
    Dim sTemp As String
    Dim f As Boolean
    ' If the file has already been saved with a valid fileName
vbwProfiler.vbwExecuteLine 8731
    If m_bSavedFlag Then

vbwProfiler.vbwExecuteLine 8732
        sTemp = StatusBar1.Panels(1).Text 'save the current text so we can restore it
vbwProfiler.vbwExecuteLine 8733
        StatusBar1.Panels(1).Picture = ImageList1.ListImages(8).Picture
vbwProfiler.vbwExecuteLine 8734
        StatusBar1.Panels(1).Text = "serializing vehicle data..."

        'todo: i could first check for the current active tabstrip tab
        '      and then check its tag property for m_sVehicleFile.  if filename
        '      is "" , thenits not saved yet.  Problem with m_sVehicleFile variable though
        '      is it either needs to be an array sorted via root node handles or
        '      just use the tabs and forget the var altogether
vbwProfiler.vbwExecuteLine 8735
        f = m_oManager.saveNode(treeVehicle.ItemData(treeVehicle.Root(0)), m_sVehicleFile, True)
vbwProfiler.vbwExecuteLine 8736
        Debug.Print "SAVE OPERATION " & f & " File output = " & m_sVehicleFile

vbwProfiler.vbwExecuteLine 8737
        StatusBar1.Panels(1).Picture = LoadPicture()
vbwProfiler.vbwExecuteLine 8738
        StatusBar1.Panels(1).Text = sTemp 'now restore the original text
vbwProfiler.vbwExecuteLine 8739
        Call UpdateVehicleVersionAndCopyRight
vbwProfiler.vbwExecuteLine 8740
        p_bChangedFlag = False ' JAW 2000.05.07 reset flag, all changes are now saved
        ' Display the save icon in the status bar
    Else
vbwProfiler.vbwExecuteLine 8741 'B
vbwProfiler.vbwExecuteLine 8742
        mnuSaveAs_Click
    End If
vbwProfiler.vbwExecuteLine 8743 'B
vbwProfiler.vbwProcOut 310
vbwProfiler.vbwExecuteLine 8744
End Sub

Private Sub mnuSave_Click()
vbwProfiler.vbwProcIn 311
vbwProfiler.vbwExecuteLine 8745
    Call SaveVehicle
vbwProfiler.vbwProcOut 311
vbwProfiler.vbwExecuteLine 8746
End Sub

Private Sub mnuSaveAs_Click()
    ' Code to load the SaveAs common dialog and to handle the saving
    ' of the file if the user does want to save the file
vbwProfiler.vbwProcIn 312
    Dim Cancel As Boolean
    Dim sTemp As String
    Dim oCDLG As clsCmdlg

vbwProfiler.vbwExecuteLine 8747
    On Error GoTo errorhandler
vbwProfiler.vbwExecuteLine 8748
    Cancel = False
vbwProfiler.vbwExecuteLine 8749
    Set oCDLG = New clsCmdlg
vbwProfiler.vbwExecuteLine 8750
    With oCDLG
        Dim oFile As FileSystemObject
vbwProfiler.vbwExecuteLine 8751
        Set oFile = New FileSystemObject
vbwProfiler.vbwExecuteLine 8752
        If oFile.FolderExists(Settings.VehiclesSavePath) Then
vbwProfiler.vbwExecuteLine 8753
            .InitialDir = Settings.VehiclesSavePath
        Else
vbwProfiler.vbwExecuteLine 8754 'B
vbwProfiler.vbwExecuteLine 8755
            .InitialDir = App.Path
        End If
vbwProfiler.vbwExecuteLine 8756 'B
vbwProfiler.vbwExecuteLine 8757
        If m_sVehicleFile <> NEW_VEHICLE_DEFAULT_FILENAME Then
vbwProfiler.vbwExecuteLine 8758
            .DefaultFilename = m_sVehicleFile
        Else
vbwProfiler.vbwExecuteLine 8759 'B
vbwProfiler.vbwExecuteLine 8760
            .DefaultFilename = ""
        End If
vbwProfiler.vbwExecuteLine 8761 'B
        '.DefaultExt = ".veh"
vbwProfiler.vbwExecuteLine 8762
        .Filter = OPEN_SAVE_DIALOG_FILTER
vbwProfiler.vbwExecuteLine 8763
        .CancelError = True
vbwProfiler.vbwExecuteLine 8764
        .MultiSelect = False
        '.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
vbwProfiler.vbwExecuteLine 8765
    End With

vbwProfiler.vbwExecuteLine 8766
    Cancel = oCDLG.ShowSave(Me.hwnd)
vbwProfiler.vbwExecuteLine 8767
    If Not Cancel Then
        ' A fileName was selected. Add the code to save the file here
vbwProfiler.vbwExecuteLine 8768
        m_sVehicleFile = oCDLG.cFileName(0)
        'remember this path
vbwProfiler.vbwExecuteLine 8769
        Settings.VehiclesSavePath = ExtractPathFromFile(m_sVehicleFile)
vbwProfiler.vbwExecuteLine 8770
        p_sGUID = CreateGUID 'MPJ 07/25/2000 'all "save as" vehicles are given new guid
vbwProfiler.vbwExecuteLine 8771
        m_bSavedFlag = True ' the vehicle has been saved
vbwProfiler.vbwExecuteLine 8772
        p_bChangedFlag = False ' JAW 2000.05.07 no unsaved changes now remain
vbwProfiler.vbwExecuteLine 8773
        DoEvents
vbwProfiler.vbwExecuteLine 8774
        Call SaveVehicle
vbwProfiler.vbwExecuteLine 8775
        frmDesigner.Caption = oCDLG.cFileTitle(0) & App.Title & App.Major & "." & App.Minor
         ' save the name and path of this file to our Settings UDT and our File Menu
        ' todo: Call AddRecentFile(m_sVehicleFile, oCDLG.cFileTitle(0))
    End If
vbwProfiler.vbwExecuteLine 8776 'B
vbwProfiler.vbwProcOut 312
vbwProfiler.vbwExecuteLine 8777
    Exit Sub
errorhandler:
vbwProfiler.vbwExecuteLine 8778
    InfoPrint 1, "Error in mnuSaveAs_Click:  " & CStr(err.Number) & " " & err.Description
vbwProfiler.vbwExecuteLine 8779
    Resume Next
vbwProfiler.vbwProcOut 312
vbwProfiler.vbwExecuteLine 8780
End Sub


Sub ShowCustomDropDown()
vbwProfiler.vbwProcIn 313
vbwProfiler.vbwExecuteLine 8781
    If m_oCBO Is Nothing Then
vbwProfiler.vbwExecuteLine 8782
        Set m_oCBO = New clsCompList
        ' todo: Move this string into the configuration dialog... maybe even use a hidden dialog to
    ' make this configureably by me, but not by users.  Or maybe users will want to
    ' have their own versions of this text
vbwProfiler.vbwExecuteLine 8783
        Call m_oCBO.SetFileName(App.Path & "\data\parts.txt")
    End If
vbwProfiler.vbwExecuteLine 8784 'B
vbwProfiler.vbwExecuteLine 8785
    Call m_oCBO.ShowDropDown
vbwProfiler.vbwProcOut 313
vbwProfiler.vbwExecuteLine 8786
End Sub

Private Sub cboComponents_KeyDown(KeyCode As Integer, Shift As Integer)
vbwProfiler.vbwProcIn 314
vbwProfiler.vbwExecuteLine 8787
    KeyCode = 0
vbwProfiler.vbwProcOut 314
vbwProfiler.vbwExecuteLine 8788
End Sub

Private Sub cboComponents_KeyPress(KeyAscii As Integer)
    ' dont allow the user to manually edit this box
vbwProfiler.vbwProcIn 315
vbwProfiler.vbwExecuteLine 8789
    KeyAscii = 0
vbwProfiler.vbwProcOut 315
vbwProfiler.vbwExecuteLine 8790
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
vbwProfiler.vbwProcIn 316
vbwProfiler.vbwExecuteLine 8791
    On Error Resume Next
    Static iPrevKey As Integer

vbwProfiler.vbwExecuteLine 8792
    If KeyCode = vbUpArrow Then
vbwProfiler.vbwExecuteLine 8793
        treeVehicle.AutoHScroll = True
vbwProfiler.vbwExecuteLine 8794
        treeVehicle.AutoVScroll = True
    End If
vbwProfiler.vbwExecuteLine 8795 'B
vbwProfiler.vbwExecuteLine 8796
    iPrevKey = KeyCode
vbwProfiler.vbwProcOut 316
vbwProfiler.vbwExecuteLine 8797
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'todo: can i intercept mouseup globally to make sure the drag booleans are always turned off under this condition?
vbwProfiler.vbwProcIn 317
vbwProfiler.vbwProcOut 317
vbwProfiler.vbwExecuteLine 8798
End Sub

Private Sub ConfigureMainSplitters()
vbwProfiler.vbwProcIn 318
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim r As Rect

    ' configure main splitters
vbwProfiler.vbwExecuteLine 8799
    Set oVSplitter1 = New cSplitter
vbwProfiler.vbwExecuteLine 8800
    Set oHSplitter = New cSplitter
vbwProfiler.vbwExecuteLine 8801
    Set oVSplitter2 = New cSplitter

    ' set the colors RED FOR DEBUG
vbwProfiler.vbwExecuteLine 8802
    oVSplitter1.SplitterColor = vbRed
vbwProfiler.vbwExecuteLine 8803
    oHSplitter.SplitterColor = vbRed
vbwProfiler.vbwExecuteLine 8804
    oVSplitter2.SplitterColor = vbRed

vbwProfiler.vbwExecuteLine 8805
    lngTop = tabVehicle.Top / Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 8806
    lngLeft = Settings.Splitter1
vbwProfiler.vbwExecuteLine 8807
    lngWidth = 5 ' 5 pixels
vbwProfiler.vbwExecuteLine 8808
    lngHeight = tabVehicle.Height / Screen.TwipsPerPixelY

vbwProfiler.vbwExecuteLine 8809
    With oVSplitter1
vbwProfiler.vbwExecuteLine 8810
        .SetPosition Me, "vSplitter1", lngTop, lngLeft, lngWidth, lngHeight
vbwProfiler.vbwExecuteLine 8811
        .Orientation = splitvertical
vbwProfiler.vbwExecuteLine 8812
        .SetPadding (Toolbar1.Height / Screen.TwipsPerPixelY), 5, 5, (StatusBar1.Height / Screen.TwipsPerPixelY) + 5

vbwProfiler.vbwExecuteLine 8813
        .AddControl tabVehicle, ctlTopLeft
vbwProfiler.vbwExecuteLine 8814
        .AddControl tabMain.Item(0), ctlbottomright
        ' note: the hsplitter must also be added as a "child" to the VSplitter since the Vsplitter controls its width
vbwProfiler.vbwExecuteLine 8815
        .AddControl oHSplitter, ctlbottomright
vbwProfiler.vbwExecuteLine 8816
        .AddControl tabSub.Item(1), ctlbottomright
vbwProfiler.vbwExecuteLine 8817
    End With

    ' VSplitter2 will always default to midway in tabmain
vbwProfiler.vbwExecuteLine 8818
    GetClientRect tabMain(0).hwnd, r
vbwProfiler.vbwExecuteLine 8819
    lngLeft = r.Right - r.Left

vbwProfiler.vbwExecuteLine 8820
    With oVSplitter2
vbwProfiler.vbwExecuteLine 8821
        .SetPosition tabMain.Item(0), "vSplitter2", 10, lngLeft, 5, tabMain.Item(0).Height
vbwProfiler.vbwExecuteLine 8822
        Debug.Print "VSPlitter2 Left = " & tabMain(0).Width / Screen.TwipsPerPixelX / 2
vbwProfiler.vbwExecuteLine 8823
        .Orientation = splitvertical
vbwProfiler.vbwExecuteLine 8824
        .SetPadding 30, 15, 15, 15
vbwProfiler.vbwExecuteLine 8825
    End With

vbwProfiler.vbwExecuteLine 8826
    lngTop = Settings.HSplitter '(tabSub.item(1).Top - 60) / Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 8827
    lngLeft = oVSplitter1.Left + oVSplitter1.Width
vbwProfiler.vbwExecuteLine 8828
    lngWidth = tabSub.Item(1).Width / Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 8829
    lngHeight = 5 ' pixels

vbwProfiler.vbwExecuteLine 8830
    With oHSplitter
vbwProfiler.vbwExecuteLine 8831
        .SetPosition Me, "hSplitter", lngTop, lngLeft, lngWidth, lngHeight
vbwProfiler.vbwExecuteLine 8832
        .Orientation = splithorizontal
vbwProfiler.vbwExecuteLine 8833
        .SetPadding (Toolbar1.Height / Screen.TwipsPerPixelY), 5, 5, (StatusBar1.Height / Screen.TwipsPerPixelY) + 5
vbwProfiler.vbwExecuteLine 8834
        .AddControl tabMain.Item(0), ctlTopLeft
vbwProfiler.vbwExecuteLine 8835
        .AddControl tabSub.Item(1), ctlbottomright
vbwProfiler.vbwExecuteLine 8836
    End With

vbwProfiler.vbwExecuteLine 8837
   lngTop = (tabMain(0).Top + 50) / Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 8838
   lngLeft = (tabMain(0).Left + (tabMain(0).Width / 2)) / Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 8839
   lngWidth = 5 'pixels
vbwProfiler.vbwExecuteLine 8840
   lngHeight = tabMain(0).Height / Screen.TwipsPerPixelY
vbwProfiler.vbwProcOut 318
vbwProfiler.vbwExecuteLine 8841
End Sub

Public Sub FillListViewDefaults()
    ' todo: Fill the listview with Components to start
vbwProfiler.vbwProcIn 319
vbwProfiler.vbwExecuteLine 8842
    Debug.Print "frmDesigner.FillListViewDefaults - TODO"
vbwProfiler.vbwProcOut 319
vbwProfiler.vbwExecuteLine 8843
End Sub

Private Sub mnuConfigureGVD_Click()
vbwProfiler.vbwProcIn 320
vbwProfiler.vbwExecuteLine 8844
    On Error Resume Next
vbwProfiler.vbwExecuteLine 8845
    frmConfigure.Show vbModal, frmDesigner
vbwProfiler.vbwExecuteLine 8846
    Set frmConfigure = Nothing
    'UpdateVehicle 'todo: uncomment after UpdateVehicle routine is working again
vbwProfiler.vbwProcOut 320
vbwProfiler.vbwExecuteLine 8847
End Sub

Private Sub mnuTextSlim_Click()
vbwProfiler.vbwProcIn 321
vbwProfiler.vbwExecuteLine 8848
    ExportFile "Text Slim"
vbwProfiler.vbwProcOut 321
vbwProfiler.vbwExecuteLine 8849
End Sub
Private Sub mnuText_Click()
vbwProfiler.vbwProcIn 322
vbwProfiler.vbwExecuteLine 8850
    ExportFile "Text"
vbwProfiler.vbwProcOut 322
vbwProfiler.vbwExecuteLine 8851
End Sub
Private Sub mnuHTML_classic_Click()
vbwProfiler.vbwProcIn 323
vbwProfiler.vbwExecuteLine 8852
    ExportFile "Class HTML"
vbwProfiler.vbwProcOut 323
vbwProfiler.vbwExecuteLine 8853
End Sub
Private Sub mnuHTML_Tables_Click()
vbwProfiler.vbwProcIn 324
vbwProfiler.vbwExecuteLine 8854
    ExportFile "New HTML"
vbwProfiler.vbwProcOut 324
vbwProfiler.vbwExecuteLine 8855
End Sub
Private Sub mnuUnitConversion_Click()
vbwProfiler.vbwProcIn 325
vbwProfiler.vbwExecuteLine 8856
    frmUnitConverter.Show vbModal, Me
vbwProfiler.vbwProcOut 325
vbwProfiler.vbwExecuteLine 8857
End Sub

Private Sub mnuUseSurfaceAreaTable_Click()
vbwProfiler.vbwProcIn 326
vbwProfiler.vbwExecuteLine 8858
    On Error Resume Next
vbwProfiler.vbwExecuteLine 8859
    If mnuUseSurfaceAreaTable.Checked = True Then
vbwProfiler.vbwExecuteLine 8860
        frmDesigner.Toolbar1.Buttons.Item(9).value = tbrUnpressed
vbwProfiler.vbwExecuteLine 8861
        frmDesigner.mnuUseSurfaceAreaTable.Checked = False
vbwProfiler.vbwExecuteLine 8862
        Settings.bUseSurfaceAreaTable = False
vbwProfiler.vbwExecuteLine 8863
        m_oCurrentVeh.Options.UseSurfaceAreaTable = False
    Else
vbwProfiler.vbwExecuteLine 8864 'B
vbwProfiler.vbwExecuteLine 8865
        frmDesigner.Toolbar1.Buttons.Item(9).value = tbrPressed
vbwProfiler.vbwExecuteLine 8866
        frmDesigner.mnuUseSurfaceAreaTable.Checked = True
vbwProfiler.vbwExecuteLine 8867
        Settings.bUseSurfaceAreaTable = True
vbwProfiler.vbwExecuteLine 8868
        m_oCurrentVeh.Options.UseSurfaceAreaTable = True
    End If
vbwProfiler.vbwExecuteLine 8869 'B
    'recalc all the stats
vbwProfiler.vbwExecuteLine 8870
    p_bChangedFlag = True ' JAW 2000.05.07
    'UpdateVehicle 'todo: uncomment when fixed
vbwProfiler.vbwProcOut 326
vbwProfiler.vbwExecuteLine 8871
End Sub

Private Sub mnuVersion_Click()
vbwProfiler.vbwProcIn 327
vbwProfiler.vbwExecuteLine 8872
    MsgBox "GVD.exe v" & App.Major & "." & App.Minor & "." & App.Revision
vbwProfiler.vbwProcOut 327
vbwProfiler.vbwExecuteLine 8873
End Sub

Private Sub mnuRecent_Click(index As Integer)
vbwProfiler.vbwProcIn 328
vbwProfiler.vbwExecuteLine 8874
    If OpenNewVehicle Then
vbwProfiler.vbwExecuteLine 8875
        If LoadVehicle(mnuRecent(index).Caption) Then
vbwProfiler.vbwExecuteLine 8876
            Call setGUID
        End If
vbwProfiler.vbwExecuteLine 8877 'B
    End If
vbwProfiler.vbwExecuteLine 8878 'B
vbwProfiler.vbwProcOut 328
vbwProfiler.vbwExecuteLine 8879
End Sub

Sub UpdateVehicleVersionAndCopyRight()
    'todo: this needs to be a command to oManager or somethhing..
    '      or actually, it needs to be apart of Vehicles.stats or some such
    'm_oCurrentVeh.Description.CopyrightDate = Format(Date, "mmmm, yyyy")
    'm_oCurrentVeh.Description.version = m_oCurrentVeh.Description.version + 0.01
vbwProfiler.vbwProcIn 329
vbwProfiler.vbwProcOut 329
vbwProfiler.vbwExecuteLine 8880
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' sort the report view according to the header clicked by the user
vbwProfiler.vbwProcIn 330
vbwProfiler.vbwExecuteLine 8881
    ListView1.SortKey = ColumnHeader.index - 1
vbwProfiler.vbwExecuteLine 8882
    ListView1.Sorted = True
vbwProfiler.vbwProcOut 330
vbwProfiler.vbwExecuteLine 8883
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
vbwProfiler.vbwProcIn 331
vbwProfiler.vbwExecuteLine 8884
    If KeyCode = vbKeyReturn Then
vbwProfiler.vbwExecuteLine 8885
        ListView1_DblClick
    End If
vbwProfiler.vbwExecuteLine 8886 'B
vbwProfiler.vbwProcOut 331
vbwProfiler.vbwExecuteLine 8887
End Sub

Private Sub ListView1_DblClick()
vbwProfiler.vbwProcIn 332
    Dim h As Long
vbwProfiler.vbwExecuteLine 8888
    If treeVehicle.Selection <> 0 Then
vbwProfiler.vbwExecuteLine 8889
        If Not ListView1.SelectedItem Is Nothing Then
vbwProfiler.vbwExecuteLine 8890
            AddComponentsFromFile ListView1.SelectedItem.Key, treeVehicle.ItemData(treeVehicle.Selection), treeVehicle.Selection
        End If
vbwProfiler.vbwExecuteLine 8891 'B
    End If
vbwProfiler.vbwExecuteLine 8892 'B
vbwProfiler.vbwProcOut 332
vbwProfiler.vbwExecuteLine 8893
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 333
vbwProfiler.vbwExecuteLine 8894
    If Button = vbLeftButton Then ' Signal a Drag operation.
vbwProfiler.vbwExecuteLine 8895
        If Not ListView1.SelectedItem Is Nothing Then ' make sure there is an item selected
            'GetDataTypeAndIcons ListView1.SelectedItem.Index Oct.01.2002 Obsolete since switching to XML Based Components
vbwProfiler.vbwExecuteLine 8896
            ListView1.DragIcon = ListView1.SelectedItem.CreateDragImage
vbwProfiler.vbwExecuteLine 8897
            ListView1.Drag vbBeginDrag ' Drag operation.
        End If
vbwProfiler.vbwExecuteLine 8898 'B
    Else
vbwProfiler.vbwExecuteLine 8899 'B
vbwProfiler.vbwExecuteLine 8900
        ListView1.MousePointer = ccCustom
    End If
vbwProfiler.vbwExecuteLine 8901 'B
vbwProfiler.vbwProcOut 333
vbwProfiler.vbwExecuteLine 8902
End Sub

Private Sub lstPropulsionSystems_ItemCheck(Item As Integer)
vbwProfiler.vbwProcIn 334
    Dim sCurrent As String
    ' if we are checkmarking items because we are restoring the checklist and the user is not actually
    ' clicking on items, then we dont want to call the user Select and DeSelect routines.
vbwProfiler.vbwExecuteLine 8903
    If lstPropulsionSystems.Tag = CHECKLIST_STATE_RESTORE Then
vbwProfiler.vbwProcOut 334
vbwProfiler.vbwExecuteLine 8904
         Exit Sub
    End If
vbwProfiler.vbwExecuteLine 8905 'B

vbwProfiler.vbwExecuteLine 8906
    If lstPropulsionSystems.Selected(Item) Then
vbwProfiler.vbwExecuteLine 8907
        Call PropulsionSelect(sCurrent, CLng(Item))
    Else
vbwProfiler.vbwExecuteLine 8908 'B
vbwProfiler.vbwExecuteLine 8909
        Call PropulsionDeSelect(sCurrent, CLng(Item))
    End If
vbwProfiler.vbwExecuteLine 8910 'B
vbwProfiler.vbwProcOut 334
vbwProfiler.vbwExecuteLine 8911
End Sub

Private Sub lstviewLinks_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 335
vbwProfiler.vbwExecuteLine 8912
    lstviewLinks.SelectedItem = lstviewLinks.HitTest(x, y)
vbwProfiler.vbwExecuteLine 8913
    Set m_oLstviewLinkDragItem = lstviewLinks.SelectedItem

vbwProfiler.vbwExecuteLine 8914
    If Button = vbLeftButton Then
vbwProfiler.vbwExecuteLine 8915
        If Not m_oLstviewLinkDragItem Is Nothing Then
vbwProfiler.vbwExecuteLine 8916
            lstviewLinks.DragIcon = ImageList1.ListImages(2).ExtractIcon ' use extractIcon and not .Picture because the image list is not associated with this listview
vbwProfiler.vbwExecuteLine 8917
            lstviewLinks.Drag vbEndDrag
vbwProfiler.vbwExecuteLine 8918
            lstviewLinks.Drag vbBeginDrag

vbwProfiler.vbwExecuteLine 8919
            m_bTreeLinksDrag = True
        End If
vbwProfiler.vbwExecuteLine 8920 'B
    End If
vbwProfiler.vbwExecuteLine 8921 'B
vbwProfiler.vbwProcOut 335
vbwProfiler.vbwExecuteLine 8922
End Sub

Private Sub lstviewLinks_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 336
vbwProfiler.vbwExecuteLine 8923
    Set m_oLstviewLinkDragItem = Nothing
vbwProfiler.vbwExecuteLine 8924
    m_bTreeLinksDrag = False
vbwProfiler.vbwProcOut 336
vbwProfiler.vbwExecuteLine 8925
End Sub

Private Sub m_oCBO_Click(ByRef s As String)
vbwProfiler.vbwProcIn 337
vbwProfiler.vbwExecuteLine 8926
    cboComponents.Text = s
vbwProfiler.vbwExecuteLine 8927
    Call LoadListView(s)
vbwProfiler.vbwExecuteLine 8928
    tabMain(0).Tab = 0
vbwProfiler.vbwProcOut 337
vbwProfiler.vbwExecuteLine 8929
End Sub

Private Sub mnuAddPerformanceProfile_Click(index As Integer)
vbwProfiler.vbwProcIn 338
    Dim sKey As String
    Dim lngDatatype As Long
    Dim sProfileName As String

'    sKey = GetNextKey
'
'    'todo: this is temporary setting of tag on lstperformanceprofiles to try and
'    ' establish the "Current" performance profile
'    m_colVehicles(1).m_oCurrentVeh.ActiveCheckList = sKey
'    m_colVehicles(1).m_oCurrentVeh.ActiveCheckListType = PERFORMANCE_CHECKLIST

vbwProfiler.vbwExecuteLine 8930
    sProfileName = InputBox("Enter a name for this performance profile:", "New Performance Profile")
vbwProfiler.vbwExecuteLine 8931
    If Not IsValidFilename(sProfileName) Then
vbwProfiler.vbwProcOut 338
vbwProfiler.vbwExecuteLine 8932
        Exit Sub
    End If
vbwProfiler.vbwExecuteLine 8933 'B

vbwProfiler.vbwExecuteLine 8934
    Select Case index

'vbwLine 8935:        Case 0 ' wheeled
        Case IIf(vbwProfiler.vbwExecuteLine(8935), VBWPROFILER_EMPTY, _
        0 )' wheeled
vbwProfiler.vbwExecuteLine 8936
                lngDatatype = PERFORMANCEWHEEL
'vbwLine 8937:        Case 1 ' tracks
        Case IIf(vbwProfiler.vbwExecuteLine(8937), VBWPROFILER_EMPTY, _
        1 )' tracks
vbwProfiler.vbwExecuteLine 8938
                lngDatatype = PERFORMANCETRACK
'vbwLine 8939:        Case 2 ' skids
        Case IIf(vbwProfiler.vbwExecuteLine(8939), VBWPROFILER_EMPTY, _
        2 )' skids
vbwProfiler.vbwExecuteLine 8940
                lngDatatype = PERFORMANCESKID
'vbwLine 8941:        Case 3 ' legs
        Case IIf(vbwProfiler.vbwExecuteLine(8941), VBWPROFILER_EMPTY, _
        3 )' legs
vbwProfiler.vbwExecuteLine 8942
                lngDatatype = PERFORMANCELEG
'vbwLine 8943:        Case 4 'flexibody
        Case IIf(vbwProfiler.vbwExecuteLine(8943), VBWPROFILER_EMPTY, _
        4 )'flexibody
vbwProfiler.vbwExecuteLine 8944
                lngDatatype = PERFORMANCEFLEX
'vbwLine 8945:        Case 5 'air
        Case IIf(vbwProfiler.vbwExecuteLine(8945), VBWPROFILER_EMPTY, _
        5 )'air
vbwProfiler.vbwExecuteLine 8946
                lngDatatype = PERFORMANCEAIR
'vbwLine 8947:        Case 6 ' hover
        Case IIf(vbwProfiler.vbwExecuteLine(8947), VBWPROFILER_EMPTY, _
        6 )' hover
vbwProfiler.vbwExecuteLine 8948
                lngDatatype = PERFORMANCEHOVER
'vbwLine 8949:        Case 7 ' mag-lev
        Case IIf(vbwProfiler.vbwExecuteLine(8949), VBWPROFILER_EMPTY, _
        7 )' mag-lev
vbwProfiler.vbwExecuteLine 8950
                lngDatatype = PERFORMANCEMAGLEV
'vbwLine 8951:        Case 8 ' water
        Case IIf(vbwProfiler.vbwExecuteLine(8951), VBWPROFILER_EMPTY, _
        8 )' water
vbwProfiler.vbwExecuteLine 8952
                lngDatatype = PERFORMANCEWATER
'vbwLine 8953:        Case 9 ' submerged
        Case IIf(vbwProfiler.vbwExecuteLine(8953), VBWPROFILER_EMPTY, _
        9 )' submerged
vbwProfiler.vbwExecuteLine 8954
                lngDatatype = PERFORMANCESUB
'vbwLine 8955:        Case 10 ' space
        Case IIf(vbwProfiler.vbwExecuteLine(8955), VBWPROFILER_EMPTY, _
        10 )' space
vbwProfiler.vbwExecuteLine 8956
                lngDatatype = PERFORMANCESPACE
    End Select
vbwProfiler.vbwExecuteLine 8957 'B

'    m_oCurrentVeh.AddPerformanceProfile lngDatatype, sKey, PERFORMANCE_KEY, 1, sProfileName
'    treeVehicle.Nodes.Add PERFORMANCE_KEY, tvwChild, sKey, sProfileName, 1
'    addnewchildnode performance_key, "settings.ico",
'    treeVehicle.Nodes(sKey).Tag = lngDatatype
vbwProfiler.vbwProcOut 338
vbwProfiler.vbwExecuteLine 8958
End Sub

Private Sub mnuClearMessages_Click()
vbwProfiler.vbwProcIn 339
vbwProfiler.vbwExecuteLine 8959
    txtInfo = ""
vbwProfiler.vbwProcOut 339
vbwProfiler.vbwExecuteLine 8960
End Sub

Private Sub mnuDesignCheck_Click()
vbwProfiler.vbwProcIn 340
vbwProfiler.vbwExecuteLine 8961
    frmDesignCheck.Show vbModal, Me
vbwProfiler.vbwExecuteLine 8962
    Set frmDesignCheck = Nothing
vbwProfiler.vbwProcOut 340
vbwProfiler.vbwExecuteLine 8963
End Sub

Private Sub mnuExit_Click()
vbwProfiler.vbwProcIn 341
vbwProfiler.vbwExecuteLine 8964
    Unload Me
vbwProfiler.vbwProcOut 341
vbwProfiler.vbwExecuteLine 8965
End Sub

Private Sub mnuNotes_Click()
vbwProfiler.vbwProcIn 342
vbwProfiler.vbwExecuteLine 8966
    frmNotes.Show vbModal, frmDesigner
vbwProfiler.vbwExecuteLine 8967
    Set frmNotes = Nothing
vbwProfiler.vbwProcOut 342
vbwProfiler.vbwExecuteLine 8968
End Sub

Private Sub mnuPowerCreateNew_Click()
vbwProfiler.vbwProcIn 343
    Dim sProfileName As String
    Dim sKey As String
    Dim oNode As cINode
    Dim hComponent As Long
    Dim oVehicle As cVehicle
    Dim lptr As Long

    ' this Sub will simulate the creation of NEW Power Profile for testing purposes
vbwProfiler.vbwExecuteLine 8969
    sProfileName = InputBox("Enter a name for this power configuration:", "New Power Configuration")
vbwProfiler.vbwExecuteLine 8970
    If IsValidFilename(sProfileName) Then
        ' create a test power system profile
vbwProfiler.vbwExecuteLine 8971
        lptr = getCurrentVehicle
vbwProfiler.vbwExecuteLine 8972
        CopyMemory oNode, lptr, 4

vbwProfiler.vbwExecuteLine 8973
        Set oVehicle = oNode
        'todo: umm... or dont we call the factory to create this?
       ' oVehicle.addProfile , sKey, POWER_PROFILE, sProfileName
vbwProfiler.vbwExecuteLine 8974
        GraphVehicle treeVehicle, lptr, hComponent

    End If
vbwProfiler.vbwExecuteLine 8975 'B
vbwProfiler.vbwProcOut 343
vbwProfiler.vbwExecuteLine 8976
End Sub

Private Sub mnuFuelCreateNew_Click()
vbwProfiler.vbwProcIn 344
    Dim sProfileName As String
    Dim sKey As String

vbwProfiler.vbwExecuteLine 8977
    sProfileName = InputBox("Enter a name for this fuel link configuration:", "New Fuel Configuration")
vbwProfiler.vbwExecuteLine 8978
    If IsValidFilename(sProfileName) Then
vbwProfiler.vbwExecuteLine 8979
        sKey = GetNextKey
vbwProfiler.vbwExecuteLine 8980
        m_oCurrentVeh.addProfile sKey, FUEL_PROFILE, sProfileName
vbwProfiler.vbwExecuteLine 8981
        treeVehicle.Nodes.Add FUELSYSTEMS_KEY, tvwChild, sKey, sProfileName, 1
vbwProfiler.vbwExecuteLine 8982
        treeVehicle.Nodes(sKey).Tag = FUELSYSTEMS_NODE
    End If
vbwProfiler.vbwExecuteLine 8983 'B
vbwProfiler.vbwProcOut 344
vbwProfiler.vbwExecuteLine 8984
End Sub

Private Sub mnuWeaponCreateNew_Click()
vbwProfiler.vbwProcIn 345
    Dim sWeaponLinkName As String
    Dim sKey As String

vbwProfiler.vbwExecuteLine 8985
    sWeaponLinkName = InputBox("Enter a name for this weapon link:", "New Weapon Link")
vbwProfiler.vbwExecuteLine 8986
    If IsValidFilename(sWeaponLinkName) Then
vbwProfiler.vbwExecuteLine 8987
        sKey = GetNextKey

vbwProfiler.vbwExecuteLine 8988
        m_oCurrentVeh.addweaponlink sKey, sWeaponLinkName
vbwProfiler.vbwExecuteLine 8989
        treeVehicle.Nodes.Add WEAPON_LINKS_KEY, tvwChild, sKey, sWeaponLinkName, 1
vbwProfiler.vbwExecuteLine 8990
        treeVehicle.Nodes(sKey).Tag = WEAPON_LINKS_NODE
    End If
vbwProfiler.vbwExecuteLine 8991 'B
vbwProfiler.vbwProcOut 345
vbwProfiler.vbwExecuteLine 8992
End Sub

Private Sub mnuPrint_Click()
vbwProfiler.vbwProcIn 346
vbwProfiler.vbwExecuteLine 8993
    Call PrintRTF(rtbOutput, 700, 700, 700, 700)
vbwProfiler.vbwProcOut 346
vbwProfiler.vbwExecuteLine 8994
End Sub

Private Sub mnuPrintSetup_Click()
vbwProfiler.vbwProcIn 347
vbwProfiler.vbwExecuteLine 8995
    If m_oCmdlg Is Nothing Then
vbwProfiler.vbwExecuteLine 8996
        Set m_oCmdlg = New clsCmdlg
    End If
vbwProfiler.vbwExecuteLine 8997 'B
vbwProfiler.vbwExecuteLine 8998
    m_oCmdlg.ShowPageSetup (Me.hwnd)   'Show Page Setup dialog
vbwProfiler.vbwProcOut 347
vbwProfiler.vbwExecuteLine 8999
End Sub

Private Sub mnuPublish_Click()
   ' Call Publish
vbwProfiler.vbwProcIn 348
vbwProfiler.vbwProcOut 348
vbwProfiler.vbwExecuteLine 9000
End Sub

Private Sub mnuAbout_Click()
vbwProfiler.vbwProcIn 349
vbwProfiler.vbwExecuteLine 9001
    Load frmSplash
vbwProfiler.vbwExecuteLine 9002
    frmSplash.Show vbModal, frmDesigner
vbwProfiler.vbwExecuteLine 9003
    Set frmSplash = Nothing
vbwProfiler.vbwProcOut 349
vbwProfiler.vbwExecuteLine 9004
End Sub

Private Sub mnuRegister_click()
vbwProfiler.vbwProcIn 350
vbwProfiler.vbwExecuteLine 9005
    Load frmCredits
vbwProfiler.vbwExecuteLine 9006
    frmCredits.Tag = "register"
vbwProfiler.vbwExecuteLine 9007
    frmCredits.Show vbModal, frmDesigner
vbwProfiler.vbwExecuteLine 9008
    Set frmCredits = Nothing
vbwProfiler.vbwProcOut 350
vbwProfiler.vbwExecuteLine 9009
End Sub

Private Sub PLC1_BeforePropertyChanged(ByVal index As Long, Cancel As Boolean)
vbwProfiler.vbwProcIn 351
    Dim i As Long
    Dim sCharacter As String
    Dim NewString As String
    Dim iData As Integer
vbwProfiler.vbwExecuteLine 9010
    iData = PLC1.ItemData(index)

vbwProfiler.vbwExecuteLine 9011
    If iData = wdText Then
        'check for use of reserved characters
vbwProfiler.vbwExecuteLine 9012
        NewString = PLC1.value(index)
vbwProfiler.vbwExecuteLine 9013
        If Not IsValidFilename(NewString) Then
vbwProfiler.vbwExecuteLine 9014
            Cancel = True
        End If
vbwProfiler.vbwExecuteLine 9015 'B
    ' check to see that a Integer value has been added for any Number
'vbwLine 9016:    ElseIf iData = wdNumber Then
    ElseIf vbwProfiler.vbwExecuteLine(9016) Or iData = wdNumber Then
        'check for numbers less than 0
vbwProfiler.vbwExecuteLine 9017
        If PLC1.value(index) < 0 Then
vbwProfiler.vbwExecuteLine 9018
            InfoPrint 1, "This field does not accept negative numbers.  Please use positive numbers only." 'todo: need constant for this
vbwProfiler.vbwExecuteLine 9019
            Cancel = True
        End If
vbwProfiler.vbwExecuteLine 9020 'B

        'check to make sure the Quantity field has at least a 1 value and no more than 1,000
vbwProfiler.vbwExecuteLine 9021
        If PLC1.DescriptionString(index) = "Quantity" Then ' todo Need constant
vbwProfiler.vbwExecuteLine 9022
            If (PLC1.value(index) < 1) Or (PLC1.value(index) > 1000) Then 'todo: need constant for 1000
vbwProfiler.vbwExecuteLine 9023
                InfoPrint 1, "The Quantity field must be an integer value from 1 to 1,000." 'todo: need constant
vbwProfiler.vbwExecuteLine 9024
                Cancel = True
            End If
vbwProfiler.vbwExecuteLine 9025 'B
        End If
vbwProfiler.vbwExecuteLine 9026 'B
'vbwLine 9027:    ElseIf iData = wdDouble Then
    ElseIf vbwProfiler.vbwExecuteLine(9027) Or iData = wdDouble Then
        'make sure the actual user enter-able number is no greater than a Single
        'which is 3.402823E+38 and no less than .001 and NO NEGATIVE numbers at all
        'actually TODO i might want to use the square root of a single as max to keep things sane?
        Dim num As Single
vbwProfiler.vbwExecuteLine 9028
        num = PLC1.value(index)
vbwProfiler.vbwExecuteLine 9029
        If PLC1.DescriptionString(index) = "EmptySpace" Then
vbwProfiler.vbwExecuteLine 9030
            If (num > 1E+20) Or (num < 0) Then
vbwProfiler.vbwExecuteLine 9031
                InfoPrint 1, "Allowed numeric range is 0 to 1.0E20"
vbwProfiler.vbwExecuteLine 9032
                Cancel = True
            End If
vbwProfiler.vbwExecuteLine 9033 'B
'vbwLine 9034:        ElseIf PLC1.DescriptionString(index) = "UserWeight" Then 'custom component uses these and should allow 0
        ElseIf vbwProfiler.vbwExecuteLine(9034) Or PLC1.DescriptionString(index) = "UserWeight" Then 'custom component uses these and should allow 0
        'TODO: Need constants for these
'vbwLine 9035:        ElseIf PLC1.DescriptionString(index) = "UserCost" Then 'custom component uses these and should allow 0
        ElseIf vbwProfiler.vbwExecuteLine(9035) Or PLC1.DescriptionString(index) = "UserCost" Then 'custom component uses these and should allow 0
'vbwLine 9036:        ElseIf PLC1.DescriptionString(index) = "UserVolume" Then 'custom component uses these and should allow 0
        ElseIf vbwProfiler.vbwExecuteLine(9036) Or PLC1.DescriptionString(index) = "UserVolume" Then 'custom component uses these and should allow 0

        'note: Commenting this out... testing out using range checking within the DEF file via "userinputurange#" and "userinputlrange#"
        ' TODO:, i need to move all the code for callbyname in the plc1_peropertychanged into here so thta i can
        ' test if we need to Cancel = true.  Since its not til callbyname attempts to change the value that I can
        ' find out if the bounds are valid and the callbyname successfully modifies the value.
        'ElseIf (num > 1E+20) Or (num < 0.001) Then
        '    InfoPrint 1, "Allowed numeric range is 0.001 to 1.0E20" 'todo: need constant for this error message
        '    Cancel = True
        End If
vbwProfiler.vbwExecuteLine 9037 'B
    End If
vbwProfiler.vbwExecuteLine 9038 'B
vbwProfiler.vbwProcOut 351
vbwProfiler.vbwExecuteLine 9039
End Sub

Private Sub PLC1_PropertyChanged(ByVal index As Long)
vbwProfiler.vbwProcIn 352
vbwProfiler.vbwExecuteLine 9040
   Call modProperties.PropertyChanged(index)
vbwProfiler.vbwProcOut 352
vbwProfiler.vbwExecuteLine 9041
End Sub

Private Sub tabMain_Click(index As Integer, PreviousTab As Integer)
vbwProfiler.vbwProcIn 353
    Dim sKey As String


vbwProfiler.vbwExecuteLine 9042
    Select Case tabMain(0).Tab
'vbwLine 9043:        Case 0
        Case IIf(vbwProfiler.vbwExecuteLine(9043), VBWPROFILER_EMPTY, _
        0)
            'SetViewMode component_view
            ' call code to populate relevant controls
            ' add controls splitter
vbwProfiler.vbwExecuteLine 9044
                oVSplitter2.RemoveAllControls
vbwProfiler.vbwExecuteLine 9045
                oVSplitter2.AddControl ListView1, ctlTopLeft
vbwProfiler.vbwExecuteLine 9046
                oVSplitter2.AddControl PLC1, ctlbottomright

vbwProfiler.vbwExecuteLine 9047
                PLC1.Visible = True
vbwProfiler.vbwExecuteLine 9048
                ListView1.Visible = True
vbwProfiler.vbwExecuteLine 9049
                lstPropulsionSystems.Visible = False
vbwProfiler.vbwExecuteLine 9050
                treeLinks.Visible = False
vbwProfiler.vbwExecuteLine 9051
                lstviewLinks.Visible = False
vbwProfiler.vbwExecuteLine 9052
                rtbOutput.Visible = False
vbwProfiler.vbwExecuteLine 9053
                picVehicleImage.Visible = False

'vbwLine 9054:        Case 1 ' performance OR weapon link
        Case IIf(vbwProfiler.vbwExecuteLine(9054), VBWPROFILER_EMPTY, _
        1 )' performance OR weapon link

'                ' todo: i need to just use globals for tracking the active lists and active profiles
'                '        these are gui things and have no right being in the vehicles dll
'                If p_ActiveNode.Key <> m_oCurrentVeh.ActiveCheckList Then
'                    ' then the user clicked the tabstrip directly and is
'                    ' switching modes.
'                    If m_oCurrentVeh.ActiveCheckList <> "" Then
'                        ' if no performance or weapon link created, this will be empty
'                        treeVehicle.Nodes.item(m_oCurrentVeh.ActiveCheckList).Selected = True
'                        SetActiveNode
'                    End If
'                End If

vbwProfiler.vbwExecuteLine 9055
                oVSplitter2.RemoveAllControls
vbwProfiler.vbwExecuteLine 9056
                oVSplitter2.AddControl lstPropulsionSystems, ctlTopLeft
vbwProfiler.vbwExecuteLine 9057
                oVSplitter2.AddControl PLC1, ctlbottomright

vbwProfiler.vbwExecuteLine 9058
                PLC1.Visible = True
vbwProfiler.vbwExecuteLine 9059
                ListView1.Visible = False
vbwProfiler.vbwExecuteLine 9060
                lstPropulsionSystems.Visible = True
vbwProfiler.vbwExecuteLine 9061
                treeLinks.Visible = False
vbwProfiler.vbwExecuteLine 9062
                lstviewLinks.Visible = False
vbwProfiler.vbwExecuteLine 9063
                rtbOutput.Visible = False
vbwProfiler.vbwExecuteLine 9064
                picVehicleImage.Visible = False


'vbwLine 9065:        Case 2 ' power or fuel links
        Case IIf(vbwProfiler.vbwExecuteLine(9065), VBWPROFILER_EMPTY, _
        2 )' power or fuel links

'                If p_ActiveNode.Key <> m_oCurrentVeh.ActiveProfile Then
'                    ' then the user clicked the tabstrip directly and is
'                    ' switching modes.
'                    If m_oCurrentVeh.ActiveProfile <> "" Then
'                        ' if no performance or weapon link created, this will be empty
'                        treeVehicle.Nodes.item(m_oCurrentVeh.ActiveProfile).Selected = True
'                        SetActiveNode
'                    End If
'                End If

vbwProfiler.vbwExecuteLine 9066
                oVSplitter2.RemoveAllControls
vbwProfiler.vbwExecuteLine 9067
                oVSplitter2.AddControl treeLinks, ctlTopLeft
vbwProfiler.vbwExecuteLine 9068
                oVSplitter2.AddControl lstviewLinks, ctlbottomright


vbwProfiler.vbwExecuteLine 9069
                treeLinks.Visible = True
vbwProfiler.vbwExecuteLine 9070
                DoEvents
vbwProfiler.vbwExecuteLine 9071
                lstviewLinks.Visible = True '<-- this is the one leaving the little crap image before its resized
vbwProfiler.vbwExecuteLine 9072
                PLC1.Visible = False
vbwProfiler.vbwExecuteLine 9073
                ListView1.Visible = False
vbwProfiler.vbwExecuteLine 9074
                lstPropulsionSystems.Visible = False
vbwProfiler.vbwExecuteLine 9075
                rtbOutput.Visible = False

vbwProfiler.vbwExecuteLine 9076
                picVehicleImage.Visible = False


'vbwLine 9077:        Case 3
        Case IIf(vbwProfiler.vbwExecuteLine(9077), VBWPROFILER_EMPTY, _
        3)
            'SetViewMode output_view
vbwProfiler.vbwExecuteLine 9078
            oVSplitter2.RemoveAllControls
                'oVSplitter2.Visible = False
vbwProfiler.vbwExecuteLine 9079
                PLC1.Visible = False
vbwProfiler.vbwExecuteLine 9080
                ListView1.Visible = False
vbwProfiler.vbwExecuteLine 9081
                lstPropulsionSystems.Visible = False
vbwProfiler.vbwExecuteLine 9082
                treeLinks.Visible = False
vbwProfiler.vbwExecuteLine 9083
                lstviewLinks.Visible = False
vbwProfiler.vbwExecuteLine 9084
                rtbOutput.Visible = True           '<-- Only the output RTF is visible
vbwProfiler.vbwExecuteLine 9085
                rtbOutput.ZOrder
vbwProfiler.vbwExecuteLine 9086
                picVehicleImage.Visible = False


                ' test print oupt for power systems
                #If DEBUG_MODE Then
                    Dim sTemp As String

                    'get the Power plant info
vbwProfiler.vbwExecuteLine 9087
                    sTemp = createGURPSText("Text")
vbwProfiler.vbwExecuteLine 9088
                    rtbOutput.Text = sTemp
                #End If
'vbwLine 9089:        Case 4
        Case IIf(vbwProfiler.vbwExecuteLine(9089), VBWPROFILER_EMPTY, _
        4)
            'SetViewMode image_view
vbwProfiler.vbwExecuteLine 9090
            oVSplitter2.RemoveAllControls

                'oVSplitter2.Visible = False
vbwProfiler.vbwExecuteLine 9091
                PLC1.Visible = False
vbwProfiler.vbwExecuteLine 9092
                ListView1.Visible = False
vbwProfiler.vbwExecuteLine 9093
                lstPropulsionSystems.Visible = False
vbwProfiler.vbwExecuteLine 9094
                treeLinks.Visible = False
vbwProfiler.vbwExecuteLine 9095
                lstviewLinks.Visible = False
vbwProfiler.vbwExecuteLine 9096
                rtbOutput.Visible = False
vbwProfiler.vbwExecuteLine 9097
                picVehicleImage.Visible = True              '<-- Only the image picturebox is visible
vbwProfiler.vbwExecuteLine 9098
                picVehicleImage.ZOrder

    End Select
vbwProfiler.vbwExecuteLine 9099 'B

vbwProfiler.vbwExecuteLine 9100
    Call TabStrip_Resize
vbwProfiler.vbwExecuteLine 9101
    Call Form_Resize
vbwProfiler.vbwProcOut 353
vbwProfiler.vbwExecuteLine 9102
End Sub

Private Sub tabSub_Click(index As Integer, PreviousTab As Integer)
vbwProfiler.vbwProcIn 354
vbwProfiler.vbwExecuteLine 9103
    Select Case tabSub(1).Tab
'vbwLine 9104:        Case 0
        Case IIf(vbwProfiler.vbwExecuteLine(9104), VBWPROFILER_EMPTY, _
        0)
vbwProfiler.vbwExecuteLine 9105
            lstStats.Visible = False
vbwProfiler.vbwExecuteLine 9106
            txtInfo.Visible = True

'vbwLine 9107:        Case 1
        Case IIf(vbwProfiler.vbwExecuteLine(9107), VBWPROFILER_EMPTY, _
        1)
vbwProfiler.vbwExecuteLine 9108
            lstStats.Visible = True
vbwProfiler.vbwExecuteLine 9109
            txtInfo.Visible = False
    End Select
vbwProfiler.vbwExecuteLine 9110 'B
vbwProfiler.vbwProcOut 354
vbwProfiler.vbwExecuteLine 9111
End Sub

' todo: currently this sub cant be private since ShowProps in modProperties calls it.
' may want to move this code to a seperate module and make it public
Public Sub SetViewMode(eModeType As VIEW_MODE)
vbwProfiler.vbwProcIn 355
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long

    ' configure splitters and visible controls based on the selected mode
vbwProfiler.vbwExecuteLine 9112
    Select Case eModeType
'vbwLine 9113:        Case component_view
        Case IIf(vbwProfiler.vbwExecuteLine(9113), VBWPROFILER_EMPTY, _
        component_view)
vbwProfiler.vbwExecuteLine 9114
            tabMain(0).Tab = 0

'vbwLine 9115:        Case image_view
        Case IIf(vbwProfiler.vbwExecuteLine(9115), VBWPROFILER_EMPTY, _
        image_view)
vbwProfiler.vbwExecuteLine 9116
            tabMain(0).Tab = 4

'vbwLine 9117:        Case output_view
        Case IIf(vbwProfiler.vbwExecuteLine(9117), VBWPROFILER_EMPTY, _
        output_view)
vbwProfiler.vbwExecuteLine 9118
            tabMain(0).Tab = 3

        ' Weapon or Performance Profiles
'vbwLine 9119:        Case performance_view
        Case IIf(vbwProfiler.vbwExecuteLine(9119), VBWPROFILER_EMPTY, _
        performance_view)
vbwProfiler.vbwExecuteLine 9120
            tabMain(0).Tab = 1

        ' Power or Fuel Links
'vbwLine 9121:        Case links_view
        Case IIf(vbwProfiler.vbwExecuteLine(9121), VBWPROFILER_EMPTY, _
        links_view)
vbwProfiler.vbwExecuteLine 9122
            tabMain(0).Tab = 2
    End Select
vbwProfiler.vbwExecuteLine 9123 'B

vbwProfiler.vbwProcOut 355
vbwProfiler.vbwExecuteLine 9124
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
vbwProfiler.vbwProcIn 356
vbwProfiler.vbwExecuteLine 9125
    Select Case Button.index
'vbwLine 9126:        Case 1 ' New Vehicle
        Case IIf(vbwProfiler.vbwExecuteLine(9126), VBWPROFILER_EMPTY, _
        1 )' New Vehicle
vbwProfiler.vbwExecuteLine 9127
            mnuNewVehicle_Click

'vbwLine 9128:        Case 2 ' Open
        Case IIf(vbwProfiler.vbwExecuteLine(9128), VBWPROFILER_EMPTY, _
        2 )' Open
vbwProfiler.vbwExecuteLine 9129
            mnuOpen_Click
'vbwLine 9130:        Case 3 ' Save
        Case IIf(vbwProfiler.vbwExecuteLine(9130), VBWPROFILER_EMPTY, _
        3 )' Save
vbwProfiler.vbwExecuteLine 9131
            SaveVehicle
'vbwLine 9132:        Case 5 ' Print
        Case IIf(vbwProfiler.vbwExecuteLine(9132), VBWPROFILER_EMPTY, _
        5 )' Print
vbwProfiler.vbwExecuteLine 9133
            ExportFile "Text"


'vbwLine 9134:        Case 7 'Use surface Area formula
        Case IIf(vbwProfiler.vbwExecuteLine(9134), VBWPROFILER_EMPTY, _
        7 )'Use surface Area formula
vbwProfiler.vbwExecuteLine 9135
            mnuUseSurfaceAreaTable_Click
'vbwLine 9136:        Case 9 'publish vehicle
        Case IIf(vbwProfiler.vbwExecuteLine(9136), VBWPROFILER_EMPTY, _
        9 )'publish vehicle
vbwProfiler.vbwExecuteLine 9137
            mnuPublish_Click
    End Select
vbwProfiler.vbwExecuteLine 9138 'B
vbwProfiler.vbwProcOut 356
vbwProfiler.vbwExecuteLine 9139
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
vbwProfiler.vbwProcIn 357

vbwProfiler.vbwExecuteLine 9140
    Select Case ButtonMenu.Parent.index
'vbwLine 9141:        Case 1
        Case IIf(vbwProfiler.vbwExecuteLine(9141), VBWPROFILER_EMPTY, _
        1)
vbwProfiler.vbwExecuteLine 9142
            Select Case ButtonMenu.index
'vbwLine 9143:                Case 1
                Case IIf(vbwProfiler.vbwExecuteLine(9143), VBWPROFILER_EMPTY, _
        1)
vbwProfiler.vbwExecuteLine 9144
                    mnuNewVehicle_Click
'vbwLine 9145:                Case 2
                Case IIf(vbwProfiler.vbwExecuteLine(9145), VBWPROFILER_EMPTY, _
        2)
                    'mnuBattleSuitFF_Click 'todo: these are callng paths to load files which contain this stuff already
'vbwLine 9146:                Case 3
                Case IIf(vbwProfiler.vbwExecuteLine(9146), VBWPROFILER_EMPTY, _
        3)
                    'mnuBattleSuitBody_Click
'vbwLine 9147:                Case 4
                Case IIf(vbwProfiler.vbwExecuteLine(9147), VBWPROFILER_EMPTY, _
        4)
                    'mnuBattleSuitTurret_Click
            End Select
vbwProfiler.vbwExecuteLine 9148 'B
'vbwLine 9149:        Case 5
        Case IIf(vbwProfiler.vbwExecuteLine(9149), VBWPROFILER_EMPTY, _
        5)
vbwProfiler.vbwExecuteLine 9150
            Select Case ButtonMenu.index
'vbwLine 9151:                Case 1
                Case IIf(vbwProfiler.vbwExecuteLine(9151), VBWPROFILER_EMPTY, _
        1)
vbwProfiler.vbwExecuteLine 9152
                    mnuText_Click
'vbwLine 9153:                Case 2
                Case IIf(vbwProfiler.vbwExecuteLine(9153), VBWPROFILER_EMPTY, _
        2)
vbwProfiler.vbwExecuteLine 9154
                    mnuHTML_classic_Click
'vbwLine 9155:                Case 3
                Case IIf(vbwProfiler.vbwExecuteLine(9155), VBWPROFILER_EMPTY, _
        3)
vbwProfiler.vbwExecuteLine 9156
                    mnuHTML_Tables_Click
            End Select
vbwProfiler.vbwExecuteLine 9157 'B
    End Select
vbwProfiler.vbwExecuteLine 9158 'B
vbwProfiler.vbwProcOut 357
vbwProfiler.vbwExecuteLine 9159
End Sub

Private Sub treeLinks_DragDrop(Source As Control, x As Single, y As Single)
vbwProfiler.vbwProcIn 358
    Dim hDropNode As Long

vbwProfiler.vbwExecuteLine 9160
    If m_oCurrentVeh.ActiveProfile = "" Then
vbwProfiler.vbwProcOut 358
vbwProfiler.vbwExecuteLine 9161
         Exit Sub
    End If
vbwProfiler.vbwExecuteLine 9162 'B
vbwProfiler.vbwExecuteLine 9163
    hDropNode = treeLinks.HitTest(x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY)
vbwProfiler.vbwExecuteLine 9164
    If hDropNode > 0 Then
vbwProfiler.vbwExecuteLine 9165
        If TypeName(Source) = "ListView" Then
vbwProfiler.vbwExecuteLine 9166
            If treeLinks.ItemText(hDropNode) = CHILD_GROUP_NAME Then
vbwProfiler.vbwExecuteLine 9167
                Debug.Assert m_oCurrentVeh.ActiveProfiletype > 0
vbwProfiler.vbwExecuteLine 9168
                Call m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).assignconsumer(m_oLstviewLinkDragItem.Key, treeLinks.ItemData(hDropNode))
vbwProfiler.vbwExecuteLine 9169
                If m_oCurrentVeh.ActiveProfiletype = POWER_PROFILE Then
                    ' this just updates the List of Available... should move that function into the Profile class really
vbwProfiler.vbwExecuteLine 9170
                    ShowLinks POWER_PROFILE
                Else
vbwProfiler.vbwExecuteLine 9171 'B
vbwProfiler.vbwExecuteLine 9172
                    ShowLinks FUEL_PROFILE
                End If
vbwProfiler.vbwExecuteLine 9173 'B
vbwProfiler.vbwExecuteLine 9174
                m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).Show
            Else
vbwProfiler.vbwExecuteLine 9175 'B
vbwProfiler.vbwExecuteLine 9176
                InfoPrint 1, "Consumers must be added to child group nodes."
            End If
vbwProfiler.vbwExecuteLine 9177 'B
        End If
vbwProfiler.vbwExecuteLine 9178 'B
    End If
vbwProfiler.vbwExecuteLine 9179 'B
vbwProfiler.vbwExecuteLine 9180
    m_bTreeLinksDrag = False
vbwProfiler.vbwProcOut 358
vbwProfiler.vbwExecuteLine 9181
End Sub

Private Sub treeLinks_DragOver(Source As Control, x As Single, y As Single, State As Integer)
vbwProfiler.vbwProcIn 359
    Dim hDropNode As Long
vbwProfiler.vbwExecuteLine 9182
    hDropNode = treeLinks.HitTest(x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY)
vbwProfiler.vbwExecuteLine 9183
    If hDropNode > 0 Then
vbwProfiler.vbwExecuteLine 9184
         treeLinks.SelectItem(hDropNode) = True
    End If
vbwProfiler.vbwExecuteLine 9185 'B
vbwProfiler.vbwProcOut 359
vbwProfiler.vbwExecuteLine 9186
End Sub

Private Sub treeLinks_ItemDblClk(ByVal hItem As Long)
vbwProfiler.vbwProcIn 360
    Dim hParent As Long
    Dim sKey As String

vbwProfiler.vbwExecuteLine 9187
    If (treeLinks.ItemText(hItem) <> GROUP_NAME) And (treeLinks.ItemText(hItem) <> CHILD_GROUP_NAME) Then
vbwProfiler.vbwExecuteLine 9188
        hParent = treeLinks.ItemParent(hItem)
vbwProfiler.vbwExecuteLine 9189
        If treeLinks.ItemText(hParent) = GROUP_NAME Then
            ' should i allow Groups to exit which have no suppliers attached?
            ' its really only good for keeping children attached and switching out suppliers
            ' Simple way would be no, and if children are attached, dont let user delete last
            ' supplier of a Group node
            ' MPJ 05/27/02 - YES.  Groups can now be void of suppliers as long as consumers are attached to the Child group

        Else
vbwProfiler.vbwExecuteLine 9190 'B
vbwProfiler.vbwExecuteLine 9191
            sKey = m_oCurrentVeh.ActiveProfile
vbwProfiler.vbwExecuteLine 9192
            m_oCurrentVeh.Profiles(sKey).UnAssignConsumer treeLinks.ItemData(hParent), treeLinks.ItemData(hItem)
vbwProfiler.vbwExecuteLine 9193
            ShowLinks m_oCurrentVeh.ActiveProfiletype  ' this just updates the List of Available... should move that function into the Profile class really
vbwProfiler.vbwExecuteLine 9194
            m_oCurrentVeh.Profiles(sKey).Show
        End If
vbwProfiler.vbwExecuteLine 9195 'B
    End If
vbwProfiler.vbwExecuteLine 9196 'B
vbwProfiler.vbwProcOut 360
vbwProfiler.vbwExecuteLine 9197
End Sub

Private Sub treeLinks_ItemDragged(ByVal hItemSource As Long, ByVal hItemTarget As Long, ByVal anDropEffect As Long, pbCancel As Boolean)
vbwProfiler.vbwProcIn 361
    Dim hParent As Long
    Dim sKey As String
    Dim lngSupplier As Long
    Dim lngOldGroup As Long
    Dim lngNewGroup As Long
    Dim hFirstChild As Long

    ' we need to determine if this is either a Group node, Child Group Node, Supplier or Consumer
    ' only suppliers can be moved and only onto other Groups
    ' only consumers can be moved only onto other Child Groups
vbwProfiler.vbwExecuteLine 9198
    Select Case treeLinks.ItemText(hItemSource)
'vbwLine 9199:        Case GROUP_NAME
        Case IIf(vbwProfiler.vbwExecuteLine(9199), VBWPROFILER_EMPTY, _
        GROUP_NAME)
vbwProfiler.vbwExecuteLine 9200
            pbCancel = True
vbwProfiler.vbwExecuteLine 9201
            InfoPrint 1, "Group nodes cannot be moved."
'vbwLine 9202:        Case CHILD_GROUP_NAME
        Case IIf(vbwProfiler.vbwExecuteLine(9202), VBWPROFILER_EMPTY, _
        CHILD_GROUP_NAME)
vbwProfiler.vbwExecuteLine 9203
            pbCancel = True
vbwProfiler.vbwExecuteLine 9204
            InfoPrint 1, "Group nodes cannot be moved."
        Case Else
vbwProfiler.vbwExecuteLine 9205 'B
vbwProfiler.vbwExecuteLine 9206
            hParent = treeLinks.ItemParent(hItemSource)
vbwProfiler.vbwExecuteLine 9207
            lngSupplier = treeLinks.ItemData(hItemSource)
vbwProfiler.vbwExecuteLine 9208
            lngOldGroup = treeLinks.ItemData(hParent)
vbwProfiler.vbwExecuteLine 9209
            lngNewGroup = treeLinks.ItemData(hItemTarget)
vbwProfiler.vbwExecuteLine 9210
            If treeLinks.ItemText(hParent) = GROUP_NAME Then
                ' its a supplier, it can only be placed on another group
vbwProfiler.vbwExecuteLine 9211
                If treeLinks.ItemText(hItemTarget) <> GROUP_NAME Then
vbwProfiler.vbwExecuteLine 9212
                    pbCancel = True
vbwProfiler.vbwExecuteLine 9213
                    InfoPrint 1, "Suppliers can only be placed on other Supply Groups"
                Else
vbwProfiler.vbwExecuteLine 9214 'B
                    'update the actual profile with the changes
vbwProfiler.vbwExecuteLine 9215
                     Call m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).movepowersystem(lngSupplier, lngOldGroup, lngNewGroup)
                End If
vbwProfiler.vbwExecuteLine 9216 'B
            Else
vbwProfiler.vbwExecuteLine 9217 'B
                ' its a consumer, it can only be placed on another child group
vbwProfiler.vbwExecuteLine 9218
                If treeLinks.ItemText(hItemTarget) <> CHILD_GROUP_NAME Then
vbwProfiler.vbwExecuteLine 9219
                    pbCancel = True
vbwProfiler.vbwExecuteLine 9220
                    InfoPrint 1, "Consumers can only be placed on other Consumer Groups"
                Else
vbwProfiler.vbwExecuteLine 9221 'B
                    ' update the actual profile with the changes
vbwProfiler.vbwExecuteLine 9222
                    Call m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).MoveConsumer(lngSupplier, lngOldGroup, lngNewGroup)
                End If
vbwProfiler.vbwExecuteLine 9223 'B
            End If
vbwProfiler.vbwExecuteLine 9224 'B

vbwProfiler.vbwExecuteLine 9225
            On Error GoTo err
vbwProfiler.vbwExecuteLine 9226
            If pbCancel = False Then
vbwProfiler.vbwExecuteLine 9227
                hFirstChild = treeLinks.ItemChild(hParent)
                ' more lameness on the TreeX control.  If we try to "refresh" the control by
                ' deleting all items and then rebuilding the tree, it crashes the vb runtime.
                ' Essentially, this control doesnt like for the .RemoveAllItems method to be called
                ' from within this sub.
                ' within this procedure, it crashes the vb runtime.
                'UPDATE: 07/14/02 - Even with the new Mabry Update, they simply trap a runtime error and
                'still dont allow .RemoveAllItems to be called.
                'todo: would be nice if i can think of better way of doing this...
vbwProfiler.vbwExecuteLine 9228
                m_lngTimerID = SetTimer(0, 0, TIMER_DELAY, AddressOf TimerProc)
            End If
vbwProfiler.vbwExecuteLine 9229 'B
    End Select
vbwProfiler.vbwExecuteLine 9230 'B
vbwProfiler.vbwProcOut 361
vbwProfiler.vbwExecuteLine 9231
    Exit Sub
err:
    ' if no children left under the group, an error is thrown -- delete the group
vbwProfiler.vbwExecuteLine 9232
    If err.Number = -2147417848 Then
vbwProfiler.vbwExecuteLine 9233
        treeLinks.RemoveItem (lngOldGroup)
vbwProfiler.vbwExecuteLine 9234
        Resume Next
    Else
vbwProfiler.vbwExecuteLine 9235 'B
vbwProfiler.vbwExecuteLine 9236
        Debug.Print "frmDesigner.treeLinks_ItemDragged() -- Error #" & err.Number & " - " & err.Description
    End If
vbwProfiler.vbwExecuteLine 9237 'B
vbwProfiler.vbwProcOut 361
vbwProfiler.vbwExecuteLine 9238
End Sub

Private Sub treeLinks_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 362
    Dim hNode As Long
vbwProfiler.vbwExecuteLine 9239
    If Button = vbLeftButton Then
vbwProfiler.vbwExecuteLine 9240
        If m_bTreeLinksDrag Then
vbwProfiler.vbwExecuteLine 9241
            hNode = treeLinks.HitTest(x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY)
vbwProfiler.vbwExecuteLine 9242
            If hNode > 0 Then
vbwProfiler.vbwExecuteLine 9243
                 treeLinks.SelectItem(hNode) = True
            End If
vbwProfiler.vbwExecuteLine 9244 'B
        End If
vbwProfiler.vbwExecuteLine 9245 'B
    End If
vbwProfiler.vbwExecuteLine 9246 'B
vbwProfiler.vbwProcOut 362
vbwProfiler.vbwExecuteLine 9247
End Sub

Private Sub treeLinks_RightClick(ByVal x As Single, ByVal y As Single)
vbwProfiler.vbwProcIn 363
    Dim hParent As Long
    Dim hItem As Long

vbwProfiler.vbwExecuteLine 9248
    hItem = treeLinks.HitTest(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
vbwProfiler.vbwExecuteLine 9249
    If hItem <= 0 Then
vbwProfiler.vbwProcOut 363
vbwProfiler.vbwExecuteLine 9250
         Exit Sub
    End If
vbwProfiler.vbwExecuteLine 9251 'B

vbwProfiler.vbwExecuteLine 9252
    On Error Resume Next
    ' right mouse click popup

vbwProfiler.vbwExecuteLine 9253
    With m_uTreeLinksNodeData
vbwProfiler.vbwExecuteLine 9254
        .hParent = treeLinks.ItemParent(hItem)
vbwProfiler.vbwExecuteLine 9255
        .hwnd = hItem
vbwProfiler.vbwExecuteLine 9256
        .iGroupIndex = treeLinks.ItemData(.hParent)
vbwProfiler.vbwExecuteLine 9257
        .iIndex = treeLinks.ItemData(.hwnd)
vbwProfiler.vbwExecuteLine 9258
        .sText = treeLinks.ItemText(.hwnd)
vbwProfiler.vbwExecuteLine 9259
    End With

    ' determine type of node selected
vbwProfiler.vbwExecuteLine 9260
    Select Case treeLinks.ItemText(hItem)
'vbwLine 9261:        Case GROUP_NAME
        Case IIf(vbwProfiler.vbwExecuteLine(9261), VBWPROFILER_EMPTY, _
        GROUP_NAME)
                 '      NOTE: The Visible item should always be first since
                 '      a VB error will occur if at any time, ALL popup
                 '      items are visible = False.  This can easily happen
                 '      since we are constantly hiding/showing different items
                 '     Simple way to avoid it is to order things so that the first
                 '     item is visible = true
vbwProfiler.vbwProcOut 363
vbwProfiler.vbwExecuteLine 9262
                 Exit Sub ' nothing to do here
'vbwLine 9263:        Case CHILD_GROUP_NAME
        Case IIf(vbwProfiler.vbwExecuteLine(9263), VBWPROFILER_EMPTY, _
        CHILD_GROUP_NAME)
vbwProfiler.vbwExecuteLine 9264
            m_uTreeLinksNodeData.iGroupIndex = treeLinks.ItemData(hItem) ' group nodes use their own index
vbwProfiler.vbwExecuteLine 9265
            mnuConsumerDeleteAll.Visible = True
vbwProfiler.vbwExecuteLine 9266
            mnuConsumerDelete.Visible = False
vbwProfiler.vbwExecuteLine 9267
            mnuSupplierAddToNewGroup.Visible = False
        Case Else
vbwProfiler.vbwExecuteLine 9268 'B
vbwProfiler.vbwExecuteLine 9269
            hParent = treeLinks.ItemParent(hItem)
vbwProfiler.vbwExecuteLine 9270
            Select Case treeLinks.ItemText(hParent)
'vbwLine 9271:                Case GROUP_NAME
                Case IIf(vbwProfiler.vbwExecuteLine(9271), VBWPROFILER_EMPTY, _
        GROUP_NAME)
vbwProfiler.vbwExecuteLine 9272
                    mnuSupplierAddToNewGroup.Visible = True
vbwProfiler.vbwExecuteLine 9273
                    mnuConsumerDeleteAll.Visible = False
vbwProfiler.vbwExecuteLine 9274
                    mnuConsumerDelete.Visible = False
'vbwLine 9275:                Case CHILD_GROUP_NAME
                Case IIf(vbwProfiler.vbwExecuteLine(9275), VBWPROFILER_EMPTY, _
        CHILD_GROUP_NAME)
vbwProfiler.vbwExecuteLine 9276
                    mnuConsumerDelete.Visible = True
vbwProfiler.vbwExecuteLine 9277
                    mnuSupplierAddToNewGroup.Visible = False
vbwProfiler.vbwExecuteLine 9278
                    mnuConsumerDeleteAll.Visible = False
                Case Else
vbwProfiler.vbwExecuteLine 9279 'B
vbwProfiler.vbwExecuteLine 9280
                    Debug.Print "frmDesigner:treeLinks_ItemSelect -- ItemData holds unsupported Parent Type"
vbwProfiler.vbwExecuteLine 9281
                    Debug.Print "ItemData = " & treeLinks.ItemData(hItem)
vbwProfiler.vbwProcOut 363
vbwProfiler.vbwExecuteLine 9282
                    Exit Sub
            End Select
vbwProfiler.vbwExecuteLine 9283 'B
    End Select
vbwProfiler.vbwExecuteLine 9284 'B
vbwProfiler.vbwExecuteLine 9285
    PopupMenu mnuLinksPopup
vbwProfiler.vbwProcOut 363
vbwProfiler.vbwExecuteLine 9286
End Sub

Private Sub mnuSupplierAddToNewGroup_Click()
vbwProfiler.vbwProcIn 364
    Dim iNewGroup As Long

vbwProfiler.vbwExecuteLine 9287
    iNewGroup = m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).CreateNewscGroup
vbwProfiler.vbwExecuteLine 9288
    Call m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).movepowersystem(m_uTreeLinksNodeData.iIndex, m_uTreeLinksNodeData.iGroupIndex, iNewGroup)
vbwProfiler.vbwExecuteLine 9289
    m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).Show
vbwProfiler.vbwProcOut 364
vbwProfiler.vbwExecuteLine 9290
End Sub

Private Sub mnuConsumerDelete_Click()
vbwProfiler.vbwProcIn 365
vbwProfiler.vbwExecuteLine 9291
    InfoPrint 1, "HINT:  Double clicking a consumer will also remove it from the group."
vbwProfiler.vbwExecuteLine 9292
    m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).UnAssignConsumer m_uTreeLinksNodeData.iGroupIndex, m_uTreeLinksNodeData.iIndex
vbwProfiler.vbwExecuteLine 9293
    ShowLinks m_oCurrentVeh.ActiveProfiletype  ' this just updates the List of Available... should move that function into the Profile class really
vbwProfiler.vbwExecuteLine 9294
    m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).Show
vbwProfiler.vbwProcOut 365
vbwProfiler.vbwExecuteLine 9295
End Sub

Private Sub mnuConsumerDeleteAll_Click()
vbwProfiler.vbwProcIn 366
    Dim hItem As Long
    Dim iIndex As Long

vbwProfiler.vbwExecuteLine 9296
    On Error GoTo err
vbwProfiler.vbwExecuteLine 9297
    hItem = m_uTreeLinksNodeData.hwnd

vbwProfiler.vbwExecuteLine 9298
    hItem = treeLinks.ItemChild(hItem)
vbwProfiler.vbwExecuteLine 9299
    iIndex = treeLinks.ItemData(hItem) ' index should always be 1 since whenever we delete one, the next index is always re-ordered to start at 1 again

vbwProfiler.vbwExecuteLine 9300
    Do
vbwProfiler.vbwExecuteLine 9301
        m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).UnAssignConsumer m_uTreeLinksNodeData.iGroupIndex, iIndex
vbwProfiler.vbwExecuteLine 9302
        hItem = treeLinks.ItemNext(hItem)
vbwProfiler.vbwExecuteLine 9303
        If hItem = 0 Then
vbwProfiler.vbwExecuteLine 9304
             Exit Do
        End If
vbwProfiler.vbwExecuteLine 9305 'B
vbwProfiler.vbwExecuteLine 9306
    Loop

    'todo: why is this missing an exit sub before the err?
err:

vbwProfiler.vbwExecuteLine 9307
    ShowLinks m_oCurrentVeh.ActiveProfiletype ' this just updates the List of Available... should move that function into the Profile class really
vbwProfiler.vbwExecuteLine 9308
    m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).Show
vbwProfiler.vbwProcOut 366
vbwProfiler.vbwExecuteLine 9309
End Sub

Private Sub treeVehicle_DblClick()
'    treeVehicle.Sorted = True
vbwProfiler.vbwProcIn 367
vbwProfiler.vbwProcOut 367
vbwProfiler.vbwExecuteLine 9310
End Sub
Private Sub treeVehicle_ItemDraggedTo(ByVal hItemSource As Long, ByVal pTargetTree As TreeXLibCtl.ITreeX, ByVal hItemTarget As Long, ByVal anDropEffect As Long, pbCancel As Boolean)
vbwProfiler.vbwProcIn 368
    Dim hSrc As Long
    Dim hDest As Long

vbwProfiler.vbwExecuteLine 9311
    hSrc = treeVehicle.ItemData(hItemSource)
vbwProfiler.vbwExecuteLine 9312
    hDest = treeVehicle.ItemData(hItemTarget)
vbwProfiler.vbwExecuteLine 9313
    If Not MoveExistingComponent(hSrc, hDest) Then
vbwProfiler.vbwExecuteLine 9314
        pbCancel = True
    End If
vbwProfiler.vbwExecuteLine 9315 'B
vbwProfiler.vbwProcOut 368
vbwProfiler.vbwExecuteLine 9316
End Sub
Private Sub treeVehicle_DragDrop(Source As Control, x As Single, y As Single)
vbwProfiler.vbwProcIn 369
    Dim h As Long
vbwProfiler.vbwExecuteLine 9317
    h = treeVehicle.HitTest(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
vbwProfiler.vbwExecuteLine 9318
    If h Then
vbwProfiler.vbwExecuteLine 9319
        If Source.Name = "ListView1" Then
vbwProfiler.vbwExecuteLine 9320
            Call AddComponentsFromFile(Source.SelectedItem.Key, treeVehicle.ItemData(h), h)
        Else
vbwProfiler.vbwExecuteLine 9321 'B
vbwProfiler.vbwExecuteLine 9322
            MsgBox TypeName(Me) & ":treeVehicle_DragDrop() -- Error.  Item not a valid component type."
        End If
vbwProfiler.vbwExecuteLine 9323 'B
vbwProfiler.vbwExecuteLine 9324
        p_bChangedFlag = True
    End If
vbwProfiler.vbwExecuteLine 9325 'B
vbwProfiler.vbwProcOut 369
vbwProfiler.vbwExecuteLine 9326
End Sub
Private Sub treeVehicle_ItemSelect(ByVal hItem As Long)
vbwProfiler.vbwProcIn 370
vbwProfiler.vbwExecuteLine 9327
    Debug.Assert hItem > 0
vbwProfiler.vbwExecuteLine 9328
    InfoPrint 1, "frmDesigner.treeVehicle_ItemSelect -- Item Handle = " & treeVehicle.ItemData(hItem)
vbwProfiler.vbwExecuteLine 9329
    Properties_Show treeVehicle.ItemData(hItem)
vbwProfiler.vbwProcOut 370
vbwProfiler.vbwExecuteLine 9330
End Sub
Private Sub treeVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not treeVehicle.SelectedItem Is Nothing Then
'        If KeyCode = vbKeyDelete Then
'            mnuDelete_Click
'        End If
'    End If
vbwProfiler.vbwProcIn 371
vbwProfiler.vbwProcOut 371
vbwProfiler.vbwExecuteLine 9331
End Sub
Private Sub treeVehicle_KeyPress(KeyAscii As Integer)
vbwProfiler.vbwProcIn 372
vbwProfiler.vbwExecuteLine 9332
    Select Case KeyAscii
'vbwLine 9333:        Case vbKeyDelete
        Case IIf(vbwProfiler.vbwExecuteLine(9333), VBWPROFILER_EMPTY, _
        vbKeyDelete)
vbwProfiler.vbwExecuteLine 9334
            Call RemoveComponent
    End Select
vbwProfiler.vbwExecuteLine 9335 'B
vbwProfiler.vbwProcOut 372
vbwProfiler.vbwExecuteLine 9336
End Sub
Private Sub treeVehicle_KeyUp(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        '33 = PgUp, 34 = PgDn, 35 = End, 36 = Home, 37 = Left, 38 = Up , 39 = Right, 40 = Down
'        Case 33, 34, 35, 36, 38, 40, 37, 39
'            SetActiveNode
'        Case Else
'            Debug.Print "keycode = " & KeyCode
'    End Select
vbwProfiler.vbwProcIn 373
vbwProfiler.vbwProcOut 373
vbwProfiler.vbwExecuteLine 9337
End Sub
Private Sub treeVehicle_RightClick(ByVal x As Single, ByVal y As Single)
vbwProfiler.vbwProcIn 374
    Dim h As Long
    Dim oNode As cINode
    Dim lngAttributes As Long
    Dim hNode As Long
    Dim f As Boolean

vbwProfiler.vbwExecuteLine 9338
    h = treeVehicle.HitTest(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
vbwProfiler.vbwExecuteLine 9339
    If h Then
vbwProfiler.vbwExecuteLine 9340
        treeVehicle.SelectItem(h) = True
vbwProfiler.vbwExecuteLine 9341
        hNode = treeVehicle.ItemData(h)

vbwProfiler.vbwExecuteLine 9342
        CopyMemory oNode, hNode, 4
vbwProfiler.vbwExecuteLine 9343
        lngAttributes = oNode.Attributes
vbwProfiler.vbwExecuteLine 9344
        CopyMemory oNode, 0&, 4

vbwProfiler.vbwExecuteLine 9345
        f = (lngAttributes And NODE_REQUIRED)
vbwProfiler.vbwExecuteLine 9346
        If f Then
vbwProfiler.vbwExecuteLine 9347
            mnuDelete.Visible = False
        Else
vbwProfiler.vbwExecuteLine 9348 'B
vbwProfiler.vbwExecuteLine 9349
            mnuDelete.Visible = True
        End If
vbwProfiler.vbwExecuteLine 9350 'B
vbwProfiler.vbwExecuteLine 9351
        f = (lngAttributes And NODE_COPYABLE)
vbwProfiler.vbwExecuteLine 9352
        If f Then
vbwProfiler.vbwExecuteLine 9353
            mnuCopy.Visible = True
        Else
vbwProfiler.vbwExecuteLine 9354 'B
vbwProfiler.vbwExecuteLine 9355
            mnuCopy.Visible = False
        End If
vbwProfiler.vbwExecuteLine 9356 'B
vbwProfiler.vbwExecuteLine 9357
        f = (lngAttributes And NODE_RENAMEABLE)
vbwProfiler.vbwExecuteLine 9358
        If f Then
vbwProfiler.vbwExecuteLine 9359
            mnuRename.Visible = True
        Else
vbwProfiler.vbwExecuteLine 9360 'B
vbwProfiler.vbwExecuteLine 9361
            mnuRename.Visible = False
        End If
vbwProfiler.vbwExecuteLine 9362 'B

        ' todo:  remember, not allow delete for GroupComponent, Arm, ArmMotor, OrnithopterDrivetrain, LegDrivetrain, BattlesuitSystem, FormFittingBattleSuitSystem
        ' todo: cant save or copy wings, ornithopterdrivetrains, armor motors,etc.  these are strictly in the XML def's though i believe
vbwProfiler.vbwExecuteLine 9363
        PopupMenu mnuPopup
    End If
vbwProfiler.vbwExecuteLine 9364 'B
vbwProfiler.vbwProcOut 374
vbwProfiler.vbwExecuteLine 9365
End Sub


'/////////////////////////////////////////////////////////////////////////
'RESIZE CODE
'////////////////////////////////////////////////////////////////////////
Public Sub tabVehicle_resize()
vbwProfiler.vbwProcIn 375

vbwProfiler.vbwProcOut 375
vbwProfiler.vbwExecuteLine 9366
End Sub
Public Sub TabStrip_Resize()
vbwProfiler.vbwProcIn 376
vbwProfiler.vbwExecuteLine 9367
     PLC1.ColumnWidth = PLC1.Width / Screen.TwipsPerPixelX / 2
    ' resize splitter contained in TabMain(0)
vbwProfiler.vbwExecuteLine 9368
    If Not oVSplitter2 Is Nothing Then
vbwProfiler.vbwExecuteLine 9369
        oVSplitter2.ParentResized
    End If
vbwProfiler.vbwExecuteLine 9370 'B
vbwProfiler.vbwExecuteLine 9371
    Call ResizeTabbedChildren
vbwProfiler.vbwProcOut 376
vbwProfiler.vbwExecuteLine 9372
End Sub

Public Sub ResizeTabbedChildren()
vbwProfiler.vbwProcIn 377
    Dim r As Rect
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngHeight As Long
    Dim lngWidth As Long

vbwProfiler.vbwExecuteLine 9373
    On Error Resume Next

vbwProfiler.vbwExecuteLine 9374
    Call GetWindowRect(tabVehicle.hwnd, r)

vbwProfiler.vbwExecuteLine 9375
    lngLeft = 160
vbwProfiler.vbwExecuteLine 9376
    lngTop = tabVehicle.Top + 560
vbwProfiler.vbwExecuteLine 9377
    lngWidth = r.Right - r.Left '- 60 '60 for right padding
vbwProfiler.vbwExecuteLine 9378
    lngWidth = (lngWidth - 10) * Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 9379
    lngHeight = r.Bottom - r.Top
vbwProfiler.vbwExecuteLine 9380
    lngHeight = (lngHeight - 30) * Screen.TwipsPerPixelY

    ' resize the infoText in SubTab(1)
vbwProfiler.vbwExecuteLine 9381
    With treeVehicle
vbwProfiler.vbwExecuteLine 9382
        .Left = lngLeft
vbwProfiler.vbwExecuteLine 9383
        .Top = lngTop
vbwProfiler.vbwExecuteLine 9384
        .Width = lngWidth
vbwProfiler.vbwExecuteLine 9385
        .Height = lngHeight
vbwProfiler.vbwExecuteLine 9386
    End With

vbwProfiler.vbwExecuteLine 9387
    Call GetWindowRect(tabSub(1).hwnd, r)

vbwProfiler.vbwExecuteLine 9388
    lngLeft = 60
vbwProfiler.vbwExecuteLine 9389
    lngTop = 360
vbwProfiler.vbwExecuteLine 9390
    lngWidth = r.Right - r.Left '- 60 '60 for right padding
vbwProfiler.vbwExecuteLine 9391
    lngWidth = (lngWidth - 10) * Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 9392
    lngHeight = r.Bottom - r.Top
vbwProfiler.vbwExecuteLine 9393
    lngHeight = (lngHeight - 30) * Screen.TwipsPerPixelY

    ' resize the infoText in SubTab(1)
vbwProfiler.vbwExecuteLine 9394
    With txtInfo
vbwProfiler.vbwExecuteLine 9395
        .Left = lngLeft
vbwProfiler.vbwExecuteLine 9396
        .Top = lngTop
vbwProfiler.vbwExecuteLine 9397
        .Width = lngWidth
vbwProfiler.vbwExecuteLine 9398
        .Height = lngHeight
vbwProfiler.vbwExecuteLine 9399
    End With

vbwProfiler.vbwExecuteLine 9400
    With lstStats
vbwProfiler.vbwExecuteLine 9401
        .Left = lngLeft
vbwProfiler.vbwExecuteLine 9402
        .Top = lngTop
vbwProfiler.vbwExecuteLine 9403
        .Width = lngWidth
vbwProfiler.vbwExecuteLine 9404
        .Height = lngHeight
vbwProfiler.vbwExecuteLine 9405
    End With

vbwProfiler.vbwExecuteLine 9406
    Call GetWindowRect(tabMain(0).hwnd, r)

vbwProfiler.vbwExecuteLine 9407
    lngWidth = r.Right - r.Left '- 60 '60 for right padding
vbwProfiler.vbwExecuteLine 9408
    lngWidth = (lngWidth - 10) * Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 9409
    lngHeight = r.Bottom - r.Top
vbwProfiler.vbwExecuteLine 9410
    lngHeight = (lngHeight - 30) * Screen.TwipsPerPixelY

vbwProfiler.vbwExecuteLine 9411
    With picVehicleImage
vbwProfiler.vbwExecuteLine 9412
        .Left = lngLeft
vbwProfiler.vbwExecuteLine 9413
        .Top = lngTop
vbwProfiler.vbwExecuteLine 9414
        .Width = lngWidth
vbwProfiler.vbwExecuteLine 9415
        .Height = lngHeight
vbwProfiler.vbwExecuteLine 9416
    End With

vbwProfiler.vbwExecuteLine 9417
    With rtbOutput
vbwProfiler.vbwExecuteLine 9418
        .Left = lngLeft
vbwProfiler.vbwExecuteLine 9419
        .Top = lngTop
vbwProfiler.vbwExecuteLine 9420
        .Width = lngWidth
vbwProfiler.vbwExecuteLine 9421
        .Height = lngHeight
vbwProfiler.vbwExecuteLine 9422
    End With
vbwProfiler.vbwProcOut 377
vbwProfiler.vbwExecuteLine 9423
End Sub

Private Sub Form_Resize()
vbwProfiler.vbwProcIn 378
vbwProfiler.vbwExecuteLine 9424
    On Error Resume Next
    'JAW 2000.05.07
    'added IF/THEN to skip resizing if form was only minimized. This also
    'prevents an error where some bad formulas lead to negative parameters
    'being used in cSplitter.Move, but I have not had time to dig that deeply.
vbwProfiler.vbwExecuteLine 9425
    If Me.windowstate <> vbMinimized Then
        '//assumes we are dealing in Twips
vbwProfiler.vbwExecuteLine 9426
        Const RIGHT_BORDER_WIDTH = 60
vbwProfiler.vbwExecuteLine 9427
        Const BOTTOM_BORDER_HEIGHT = 60
        Dim lngWidth As Long
        Dim lngHeight As Long

vbwProfiler.vbwExecuteLine 9428
        If Me.Width < 9720 Then 'min width is 640 pixels
vbwProfiler.vbwExecuteLine 9429
             Me.Width = 9720
        End If
vbwProfiler.vbwExecuteLine 9430 'B
vbwProfiler.vbwExecuteLine 9431
        If Me.Height < 7890 Then 'min height is 480 pixels
vbwProfiler.vbwExecuteLine 9432
             Me.Height = 7890
        End If
vbwProfiler.vbwExecuteLine 9433 'B
vbwProfiler.vbwExecuteLine 9434
        oVSplitter1.ParentResized
vbwProfiler.vbwExecuteLine 9435
        oHSplitter.ParentResized
   End If
vbwProfiler.vbwExecuteLine 9436 'B
vbwProfiler.vbwProcOut 378
vbwProfiler.vbwExecuteLine 9437
End Sub



'//TODO: todo:  delete all this battlesuit crap which follows.  Its all  obsolete and will be replaced with .cmp defaults
''Private Sub mnuBattleSuitBody_Click()
''    'user has selected to make a battlesuit where the Pilot is concealed entirely in the body
''    Dim i As Integer
''
''    'prepare for new vehicle
''    If InitNewVehicle = False Then Exit Sub
''    Call Init2
''
''    ' save all the settings
''    With m_oCurrentVeh.Components(BODY_KEY)
''        .TL = 8
''        .BattleSuitVolumeAdded = True
''        .component = frmDesigner.treeVehicle.Nodes.item(BODY_KEY).Text
''    End With
''    With m_oCurrentVeh.Options
''        .BattleSuit = "Pilot in Body"
''        .UseSurfaceAreaTable = Settings.bUseSurfaceAreaTable
''    End With
''
''    ' battlesuit system
''    frmDesigner.treeVehicle.Nodes.Add BODY_KEY, tvwChild, BATTLESUIT_KEY_SYSTEM, "non form fitting battlesuit system", 58 '44 is the 44th item in imagelist2
''    m_oCurrentVeh.addObject BattlesuitSystem, BATTLESUIT_KEY_SYSTEM, BODY_KEY, 58, "non form fitting battlesuit system", False
''    Call InitStage3
''
''End Sub
'''
''Private Sub mnuBattleSuitTurret_Click()
''    ''user has selected to make a battlesuit where the Pilot is concealed entirely in the Turret
''    Dim i As Integer
''
''    'prepare for new vehicle
''    If InitNewVehicle = False Then Exit Sub
''    Call Init2
''
''    ' save all the settings
''    With m_oCurrentVeh.Components(BODY_KEY)
''        .TL = 8
''        .BattleSuitVolumeAdded = True
''        .component = frmDesigner.treeVehicle.Nodes.item(BODY_KEY).Text
''    End With
''
''    'todo: how do these things get set via a component file? Is there a better way to handle this?
''    With m_oCurrentVeh.Options
''        .BattleSuit = "Pilot in Turret"
''        .UseSurfaceAreaTable = Settings.bUseSurfaceAreaTable
''    End With
''
''    ' battlesuit system
''    frmDesigner.treeVehicle.Nodes.Add BODY_KEY, tvwChild, BATTLESUIT_KEY_SYSTEM, "non form fitting battlesuit system", 58  '44 is the 44th item in imagelist2
''    m_oCurrentVeh.addObject BattlesuitSystem, BATTLESUIT_KEY_SYSTEM, BODY_KEY, 58, "non form fitting battlesuit system", False
''
''    ' head
''    frmDesigner.treeVehicle.Nodes.Add BODY_KEY, tvwChild, BATTLESUIT_KEY_HEAD, "turret", 24
''    m_oCurrentVeh.addObject Turret, BATTLESUIT_KEY_HEAD, BODY_KEY, 24, "turret", False
''    ' save all the settings for the turret
''    m_oCurrentVeh.Components(BATTLESUIT_KEY_HEAD).BattleSuitVolumeAdded = True
''    Call InitStage3
''End Sub
''
''Private Sub mnuBattleSuitFF_Click()
''    Dim ParentsKey As String
''    Dim i As Integer
''
''    'prepare for new vehicle
''    If InitNewVehicle = False Then Exit Sub
''
''    Call Init2
''
''    ' save all the settings
''    With m_oCurrentVeh.Components(BODY_KEY)
''        .TL = 8
''        .BattleSuitVolumeAdded = True
''        .component = frmDesigner.treeVehicle.Nodes.item(BODY_KEY).Text
''    End With
''    With m_oCurrentVeh.Options
''        .BattleSuit = "Form Fitting"
''        .UseSurfaceAreaTable = Settings.bUseSurfaceAreaTable
''    End With
''
''    'TODO: Switch this crap to a script file of some kind.
''    ' battle suit system
''    frmDesigner.treeVehicle.Nodes.Add BODY_KEY, tvwChild, BATTLESUIT_KEY_SYSTEM, "form fitting battlesuit system", 58
''    m_oCurrentVeh.addObject FormFittingBattleSuitSystem, BATTLESUIT_KEY_SYSTEM, BODY_KEY, 58, "form fitting battlesuit system", False
''
''    ' head
''    frmDesigner.treeVehicle.Nodes.Add BODY_KEY, tvwChild, BATTLESUIT_KEY_HEAD, "turret", 24
''    m_oCurrentVeh.addObject Turret, BATTLESUIT_KEY_HEAD, BODY_KEY, 24, "turret", False
''    ' save all the settings
''    m_oCurrentVeh.Components(BATTLESUIT_KEY_HEAD).BattleSuitVolumeAdded = True
''
''    ' left arm
''    frmDesigner.treeVehicle.Nodes.Add BODY_KEY, tvwChild, BATTLESUIT_KEY_LEFT_ARM, "arm", 14
''    m_oCurrentVeh.addObject Arm, BATTLESUIT_KEY_LEFT_ARM, BODY_KEY, 14, "arm", False
''    ' save all the settings
''    m_oCurrentVeh.Components(BATTLESUIT_KEY_LEFT_ARM).BattleSuitVolumeAdded = True
''
''    ' left arm motor
''    ParentsKey = BATTLESUIT_KEY_LEFT_ARM
''    frmDesigner.treeVehicle.Nodes.Add ParentsKey, tvwChild, BATTLESUIT_KEY_LEFT_ARM_MOTOR, "arm motor", 66
''    m_oCurrentVeh.addObject ArmMotor, BATTLESUIT_KEY_LEFT_ARM_MOTOR, ParentsKey, 66, "arm motor", False
''    treeVehicle.Nodes.item(BATTLESUIT_KEY_LEFT_ARM_MOTOR).EnsureVisible ' expand the tree branch
''
''    ' right arm
''    frmDesigner.treeVehicle.Nodes.Add BODY_KEY, tvwChild, BATTLESUIT_KEY_RIGHT_ARM, "arm", 14
''    m_oCurrentVeh.addObject Arm, BATTLESUIT_KEY_RIGHT_ARM, BODY_KEY, 14, "arm", False
''    ' save all the settings
''    m_oCurrentVeh.Components(BATTLESUIT_KEY_RIGHT_ARM).BattleSuitVolumeAdded = True
''    m_oCurrentVeh.Components(BATTLESUIT_KEY_RIGHT_ARM).Orientation = "right"
''
''    ' right arm motor
''    ParentsKey = BATTLESUIT_KEY_RIGHT_ARM
''    frmDesigner.treeVehicle.Nodes.Add ParentsKey, tvwChild, BATTLESUIT_KEY_RIGHT_ARM_MOTOR, "arm motor", 66
''    m_oCurrentVeh.addObject ArmMotor, BATTLESUIT_KEY_RIGHT_ARM_MOTOR, ParentsKey, 66, "arm motor", False
''    treeVehicle.Nodes.item(BATTLESUIT_KEY_RIGHT_ARM_MOTOR).EnsureVisible ' expand the tree branch
''
''    ' left leg
''    frmDesigner.treeVehicle.Nodes.Add BODY_KEY, tvwChild, BATTLESUIT_KEY_LEFT_LEG, "leg", 12
''    m_oCurrentVeh.addObject Leg, BATTLESUIT_KEY_LEFT_LEG, BODY_KEY, 12, "leg", False
''    ' save all the settings
''    m_oCurrentVeh.Components(BATTLESUIT_KEY_LEFT_LEG).BattleSuitVolumeAdded = True
''
''    ' left leg motor
''    ParentsKey = BATTLESUIT_KEY_LEFT_LEG
''    frmDesigner.treeVehicle.Nodes.Add ParentsKey, tvwChild, BATTLESUIT_KEY_LEFT_LEG_MOTOR, "leg motor", 110
''    m_oCurrentVeh.addObject LegDrivetrain, BATTLESUIT_KEY_LEFT_LEG_MOTOR, ParentsKey, 110, "leg motor", False
''    treeVehicle.Nodes.item(BATTLESUIT_KEY_LEFT_LEG_MOTOR).EnsureVisible ' expand the tree branch
''
''    ' right leg
''    frmDesigner.treeVehicle.Nodes.Add BODY_KEY, tvwChild, BATTLESUIT_KEY_RIGHT_LEG, "leg", 12
''    m_oCurrentVeh.addObject Leg, BATTLESUIT_KEY_RIGHT_LEG, BODY_KEY, 12, "leg", False
''    ' save all the settings
''    m_oCurrentVeh.Components(BATTLESUIT_KEY_RIGHT_LEG).BattleSuitVolumeAdded = True
''
''    ' right  leg motor
''    ParentsKey = BATTLESUIT_KEY_RIGHT_LEG
''    frmDesigner.treeVehicle.Nodes.Add ParentsKey, tvwChild, BATTLESUIT_KEY_RIGHT_LEG_MOTOR, "leg motor", 110
''    m_oCurrentVeh.addObject LegDrivetrain, BATTLESUIT_KEY_RIGHT_LEG_MOTOR, ParentsKey, 110, "leg motor", False
''    treeVehicle.Nodes.item(BATTLESUIT_KEY_RIGHT_LEG_MOTOR).EnsureVisible ' expand the tree branch
''
''    Call InitStage3
''End Sub




