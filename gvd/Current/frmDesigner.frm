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
    Call SetHook(cboComponents.hwnd, "test1", True)
    ' subclass the tabstrips so that we know when they are resized
    Call SetHook(tabMain(0).hwnd, "test2", True)
    'Call SetHook(tabSub(1).hWnd, True)
End Sub

Private Sub RemoveHooks()
    ' remove the subclassing hook from the cboComponents
    Call SetHook(cboComponents.hwnd, "test1", False)
    Call SetHook(tabMain(0).hwnd, "test2", False)
    'Call SetHook(tabSub(1).hWnd, False)
End Sub

Private Sub Form_Load()
    Dim lngStyle As Long
           
    ' set our subclass hooks
    Call SetHooks
    
    'set our common dialog object
    'todo: verify in .Unload() all these are being unloaded
    Set m_oCmdlg = New clsCmdlg
    Set m_oRecentFiles = New clsRecentFileManager
    Set m_oRecentFiles.Menu = frmDesigner.mnuRecent(0)
    Set m_oManager = New Vehicles.cManager
    'todo: 'set a reference to our text dsplay area
    'm_oFactory.SetMessageTextBox txtInfo
    'm_oFactory.FormatString = Settings.FormatString
    
    Me.Caption = APP_CAPTION
    Me.ScaleMode = vbTwips
    lstviewLinks.LabelEdit = lvwManual
    cboComponents.Text = "Components"
    tabVehicle.Tabs.Remove (1) ' dont know why this stupid control always adds a tab after ive deleted them all at design time... oh well.
    
    ' Set listbox to report mode
    With ListView1
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Component"
        'hide the column headers
        .HideColumnHeaders = True
    End With
    PLC1.Clear
              
    'must now set the screen display
    With Settings
        GVDVehiclesPath = .InitialDir
        Me.mnuUseSurfaceAreaTable.Checked = Abs(.bUseSurfaceAreaTable)
        Me.Toolbar1.Buttons.Item(9).value = Abs(.bUseSurfaceAreaTable)
            
        'check to see if the desktop resolution has changed since the last run
        If (Screen.Width <> .DesktopX) And (Screen.Height <> .DesktopY) Then
            MsgBox "Desktop settings have changed since last run. GVD will use default settings."
            If .windowstate = vbMinimized Then .windowstate = vbNormal
            Me.windowstate = .windowstate
            MoveWindow Me.hwnd, 0, 0, 640, 480, 0
            .Splitter1 = DEFAULT_SPLITTER1_LEFT
            .Splitter2 = tabMain(0).Width / Screen.TwipsPerPixelX / 2
            .HSplitter = DEFAULT_HSPLITTER_TOP
        Else
            Me.windowstate = .windowstate
            MoveWindow Me.hwnd, .FormLeft, .FormTop, .FormWidth, .FormHeight, 0
        End If
        
        Call ConfigureMainSplitters
        ' set the main tab to display the "Components Tab"
        Call SetViewMode(component_view)
        ' set the sub tab to start with the "Alerts Tab"
        tabSub(1).Tab = 0
        
        'set our recent file list
        m_oRecentFiles.AddRecentFile Settings.Recent2
        m_oRecentFiles.AddRecentFile Settings.Recent2
        m_oRecentFiles.AddRecentFile Settings.Recent3
        m_oRecentFiles.AddRecentFile Settings.Recent4
        m_oRecentFiles.AddRecentFile Settings.Recent5
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' If there is already a Vehicle loaded
    If m_bLoadedFlag And p_bChangedFlag Then
        Select Case MsgBox("Save changes to current vehicle?", vbYesNoCancel + 64, "Save Confirmation")
            Case vbYes
                SaveVehicle  ' call the general save sub
            Case vbCancel
                Cancel = True ' cancel out of the unload event
        End Select
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    DoEvents
    On Error Resume Next
    
    Call RemoveHooks
    Set m_oLstviewLinkDragItem = Nothing
    Set m_oManager = Nothing
    
    ' prior to calling WriteINI and prior to destroying the splitters, update the Settings UDT
    With Settings
        .FormLeft = Me.ScaleLeft / Screen.TwipsPerPixelX
        .FormTop = Me.ScaleTop / Screen.TwipsPerPixelY
        .FormWidth = Me.ScaleWidth / Screen.TwipsPerPixelX
        .FormHeight = Me.ScaleHeight / Screen.TwipsPerPixelY
        .HSplitter = oHSplitter.Top
        .Splitter1 = oVSplitter1.Left
    End With
    
    ' destroy the splitters
    Set oHSplitter = Nothing
    Set oVSplitter1 = Nothing
    Set oVSplitter2 = Nothing
    
    ' destroy our dropdown
    Set m_oCBO = Nothing
    Set m_oCmdlg = Nothing
    Set m_oRecentFiles = Nothing
    
    WriteINI ' call the procedure to save the users settings to the "designer.ini"
    WriteLicenseFile
    
    UnloadAllForms Me.Name
End Sub

'Load a Standard Vehicle, Not a battlesuit
Private Sub mnuNewVehicle_Click()
    LoadNewVehicle
End Sub

Public Sub LoadNewVehicle()
    Dim sFilePath As String
    sFilePath = App.Path & NEWE_VEHICLE_COMPONENT
    If LoadVehicle(sFilePath) Then
        Call setGUID
        ' NOTE: This must be done AFTER a vehicle is loaded (either new or from command line)
        ' set the toolbar states for registered or unregistered version
        SetRegisteredToolbarButtonStates
    End If
End Sub

Private Function InitVehicleInstance() As Long
    Dim sKey As String
    Dim lngTempIcon As Long
    Dim lp As Long
    Dim oBase As Vehicles.cINode
    
    lngTempIcon = 1
    
    'clear any ghost GUI stuff from previous vehicle if any
    lstviewLinks.ListItems.Clear              '<--- Make sure we dont leave ghosts in the lists
    treeLinks.RemoveAllItems                  '<--- Make sure we dont leave ghosts in the lists
    lstPropulsionSystems.Clear
    txtInfo = ""
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
                          
    p_ActiveNode.CustomDescription = ""
    p_ActiveNode.Parent = ""
    p_ActiveNode.Datatype = 0
    p_ActiveNode.ParentDataType = 0
    
    InfoPrint 1, FIRST_TASK_MESSAGE
    frmDesigner.Caption = NEW_VEHICLE_CAPTION
    m_sVehicleFile = NEW_VEHICLE_DEFAULT_FILENAME
    frmDesigner.treeVehicle.RemoveAllItems
    Call FillListViewDefaults
        
    'Set the flags and states
    p_bChangedFlag = False 'JAW 2000.05.07 of course no changes can have been made yet
    InitVehicleInstance = lp
    Exit Function
err:
    Debug.Print "frmDesigner.InitVehicleInstance -- Error#" & err.Number & "  " & err.Description
End Function

Private Function OpenNewVehicle() As Long
    ' Check if there is already a vehicle loaded and give user option to save it
    If m_bLoadedFlag And p_bChangedFlag Then
        Select Case MsgBox("Save changes to current vehicle?", vbYesNoCancel + 64, "Save Confirmation")
            Case vbYes
                SaveVehicle  ' call the general save sub
            Case vbCancel
                OpenNewVehicle = False
                Exit Function 'user clicks Cancel.. Exit the sub
            Case vbNo
        End Select
    End If
    OpenNewVehicle = True
End Function

Private Sub mnuOpen_Click()
    Dim i As Integer
    Dim oCDLG As clsCmdlg
    
    ' Code to display the Open commond dialog and then handle the loading
    ' of a file if one is selected
    Dim bSuccess As Boolean ' detects whether the user clicks cancel at the Open dialog
    On Error GoTo errorhandler
    
    ' If there is already a Vehicle loaded
    If OpenNewVehicle Then
        Set oCDLG = New clsCmdlg
        oCDLG.Filter = OPEN_SAVE_DIALOG_FILTER
        oCDLG.CancelError = True
        
        Dim oFile As FileSystemObject
        Set oFile = New FileSystemObject
        If oFile.FolderExists(Settings.VehiclesOpenPath) Then
            oCDLG.InitialDir = Settings.VehiclesOpenPath
        Else
            oCDLG.InitialDir = App.Path
        End If
        bSuccess = oCDLG.ShowOpen(Me.hwnd)
        
        If bSuccess Then
            p_sGUID = ""
            If LoadVehicle(oCDLG.cFileName(1)) Then
                Call setGUID
            End If
        End If
    End If
    Exit Sub
errorhandler:
    Debug.Print "frmDesigner.mnuOpen_Click() - Error # " & err.Number & " " & err.Description
End Sub

Public Function LoadVehicle(sFilePath As String) As Boolean
    Dim lptr As Long
    Dim lngErrorCode As Long
    Dim index As Long
    Dim sKey As String
    
    Debug.Print sFilePath
    
        'prepare for new vehicle
    If OpenNewVehicle Then
        Call InitVehicleInstance
        lptr = m_oManager.createVehicle(sFilePath)
        If lptr <> 0 Then
            ' add it to the tabs
            index = tabVehicle.Tabs.Count + 1
            sKey = KeyFromLong(lptr)
            tabVehicle.Tabs.Add index, sKey, sKey
            
            ' graph it
            GraphVehicle treeVehicle, 0, lptr

            LoadVehicle = True
            m_bLoadedFlag = True
            m_bSavedFlag = False  ' the vehicle has not been saved yet
            ' Call m_oRecentFiles.DeleteRecentFile(sFilePath)  'todo: fix most recent file listing and uncomment this
            Exit Function
        End If
    End If
err:
     m_bLoadedFlag = False
    LoadVehicle = False
    MsgBox TypeName(Me) & ":LoadVehicle() -- Error.  Failed to load vehicle."
End Function
Private Sub mnuPaste_Click()
    Dim hTreeParent As Long
    Dim hObject As Long
    Dim f As Boolean
    
    hTreeParent = treeVehicle.Selection
    If hTreeParent <> 0 Then
        hObject = treeVehicle.ItemData(hTreeParent)
        If m_sClipboard <> "" Then
            f = AddComponentsFromFile(m_sClipboard, hTreeParent, hObject)
        Else
            MsgBox "Nothing to paste."
            Exit Sub
        End If
    End If
    
    If Not f Then
        MsgBox "Node could not be pasted."
    End If
End Sub
Private Sub mnuCopy_Click()
    If Not copyNode(treeVehicle.Selection, False) Then
        MsgBox "Node cannot be copied."
    End If
End Sub
Private Sub mnuCopyBranch_Click()
    If Not copyNode(treeVehicle.Selection, True) Then
        MsgBox "Branch cannot be copied."
    End If
End Sub
Private Sub mnuSaveComponent_Click()
    'frmSaveComponent.Show vbModal, Me
    'Set frmSaveComponent = Nothing
    ' todo: devise way to pick save file name AND wht about component categories?
    '       do we save components in same location as intrinsic? are intrinsic actually in a .PAK or ZIP?
    '       maybe use a toggle button to Show/Unshow user created components within the same listview as intrinsic?
    m_oManager.saveNode treeVehicle.ItemData(treeVehicle.Selection), "c:\testsavexml.txt", True
End Sub
Private Function copyNode(ByVal hTreeNode As Long, ByVal fRecurse As Boolean) As Boolean
    Dim sFileName As String
    Dim szBuf As String * 256
    Dim lRet As Long
    Dim sPath As String
    Const BUFFER_SIZE = 512
   
    sPath = Space$(BUFFER_SIZE)
    sFileName = Space$(BUFFER_SIZE)
    ' - create a temp file, or overright existing?
    lRet = GetTempPath(BUFFER_SIZE, sPath)
    szBuf = sPath
    lRet = GetTempFileName(szBuf, "tmp", 0, sFileName)
    ' - save node using temp filename
    m_sClipboard = Trim$(sFileName)
    If m_oManager.copyNode(treeVehicle.ItemData(hTreeNode), m_sClipboard, fRecurse) Then
        Debug.Print "CLIPBOARD = " & m_sClipboard
        copyNode = True
        Exit Function
    End If
err:
    copyNode = False
    m_sClipboard = ""
End Function
Private Sub mnuDelete_Click()
    Call RemoveComponent
End Sub
Private Sub mnuRename_Click()
    Call RenameComponent
End Sub
Private Sub RenameComponent()
    Dim hNode As Long
    Dim sName As String
    Dim hSelected As Long
    
    sName = InputBox("Enter new name", "Rename")
    If IsValidFilename(sName) Then
        hSelected = treeVehicle.Selection
        hNode = treeVehicle.ItemData(hSelected)
        If Not m_oManager.renameNode(hNode, sName) Then
            MsgBox "You are not allowed to rename this node."
        Else
            treeVehicle.ItemText(hSelected) = sName
        End If
    Else
        MsgBox "Name contains invalid characters."
    End If
End Sub
Private Sub mnuRevert_Click()
    ' todo: if we allow this, there are two options
    ' 1) we call a .Revert function of sorts from the node via oManager which reads its original name from the XML
    ' 2) we store the origianl read only name in the node itself and call oManager.Revert (hNode) as string
End Sub
Private Sub RemoveComponent()
    Dim f As Boolean
    Dim hObject As Long
    Dim hSelected As Long
    
    ' obtain the selected tree node, retreive the pointer to the vehicle component it repesents
    ' and delete them both
    hSelected = treeVehicle.Selection
    If hSelected Then
        hObject = treeVehicle.ItemData(hSelected)
        f = m_oManager.DeleteNode(hObject)
        If f Then
            treeVehicle.RemoveItem (hSelected)
        Else
            Debug.Print TypeName(Me) & ":RemoveComponent() -- Could not delete node"
        End If
    End If
    p_bChangedFlag = True ' JAW 2000.05.07
End Sub
Private Sub AddComponentsFromFile(ByRef sSourceKey As String, ByVal hNodeParent As Long, ByVal hTreeParent As Long)
    Dim sFilePath As String
    Dim hChild As Long
    Dim lngNodeCount As Long
    
    lngNodeCount = 20 'todo: need function for calc'ing nodeCount since apparently there isnt a count property in the control
                      ' todo: Its still possible to add more nodes by simply dragging a saved file that has tons of children
                      ' onto the tree since we never count how many child nodes exist inthe saved file.  Investigate solutions.
    If Not lngNodeCount = MAX_NODES Then
        sFilePath = sSourceKey  ' from listview, the key is actually the full path
        hChild = m_oManager.addNode(sFilePath, hNodeParent)
        If hChild <> 0 Then
            ' add it to the tree
            GraphVehicle treeVehicle, hTreeParent, hChild
        Else
            ' could not add the child to the parent object, possible reasons are
            ' its a leaf node, invalid location given the type, max nodes, etc
            InfoPrint 1, "frmDesigner:AddComponentsFromFile() --  Could not add node to parent. Possible reasons are its a leaf node, invalid location given the type, max node cound reached?"
        End If
    Else
        MsgBox "Node count reached.  You cannot have more than '" & MAX_NODES & "' nodes in the tree."
    End If
End Sub

Private Function MoveExistingComponent(ByVal hSrc As Long, ByVal hDest As Long) As Boolean
    MoveExistingComponent = m_oManager.moveNode(hSrc, hDest)
End Function

Private Function getCurrentVehicle() As Long
' returns handle to cVehicle
    Dim oTabs As Tabs
    Dim i As Long
    
    For i = 1 To tabVehicle.Tabs.Count
        If tabVehicle.Tabs.Item(i).Selected Then
            getCurrentVehicle = Val(tabVehicle.Tabs.Item(i).Key)
            Exit Function
        End If
    Next
    getCurrentVehicle = 0
End Function
Private Function RemoveVehicle(ByVal rootHanlde As Long)
    Dim cTab As TabStrip.Tab
    Dim i As Long
    Dim sKey As String
    
    If m_oManager.deleteVehicle(rootHandle) Then
        sKey = KeyFromLong(rootHandle)
        For Each cTab In tabVehicle.Tabs
            If tabVehicle.Tabs.Item(i).Key = sKey Then
                tabVehicle.Tabs.Remove (i)
                Exit For
            End If
            i = i + 1
        Next
    Else
        MsgBox TypeName(Me) & ":RemoveVehicle() -- Error: Could not delete vehicle"
    End If
End Function

Private Sub SaveVehicle()
    Dim sTemp As String
    Dim f As Boolean
    ' If the file has already been saved with a valid fileName
    If m_bSavedFlag Then

        sTemp = StatusBar1.Panels(1).Text 'save the current text so we can restore it
        StatusBar1.Panels(1).Picture = ImageList1.ListImages(8).Picture
        StatusBar1.Panels(1).Text = "serializing vehicle data..."
        
        'todo: i could first check for the current active tabstrip tab
        '      and then check its tag property for m_sVehicleFile.  if filename
        '      is "" , thenits not saved yet.  Problem with m_sVehicleFile variable though
        '      is it either needs to be an array sorted via root node handles or
        '      just use the tabs and forget the var altogether
        f = m_oManager.saveNode(treeVehicle.ItemData(treeVehicle.Root(0)), m_sVehicleFile, True)
        Debug.Print "SAVE OPERATION " & f & " File output = " & m_sVehicleFile
        
        StatusBar1.Panels(1).Picture = LoadPicture()
        StatusBar1.Panels(1).Text = sTemp 'now restore the original text
        Call UpdateVehicleVersionAndCopyRight
        p_bChangedFlag = False ' JAW 2000.05.07 reset flag, all changes are now saved
        ' Display the save icon in the status bar
    Else
        mnuSaveAs_Click
    End If
End Sub

Private Sub mnuSave_Click()
    Call SaveVehicle
End Sub

Private Sub mnuSaveAs_Click()
    ' Code to load the SaveAs common dialog and to handle the saving
    ' of the file if the user does want to save the file
    Dim Cancel As Boolean
    Dim sTemp As String
    Dim oCDLG As clsCmdlg
    
    On Error GoTo errorhandler
    Cancel = False
    Set oCDLG = New clsCmdlg
    With oCDLG
        Dim oFile As FileSystemObject
        Set oFile = New FileSystemObject
        If oFile.FolderExists(Settings.VehiclesSavePath) Then
            .InitialDir = Settings.VehiclesSavePath
        Else
            .InitialDir = App.Path
        End If
        If m_sVehicleFile <> NEW_VEHICLE_DEFAULT_FILENAME Then
            .DefaultFilename = m_sVehicleFile
        Else
            .DefaultFilename = ""
        End If
        '.DefaultExt = ".veh"
        .Filter = OPEN_SAVE_DIALOG_FILTER
        .CancelError = True
        .MultiSelect = False
        '.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    End With
    
    Cancel = oCDLG.ShowSave(Me.hwnd)
    If Not Cancel Then
        ' A fileName was selected. Add the code to save the file here
        m_sVehicleFile = oCDLG.cFileName(0)
        'remember this path
        Settings.VehiclesSavePath = ExtractPathFromFile(m_sVehicleFile)
        p_sGUID = CreateGUID 'MPJ 07/25/2000 'all "save as" vehicles are given new guid
        m_bSavedFlag = True ' the vehicle has been saved
        p_bChangedFlag = False ' JAW 2000.05.07 no unsaved changes now remain
        DoEvents
        Call SaveVehicle
        frmDesigner.Caption = oCDLG.cFileTitle(0) & App.Title & App.Major & "." & App.Minor
         ' save the name and path of this file to our Settings UDT and our File Menu
        ' todo: Call AddRecentFile(m_sVehicleFile, oCDLG.cFileTitle(0))
    End If
    Exit Sub
errorhandler:
    InfoPrint 1, "Error in mnuSaveAs_Click:  " & CStr(err.Number) & " " & err.Description
    Resume Next
End Sub


Sub ShowCustomDropDown()
    If m_oCBO Is Nothing Then
        Set m_oCBO = New clsCompList
        ' todo: Move this string into the configuration dialog... maybe even use a hidden dialog to
    ' make this configureably by me, but not by users.  Or maybe users will want to
    ' have their own versions of this text
        Call m_oCBO.SetFileName(App.Path & "\data\parts.txt")
    End If
    Call m_oCBO.ShowDropDown
End Sub

Private Sub cboComponents_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cboComponents_KeyPress(KeyAscii As Integer)
    ' dont allow the user to manually edit this box
    KeyAscii = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Static iPrevKey As Integer
    
    If KeyCode = vbUpArrow Then
        treeVehicle.AutoHScroll = True
        treeVehicle.AutoVScroll = True
    End If
    iPrevKey = KeyCode
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'todo: can i intercept mouseup globally to make sure the drag booleans are always turned off under this condition?
End Sub

Private Sub ConfigureMainSplitters()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim r As Rect
    
    ' configure main splitters
    Set oVSplitter1 = New cSplitter
    Set oHSplitter = New cSplitter
    Set oVSplitter2 = New cSplitter
    
    ' set the colors RED FOR DEBUG
    oVSplitter1.SplitterColor = vbRed
    oHSplitter.SplitterColor = vbRed
    oVSplitter2.SplitterColor = vbRed
    
    lngTop = tabVehicle.Top / Screen.TwipsPerPixelY
    lngLeft = Settings.Splitter1
    lngWidth = 5 ' 5 pixels
    lngHeight = tabVehicle.Height / Screen.TwipsPerPixelY
    
    With oVSplitter1
        .SetPosition Me, "vSplitter1", lngTop, lngLeft, lngWidth, lngHeight
        .Orientation = splitvertical
        .SetPadding (Toolbar1.Height / Screen.TwipsPerPixelY), 5, 5, (StatusBar1.Height / Screen.TwipsPerPixelY) + 5
    
        .AddControl tabVehicle, ctlTopLeft
        .AddControl tabMain.Item(0), ctlbottomright
        ' note: the hsplitter must also be added as a "child" to the VSplitter since the Vsplitter controls its width
        .AddControl oHSplitter, ctlbottomright
        .AddControl tabSub.Item(1), ctlbottomright
    End With
    
    ' VSplitter2 will always default to midway in tabmain
    GetClientRect tabMain(0).hwnd, r
    lngLeft = r.Right - r.Left
    
    With oVSplitter2
        .SetPosition tabMain.Item(0), "vSplitter2", 10, lngLeft, 5, tabMain.Item(0).Height
        Debug.Print "VSPlitter2 Left = " & tabMain(0).Width / Screen.TwipsPerPixelX / 2
        .Orientation = splitvertical
        .SetPadding 30, 15, 15, 15
    End With
    
    lngTop = Settings.HSplitter '(tabSub.item(1).Top - 60) / Screen.TwipsPerPixelY
    lngLeft = oVSplitter1.Left + oVSplitter1.Width
    lngWidth = tabSub.Item(1).Width / Screen.TwipsPerPixelX
    lngHeight = 5 ' pixels
    
    With oHSplitter
        .SetPosition Me, "hSplitter", lngTop, lngLeft, lngWidth, lngHeight
        .Orientation = splithorizontal
        .SetPadding (Toolbar1.Height / Screen.TwipsPerPixelY), 5, 5, (StatusBar1.Height / Screen.TwipsPerPixelY) + 5
        .AddControl tabMain.Item(0), ctlTopLeft
        .AddControl tabSub.Item(1), ctlbottomright
    End With
    
   lngTop = (tabMain(0).Top + 50) / Screen.TwipsPerPixelY
   lngLeft = (tabMain(0).Left + (tabMain(0).Width / 2)) / Screen.TwipsPerPixelY
   lngWidth = 5 'pixels
   lngHeight = tabMain(0).Height / Screen.TwipsPerPixelY
End Sub

Public Sub FillListViewDefaults()
    ' todo: Fill the listview with Components to start
    Debug.Print "frmDesigner.FillListViewDefaults - TODO"
End Sub

Private Sub mnuConfigureGVD_Click()
    On Error Resume Next
    frmConfigure.Show vbModal, frmDesigner
    Set frmConfigure = Nothing
    'UpdateVehicle 'todo: uncomment after UpdateVehicle routine is working again
End Sub

Private Sub mnuTextSlim_Click()
    ExportFile "Text Slim"
End Sub
Private Sub mnuText_Click()
    ExportFile "Text"
End Sub
Private Sub mnuHTML_classic_Click()
    ExportFile "Class HTML"
End Sub
Private Sub mnuHTML_Tables_Click()
    ExportFile "New HTML"
End Sub
Private Sub mnuUnitConversion_Click()
    frmUnitConverter.Show vbModal, Me
End Sub

Private Sub mnuUseSurfaceAreaTable_Click()
    On Error Resume Next
    If mnuUseSurfaceAreaTable.Checked = True Then
        frmDesigner.Toolbar1.Buttons.Item(9).value = tbrUnpressed
        frmDesigner.mnuUseSurfaceAreaTable.Checked = False
        Settings.bUseSurfaceAreaTable = False
        m_oCurrentVeh.Options.UseSurfaceAreaTable = False
    Else
        frmDesigner.Toolbar1.Buttons.Item(9).value = tbrPressed
        frmDesigner.mnuUseSurfaceAreaTable.Checked = True
        Settings.bUseSurfaceAreaTable = True
        m_oCurrentVeh.Options.UseSurfaceAreaTable = True
    End If
    'recalc all the stats
    p_bChangedFlag = True ' JAW 2000.05.07
    'UpdateVehicle 'todo: uncomment when fixed
End Sub

Private Sub mnuVersion_Click()
    MsgBox "GVD.exe v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuRecent_Click(index As Integer)
    If OpenNewVehicle Then
        If LoadVehicle(mnuRecent(index).Caption) Then
            Call setGUID
        End If
    End If
End Sub

Sub UpdateVehicleVersionAndCopyRight()
    'todo: this needs to be a command to oManager or somethhing..
    '      or actually, it needs to be apart of Vehicles.stats or some such
    'm_oCurrentVeh.Description.CopyrightDate = Format(Date, "mmmm, yyyy")
    'm_oCurrentVeh.Description.version = m_oCurrentVeh.Description.version + 0.01
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' sort the report view according to the header clicked by the user
    ListView1.SortKey = ColumnHeader.index - 1
    ListView1.Sorted = True
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ListView1_DblClick
    End If
End Sub

Private Sub ListView1_DblClick()
    Dim h As Long
    If treeVehicle.Selection <> 0 Then
        If Not ListView1.SelectedItem Is Nothing Then
            AddComponentsFromFile ListView1.SelectedItem.Key, treeVehicle.ItemData(treeVehicle.Selection), treeVehicle.Selection
        End If
    End If
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then ' Signal a Drag operation.
        If Not ListView1.SelectedItem Is Nothing Then ' make sure there is an item selected
            'GetDataTypeAndIcons ListView1.SelectedItem.Index Oct.01.2002 Obsolete since switching to XML Based Components
            ListView1.DragIcon = ListView1.SelectedItem.CreateDragImage
            ListView1.Drag vbBeginDrag ' Drag operation.
        End If
    Else
        ListView1.MousePointer = ccCustom
    End If
End Sub

Private Sub lstPropulsionSystems_ItemCheck(Item As Integer)
    Dim sCurrent As String
    ' if we are checkmarking items because we are restoring the checklist and the user is not actually
    ' clicking on items, then we dont want to call the user Select and DeSelect routines.
    If lstPropulsionSystems.Tag = CHECKLIST_STATE_RESTORE Then Exit Sub
    
    If lstPropulsionSystems.Selected(Item) Then
        Call PropulsionSelect(sCurrent, CLng(Item))
    Else
        Call PropulsionDeSelect(sCurrent, CLng(Item))
    End If
End Sub

Private Sub lstviewLinks_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lstviewLinks.SelectedItem = lstviewLinks.HitTest(x, y)
    Set m_oLstviewLinkDragItem = lstviewLinks.SelectedItem
    
    If Button = vbLeftButton Then
        If Not m_oLstviewLinkDragItem Is Nothing Then
            lstviewLinks.DragIcon = ImageList1.ListImages(2).ExtractIcon ' use extractIcon and not .Picture because the image list is not associated with this listview
            lstviewLinks.Drag vbEndDrag
            lstviewLinks.Drag vbBeginDrag
            
            m_bTreeLinksDrag = True
        End If
    End If
End Sub

Private Sub lstviewLinks_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set m_oLstviewLinkDragItem = Nothing
    m_bTreeLinksDrag = False
End Sub

Private Sub m_oCBO_Click(ByRef s As String)
    cboComponents.Text = s
    Call LoadListView(s)
    tabMain(0).Tab = 0
End Sub

Private Sub mnuAddPerformanceProfile_Click(index As Integer)
    Dim sKey As String
    Dim lngDatatype As Long
    Dim sProfileName As String
    
'    sKey = GetNextKey
'
'    'todo: this is temporary setting of tag on lstperformanceprofiles to try and
'    ' establish the "Current" performance profile
'    m_colVehicles(1).m_oCurrentVeh.ActiveCheckList = sKey
'    m_colVehicles(1).m_oCurrentVeh.ActiveCheckListType = PERFORMANCE_CHECKLIST
    
    sProfileName = InputBox("Enter a name for this performance profile:", "New Performance Profile")
    If Not IsValidFilename(sProfileName) Then
        Exit Sub
    End If
    
    Select Case index
        
        Case 0 ' wheeled
                lngDatatype = PERFORMANCEWHEEL
        Case 1 ' tracks
                lngDatatype = PERFORMANCETRACK
        Case 2 ' skids
                lngDatatype = PERFORMANCESKID
        Case 3 ' legs
                lngDatatype = PERFORMANCELEG
        Case 4 'flexibody
                lngDatatype = PERFORMANCEFLEX
        Case 5 'air
                lngDatatype = PERFORMANCEAIR
        Case 6 ' hover
                lngDatatype = PERFORMANCEHOVER
        Case 7 ' mag-lev
                lngDatatype = PERFORMANCEMAGLEV
        Case 8 ' water
                lngDatatype = PERFORMANCEWATER
        Case 9 ' submerged
                lngDatatype = PERFORMANCESUB
        Case 10 ' space
                lngDatatype = PERFORMANCESPACE
    End Select
    
'    m_oCurrentVeh.AddPerformanceProfile lngDatatype, sKey, PERFORMANCE_KEY, 1, sProfileName
'    treeVehicle.Nodes.Add PERFORMANCE_KEY, tvwChild, sKey, sProfileName, 1
'    addnewchildnode performance_key, "settings.ico",
'    treeVehicle.Nodes(sKey).Tag = lngDatatype
End Sub

Private Sub mnuClearMessages_Click()
    txtInfo = ""
End Sub

Private Sub mnuDesignCheck_Click()
    frmDesignCheck.Show vbModal, Me
    Set frmDesignCheck = Nothing
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNotes_Click()
    frmNotes.Show vbModal, frmDesigner
    Set frmNotes = Nothing
End Sub

Private Sub mnuPowerCreateNew_Click()
    Dim sProfileName As String
    Dim sKey As String
    Dim oNode As cINode
    Dim hComponent As Long
    Dim oVehicle As cVehicle
    Dim lptr As Long
    
    ' this Sub will simulate the creation of NEW Power Profile for testing purposes
    sProfileName = InputBox("Enter a name for this power configuration:", "New Power Configuration")
    If IsValidFilename(sProfileName) Then
        ' create a test power system profile
        lptr = getCurrentVehicle
        CopyMemory oNode, lptr, 4
         
        Set oVehicle = oNode
        'todo: umm... or dont we call the factory to create this?
       ' oVehicle.addProfile , sKey, POWER_PROFILE, sProfileName
        GraphVehicle treeVehicle, lptr, hComponent
       
    End If
End Sub

Private Sub mnuFuelCreateNew_Click()
    Dim sProfileName As String
    Dim sKey As String
    
    sProfileName = InputBox("Enter a name for this fuel link configuration:", "New Fuel Configuration")
    If IsValidFilename(sProfileName) Then
        sKey = GetNextKey
        m_oCurrentVeh.addProfile sKey, FUEL_PROFILE, sProfileName
        treeVehicle.Nodes.Add FUELSYSTEMS_KEY, tvwChild, sKey, sProfileName, 1
        treeVehicle.Nodes(sKey).Tag = FUELSYSTEMS_NODE
    End If
End Sub

Private Sub mnuWeaponCreateNew_Click()
    Dim sWeaponLinkName As String
    Dim sKey As String
    
    sWeaponLinkName = InputBox("Enter a name for this weapon link:", "New Weapon Link")
    If IsValidFilename(sWeaponLinkName) Then
        sKey = GetNextKey
        
        m_oCurrentVeh.addweaponlink sKey, sWeaponLinkName
        treeVehicle.Nodes.Add WEAPON_LINKS_KEY, tvwChild, sKey, sWeaponLinkName, 1
        treeVehicle.Nodes(sKey).Tag = WEAPON_LINKS_NODE
    End If
End Sub

Private Sub mnuPrint_Click()
    Call PrintRTF(rtbOutput, 700, 700, 700, 700)
End Sub

Private Sub mnuPrintSetup_Click()
    If m_oCmdlg Is Nothing Then
        Set m_oCmdlg = New clsCmdlg
    End If
    m_oCmdlg.ShowPageSetup (Me.hwnd)   'Show Page Setup dialog
End Sub

Private Sub mnuPublish_Click()
   ' Call Publish
End Sub

Private Sub mnuAbout_Click()
    Load frmSplash
    frmSplash.Show vbModal, frmDesigner
    Set frmSplash = Nothing
End Sub

Private Sub mnuRegister_click()
    Load frmCredits
    frmCredits.Tag = "register"
    frmCredits.Show vbModal, frmDesigner
    Set frmCredits = Nothing
End Sub

Private Sub PLC1_BeforePropertyChanged(ByVal index As Long, Cancel As Boolean)
    Dim i As Long
    Dim sCharacter As String
    Dim NewString As String
    Dim iData As Integer
    iData = PLC1.ItemData(index)
    
    If iData = wdText Then
        'check for use of reserved characters
        NewString = PLC1.value(index)
        If Not IsValidFilename(NewString) Then
            Cancel = True
        End If
    ' check to see that a Integer value has been added for any Number
    ElseIf iData = wdNumber Then
        'check for numbers less than 0
        If PLC1.value(index) < 0 Then
            InfoPrint 1, "This field does not accept negative numbers.  Please use positive numbers only." 'todo: need constant for this
            Cancel = True
        End If
        
        'check to make sure the Quantity field has at least a 1 value and no more than 1,000
        If PLC1.DescriptionString(index) = "Quantity" Then ' todo Need constant
            If (PLC1.value(index) < 1) Or (PLC1.value(index) > 1000) Then 'todo: need constant for 1000
                InfoPrint 1, "The Quantity field must be an integer value from 1 to 1,000." 'todo: need constant
                Cancel = True
            End If
        End If
    ElseIf iData = wdDouble Then
        'make sure the actual user enter-able number is no greater than a Single
        'which is 3.402823E+38 and no less than .001 and NO NEGATIVE numbers at all
        'actually TODO i might want to use the square root of a single as max to keep things sane?
        Dim num As Single
        num = PLC1.value(index)
        If PLC1.DescriptionString(index) = "EmptySpace" Then
            If (num > 1E+20) Or (num < 0) Then
                InfoPrint 1, "Allowed numeric range is 0 to 1.0E20"
                Cancel = True
            End If
        ElseIf PLC1.DescriptionString(index) = "UserWeight" Then 'custom component uses these and should allow 0
        'TODO: Need constants for these
        ElseIf PLC1.DescriptionString(index) = "UserCost" Then 'custom component uses these and should allow 0
        ElseIf PLC1.DescriptionString(index) = "UserVolume" Then 'custom component uses these and should allow 0
                
        'note: Commenting this out... testing out using range checking within the DEF file via "userinputurange#" and "userinputlrange#"
        ' TODO:, i need to move all the code for callbyname in the plc1_peropertychanged into here so thta i can
        ' test if we need to Cancel = true.  Since its not til callbyname attempts to change the value that I can
        ' find out if the bounds are valid and the callbyname successfully modifies the value.
        'ElseIf (num > 1E+20) Or (num < 0.001) Then
        '    InfoPrint 1, "Allowed numeric range is 0.001 to 1.0E20" 'todo: need constant for this error message
        '    Cancel = True
        End If
    End If
End Sub

Private Sub PLC1_PropertyChanged(ByVal index As Long)
   Call modProperties.PropertyChanged(index)
End Sub

Private Sub tabMain_Click(index As Integer, PreviousTab As Integer)
    Dim sKey As String
       

    Select Case tabMain(0).Tab
        Case 0
            'SetViewMode component_view
            ' call code to populate relevant controls
            ' add controls splitter
                oVSplitter2.RemoveAllControls
                oVSplitter2.AddControl ListView1, ctlTopLeft
                oVSplitter2.AddControl PLC1, ctlbottomright
                
                PLC1.Visible = True
                ListView1.Visible = True
                lstPropulsionSystems.Visible = False
                treeLinks.Visible = False
                lstviewLinks.Visible = False
                rtbOutput.Visible = False
                picVehicleImage.Visible = False
                
        Case 1 ' performance OR weapon link
            
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
                
                oVSplitter2.RemoveAllControls
                oVSplitter2.AddControl lstPropulsionSystems, ctlTopLeft
                oVSplitter2.AddControl PLC1, ctlbottomright
                
                PLC1.Visible = True
                ListView1.Visible = False
                lstPropulsionSystems.Visible = True
                treeLinks.Visible = False
                lstviewLinks.Visible = False
                rtbOutput.Visible = False
                picVehicleImage.Visible = False
           
                     
        Case 2 ' power or fuel links
        
'                If p_ActiveNode.Key <> m_oCurrentVeh.ActiveProfile Then
'                    ' then the user clicked the tabstrip directly and is
'                    ' switching modes.
'                    If m_oCurrentVeh.ActiveProfile <> "" Then
'                        ' if no performance or weapon link created, this will be empty
'                        treeVehicle.Nodes.item(m_oCurrentVeh.ActiveProfile).Selected = True
'                        SetActiveNode
'                    End If
'                End If
                
                oVSplitter2.RemoveAllControls
                oVSplitter2.AddControl treeLinks, ctlTopLeft
                oVSplitter2.AddControl lstviewLinks, ctlbottomright
                
                
                treeLinks.Visible = True
                DoEvents
                lstviewLinks.Visible = True '<-- this is the one leaving the little crap image before its resized
                PLC1.Visible = False
                ListView1.Visible = False
                lstPropulsionSystems.Visible = False
                rtbOutput.Visible = False
                
                picVehicleImage.Visible = False
                
            
        Case 3
            'SetViewMode output_view
            oVSplitter2.RemoveAllControls
                'oVSplitter2.Visible = False
                PLC1.Visible = False
                ListView1.Visible = False
                lstPropulsionSystems.Visible = False
                treeLinks.Visible = False
                lstviewLinks.Visible = False
                rtbOutput.Visible = True           '<-- Only the output RTF is visible
                rtbOutput.ZOrder
                picVehicleImage.Visible = False
                
                
                ' test print oupt for power systems
                #If DEBUG_MODE Then
                    Dim sTemp As String
    
                    'get the Power plant info
                    sTemp = createGURPSText("Text")
                    rtbOutput.Text = sTemp
                #End If
        Case 4
            'SetViewMode image_view
            oVSplitter2.RemoveAllControls
                
                'oVSplitter2.Visible = False
                PLC1.Visible = False
                ListView1.Visible = False
                lstPropulsionSystems.Visible = False
                treeLinks.Visible = False
                lstviewLinks.Visible = False
                rtbOutput.Visible = False
                picVehicleImage.Visible = True              '<-- Only the image picturebox is visible
                picVehicleImage.ZOrder
                
    End Select
    
    Call TabStrip_Resize
    Call Form_Resize
End Sub

Private Sub tabSub_Click(index As Integer, PreviousTab As Integer)
    Select Case tabSub(1).Tab
        Case 0
            lstStats.Visible = False
            txtInfo.Visible = True
            
        Case 1
            lstStats.Visible = True
            txtInfo.Visible = False
    End Select
End Sub

' todo: currently this sub cant be private since ShowProps in modProperties calls it.
' may want to move this code to a seperate module and make it public
Public Sub SetViewMode(eModeType As VIEW_MODE)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    ' configure splitters and visible controls based on the selected mode
    Select Case eModeType
        Case component_view
            tabMain(0).Tab = 0
          
        Case image_view
            tabMain(0).Tab = 4

        Case output_view
            tabMain(0).Tab = 3

        ' Weapon or Performance Profiles
        Case performance_view
            tabMain(0).Tab = 1
            
        ' Power or Fuel Links
        Case links_view
            tabMain(0).Tab = 2
    End Select
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
        Case 1 ' New Vehicle
            mnuNewVehicle_Click
           
        Case 2 ' Open
            mnuOpen_Click
        Case 3 ' Save
            SaveVehicle
        Case 5 ' Print
            ExportFile "Text"
            
            
        Case 7 'Use surface Area formula
            mnuUseSurfaceAreaTable_Click
        Case 9 'publish vehicle
            mnuPublish_Click
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
        
    Select Case ButtonMenu.Parent.index
        Case 1
            Select Case ButtonMenu.index
                Case 1
                    mnuNewVehicle_Click
                Case 2
                    'mnuBattleSuitFF_Click 'todo: these are callng paths to load files which contain this stuff already
                Case 3
                    'mnuBattleSuitBody_Click
                Case 4
                    'mnuBattleSuitTurret_Click
            End Select
        Case 5
            Select Case ButtonMenu.index
                Case 1
                    mnuText_Click
                Case 2
                    mnuHTML_classic_Click
                Case 3
                    mnuHTML_Tables_Click
            End Select
    End Select
End Sub

Private Sub treeLinks_DragDrop(Source As Control, x As Single, y As Single)
    Dim hDropNode As Long
    
    If m_oCurrentVeh.ActiveProfile = "" Then Exit Sub
    hDropNode = treeLinks.HitTest(x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY)
    If hDropNode > 0 Then
        If TypeName(Source) = "ListView" Then
            If treeLinks.ItemText(hDropNode) = CHILD_GROUP_NAME Then
                Debug.Assert m_oCurrentVeh.ActiveProfiletype > 0
                Call m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).assignconsumer(m_oLstviewLinkDragItem.Key, treeLinks.ItemData(hDropNode))
                If m_oCurrentVeh.ActiveProfiletype = POWER_PROFILE Then
                    ' this just updates the List of Available... should move that function into the Profile class really
                    ShowLinks POWER_PROFILE
                Else
                    ShowLinks FUEL_PROFILE
                End If
                m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).Show
            Else
                InfoPrint 1, "Consumers must be added to child group nodes."
            End If
        End If
    End If
    m_bTreeLinksDrag = False
End Sub

Private Sub treeLinks_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Dim hDropNode As Long
    hDropNode = treeLinks.HitTest(x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY)
    If hDropNode > 0 Then treeLinks.SelectItem(hDropNode) = True
End Sub

Private Sub treeLinks_ItemDblClk(ByVal hItem As Long)
    Dim hParent As Long
    Dim sKey As String
    
    If (treeLinks.ItemText(hItem) <> GROUP_NAME) And (treeLinks.ItemText(hItem) <> CHILD_GROUP_NAME) Then
        hParent = treeLinks.ItemParent(hItem)
        If treeLinks.ItemText(hParent) = GROUP_NAME Then
            ' should i allow Groups to exit which have no suppliers attached?
            ' its really only good for keeping children attached and switching out suppliers
            ' Simple way would be no, and if children are attached, dont let user delete last
            ' supplier of a Group node
            ' MPJ 05/27/02 - YES.  Groups can now be void of suppliers as long as consumers are attached to the Child group
            
        Else
            sKey = m_oCurrentVeh.ActiveProfile
            m_oCurrentVeh.Profiles(sKey).UnAssignConsumer treeLinks.ItemData(hParent), treeLinks.ItemData(hItem)
            ShowLinks m_oCurrentVeh.ActiveProfiletype  ' this just updates the List of Available... should move that function into the Profile class really
            m_oCurrentVeh.Profiles(sKey).Show
        End If
    End If
End Sub

Private Sub treeLinks_ItemDragged(ByVal hItemSource As Long, ByVal hItemTarget As Long, ByVal anDropEffect As Long, pbCancel As Boolean)
    Dim hParent As Long
    Dim sKey As String
    Dim lngSupplier As Long
    Dim lngOldGroup As Long
    Dim lngNewGroup As Long
    Dim hFirstChild As Long
    
    ' we need to determine if this is either a Group node, Child Group Node, Supplier or Consumer
    ' only suppliers can be moved and only onto other Groups
    ' only consumers can be moved only onto other Child Groups
    Select Case treeLinks.ItemText(hItemSource)
        Case GROUP_NAME
            pbCancel = True
            InfoPrint 1, "Group nodes cannot be moved."
        Case CHILD_GROUP_NAME
            pbCancel = True
            InfoPrint 1, "Group nodes cannot be moved."
        Case Else
            hParent = treeLinks.ItemParent(hItemSource)
            lngSupplier = treeLinks.ItemData(hItemSource)
            lngOldGroup = treeLinks.ItemData(hParent)
            lngNewGroup = treeLinks.ItemData(hItemTarget)
            If treeLinks.ItemText(hParent) = GROUP_NAME Then
                ' its a supplier, it can only be placed on another group
                If treeLinks.ItemText(hItemTarget) <> GROUP_NAME Then
                    pbCancel = True
                    InfoPrint 1, "Suppliers can only be placed on other Supply Groups"
                Else
                    'update the actual profile with the changes
                     Call m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).movepowersystem(lngSupplier, lngOldGroup, lngNewGroup)
                End If
            Else
                ' its a consumer, it can only be placed on another child group
                If treeLinks.ItemText(hItemTarget) <> CHILD_GROUP_NAME Then
                    pbCancel = True
                    InfoPrint 1, "Consumers can only be placed on other Consumer Groups"
                Else
                    ' update the actual profile with the changes
                    Call m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).MoveConsumer(lngSupplier, lngOldGroup, lngNewGroup)
                End If
            End If
            
            On Error GoTo err
            If pbCancel = False Then
                hFirstChild = treeLinks.ItemChild(hParent)
                ' more lameness on the TreeX control.  If we try to "refresh" the control by
                ' deleting all items and then rebuilding the tree, it crashes the vb runtime.
                ' Essentially, this control doesnt like for the .RemoveAllItems method to be called
                ' from within this sub.
                ' within this procedure, it crashes the vb runtime.
                'UPDATE: 07/14/02 - Even with the new Mabry Update, they simply trap a runtime error and
                'still dont allow .RemoveAllItems to be called.
                'todo: would be nice if i can think of better way of doing this...
                m_lngTimerID = SetTimer(0, 0, TIMER_DELAY, AddressOf TimerProc)
            End If
    End Select
    Exit Sub
err:
    ' if no children left under the group, an error is thrown -- delete the group
    If err.Number = -2147417848 Then
        treeLinks.RemoveItem (lngOldGroup)
        Resume Next
    Else
        Debug.Print "frmDesigner.treeLinks_ItemDragged() -- Error #" & err.Number & " - " & err.Description
    End If
End Sub

Private Sub treeLinks_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim hNode As Long
    If Button = vbLeftButton Then
        If m_bTreeLinksDrag Then
            hNode = treeLinks.HitTest(x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY)
            If hNode > 0 Then treeLinks.SelectItem(hNode) = True
        End If
    End If
End Sub

Private Sub treeLinks_RightClick(ByVal x As Single, ByVal y As Single)
    Dim hParent As Long
    Dim hItem As Long
 
    hItem = treeLinks.HitTest(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
    If hItem <= 0 Then Exit Sub
 
    On Error Resume Next
    ' right mouse click popup
    
    With m_uTreeLinksNodeData
        .hParent = treeLinks.ItemParent(hItem)
        .hwnd = hItem
        .iGroupIndex = treeLinks.ItemData(.hParent)
        .iIndex = treeLinks.ItemData(.hwnd)
        .sText = treeLinks.ItemText(.hwnd)
    End With
    
    ' determine type of node selected
    Select Case treeLinks.ItemText(hItem)
        Case GROUP_NAME
                 '      NOTE: The Visible item should always be first since
                 '      a VB error will occur if at any time, ALL popup
                 '      items are visible = False.  This can easily happen
                 '      since we are constantly hiding/showing different items
                 '     Simple way to avoid it is to order things so that the first
                 '     item is visible = true
                 Exit Sub ' nothing to do here
        Case CHILD_GROUP_NAME
            m_uTreeLinksNodeData.iGroupIndex = treeLinks.ItemData(hItem) ' group nodes use their own index
            mnuConsumerDeleteAll.Visible = True
            mnuConsumerDelete.Visible = False
            mnuSupplierAddToNewGroup.Visible = False
        Case Else
            hParent = treeLinks.ItemParent(hItem)
            Select Case treeLinks.ItemText(hParent)
                Case GROUP_NAME
                    mnuSupplierAddToNewGroup.Visible = True
                    mnuConsumerDeleteAll.Visible = False
                    mnuConsumerDelete.Visible = False
                Case CHILD_GROUP_NAME
                    mnuConsumerDelete.Visible = True
                    mnuSupplierAddToNewGroup.Visible = False
                    mnuConsumerDeleteAll.Visible = False
                Case Else
                    Debug.Print "frmDesigner:treeLinks_ItemSelect -- ItemData holds unsupported Parent Type"
                    Debug.Print "ItemData = " & treeLinks.ItemData(hItem)
                    Exit Sub
            End Select
    End Select
    PopupMenu mnuLinksPopup
End Sub

Private Sub mnuSupplierAddToNewGroup_Click()
    Dim iNewGroup As Long
    
    iNewGroup = m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).CreateNewscGroup
    Call m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).movepowersystem(m_uTreeLinksNodeData.iIndex, m_uTreeLinksNodeData.iGroupIndex, iNewGroup)
    m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).Show
End Sub

Private Sub mnuConsumerDelete_Click()
    InfoPrint 1, "HINT:  Double clicking a consumer will also remove it from the group."
    m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).UnAssignConsumer m_uTreeLinksNodeData.iGroupIndex, m_uTreeLinksNodeData.iIndex
    ShowLinks m_oCurrentVeh.ActiveProfiletype  ' this just updates the List of Available... should move that function into the Profile class really
    m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).Show
End Sub

Private Sub mnuConsumerDeleteAll_Click()
    Dim hItem As Long
    Dim iIndex As Long
    
    On Error GoTo err
    hItem = m_uTreeLinksNodeData.hwnd
    
    hItem = treeLinks.ItemChild(hItem)
    iIndex = treeLinks.ItemData(hItem) ' index should always be 1 since whenever we delete one, the next index is always re-ordered to start at 1 again
    
    Do
        m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).UnAssignConsumer m_uTreeLinksNodeData.iGroupIndex, iIndex
        hItem = treeLinks.ItemNext(hItem)
        If hItem = 0 Then Exit Do
    Loop
    
    'todo: why is this missing an exit sub before the err?
err:
    
    ShowLinks m_oCurrentVeh.ActiveProfiletype ' this just updates the List of Available... should move that function into the Profile class really
    m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).Show
End Sub

Private Sub treeVehicle_DblClick()
'    treeVehicle.Sorted = True
End Sub
Private Sub treeVehicle_ItemDraggedTo(ByVal hItemSource As Long, ByVal pTargetTree As TreeXLibCtl.ITreeX, ByVal hItemTarget As Long, ByVal anDropEffect As Long, pbCancel As Boolean)
    Dim hSrc As Long
    Dim hDest As Long
    
    hSrc = treeVehicle.ItemData(hItemSource)
    hDest = treeVehicle.ItemData(hItemTarget)
    If Not MoveExistingComponent(hSrc, hDest) Then
        pbCancel = True
    End If
End Sub
Private Sub treeVehicle_DragDrop(Source As Control, x As Single, y As Single)
    Dim h As Long
    h = treeVehicle.HitTest(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
    If h Then
        If Source.Name = "ListView1" Then
            Call AddComponentsFromFile(Source.SelectedItem.Key, treeVehicle.ItemData(h), h)
        Else
            MsgBox TypeName(Me) & ":treeVehicle_DragDrop() -- Error.  Item not a valid component type."
        End If
        p_bChangedFlag = True
    End If
End Sub
Private Sub treeVehicle_ItemSelect(ByVal hItem As Long)
    Debug.Assert hItem > 0
    InfoPrint 1, "frmDesigner.treeVehicle_ItemSelect -- Item Handle = " & treeVehicle.ItemData(hItem)
    Properties_Show treeVehicle.ItemData(hItem)
End Sub
Private Sub treeVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not treeVehicle.SelectedItem Is Nothing Then
'        If KeyCode = vbKeyDelete Then
'            mnuDelete_Click
'        End If
'    End If
End Sub
Private Sub treeVehicle_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyDelete
            Call RemoveComponent
    End Select
End Sub
Private Sub treeVehicle_KeyUp(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        '33 = PgUp, 34 = PgDn, 35 = End, 36 = Home, 37 = Left, 38 = Up , 39 = Right, 40 = Down
'        Case 33, 34, 35, 36, 38, 40, 37, 39
'            SetActiveNode
'        Case Else
'            Debug.Print "keycode = " & KeyCode
'    End Select
End Sub
Private Sub treeVehicle_RightClick(ByVal x As Single, ByVal y As Single)
    Dim h As Long
    Dim oNode As cINode
    Dim lngAttributes As Long
    Dim hNode As Long
    Dim f As Boolean
    
    h = treeVehicle.HitTest(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
    If h Then
        treeVehicle.SelectItem(h) = True
        hNode = treeVehicle.ItemData(h)
        
        CopyMemory oNode, hNode, 4
        lngAttributes = oNode.Attributes
        CopyMemory oNode, 0&, 4
                
        f = (lngAttributes And NODE_REQUIRED)
        If f Then
            mnuDelete.Visible = False
        Else
            mnuDelete.Visible = True
        End If
        f = (lngAttributes And NODE_COPYABLE)
        If f Then
            mnuCopy.Visible = True
        Else
            mnuCopy.Visible = False
        End If
        f = (lngAttributes And NODE_RENAMEABLE)
        If f Then
            mnuRename.Visible = True
        Else
            mnuRename.Visible = False
        End If
            
        ' todo:  remember, not allow delete for GroupComponent, Arm, ArmMotor, OrnithopterDrivetrain, LegDrivetrain, BattlesuitSystem, FormFittingBattleSuitSystem
        ' todo: cant save or copy wings, ornithopterdrivetrains, armor motors,etc.  these are strictly in the XML def's though i believe
        PopupMenu mnuPopup
    End If
End Sub


'/////////////////////////////////////////////////////////////////////////
'RESIZE CODE
'////////////////////////////////////////////////////////////////////////
Public Sub tabVehicle_resize()

End Sub
Public Sub TabStrip_Resize()
     PLC1.ColumnWidth = PLC1.Width / Screen.TwipsPerPixelX / 2
    ' resize splitter contained in TabMain(0)
    If Not oVSplitter2 Is Nothing Then
        oVSplitter2.ParentResized
    End If
    Call ResizeTabbedChildren
End Sub

Public Sub ResizeTabbedChildren()
    Dim r As Rect
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngHeight As Long
    Dim lngWidth As Long
     
    On Error Resume Next
    
    Call GetWindowRect(tabVehicle.hwnd, r)
    
    lngLeft = 160
    lngTop = tabVehicle.Top + 560
    lngWidth = r.Right - r.Left '- 60 '60 for right padding
    lngWidth = (lngWidth - 10) * Screen.TwipsPerPixelX
    lngHeight = r.Bottom - r.Top
    lngHeight = (lngHeight - 30) * Screen.TwipsPerPixelY
    
    ' resize the infoText in SubTab(1)
    With treeVehicle
        .Left = lngLeft
        .Top = lngTop
        .Width = lngWidth
        .Height = lngHeight
    End With
        
    Call GetWindowRect(tabSub(1).hwnd, r)
    
    lngLeft = 60
    lngTop = 360
    lngWidth = r.Right - r.Left '- 60 '60 for right padding
    lngWidth = (lngWidth - 10) * Screen.TwipsPerPixelX
    lngHeight = r.Bottom - r.Top
    lngHeight = (lngHeight - 30) * Screen.TwipsPerPixelY
    
    ' resize the infoText in SubTab(1)
    With txtInfo
        .Left = lngLeft
        .Top = lngTop
        .Width = lngWidth
        .Height = lngHeight
    End With
    
    With lstStats
        .Left = lngLeft
        .Top = lngTop
        .Width = lngWidth
        .Height = lngHeight
    End With
    
    Call GetWindowRect(tabMain(0).hwnd, r)

    lngWidth = r.Right - r.Left '- 60 '60 for right padding
    lngWidth = (lngWidth - 10) * Screen.TwipsPerPixelX
    lngHeight = r.Bottom - r.Top
    lngHeight = (lngHeight - 30) * Screen.TwipsPerPixelY
    
    With picVehicleImage
        .Left = lngLeft
        .Top = lngTop
        .Width = lngWidth
        .Height = lngHeight
    End With
    
    With rtbOutput
        .Left = lngLeft
        .Top = lngTop
        .Width = lngWidth
        .Height = lngHeight
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'JAW 2000.05.07
    'added IF/THEN to skip resizing if form was only minimized. This also
    'prevents an error where some bad formulas lead to negative parameters
    'being used in cSplitter.Move, but I have not had time to dig that deeply.
    If Me.windowstate <> vbMinimized Then
        '//assumes we are dealing in Twips
        Const RIGHT_BORDER_WIDTH = 60
        Const BOTTOM_BORDER_HEIGHT = 60
        Dim lngWidth As Long
        Dim lngHeight As Long
        
        If Me.Width < 9720 Then Me.Width = 9720 'min width is 640 pixels
        If Me.Height < 7890 Then Me.Height = 7890 'min height is 480 pixels
        oVSplitter1.ParentResized
        oHSplitter.ParentResized
   End If
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


