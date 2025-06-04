VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4770
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   7335
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   2
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   5655
      TabIndex        =   12
      Top             =   480
      Width           =   5655
      Begin VB.Frame Frame4 
         Caption         =   "Vehicle Modifications"
         Height          =   3375
         Left            =   2880
         TabIndex        =   36
         Top             =   240
         Width           =   2655
         Begin VB.OptionButton optExplosivePin 
            Caption         =   "Explosive Pin"
            Height          =   255
            Left            =   600
            TabIndex        =   55
            Top             =   2280
            Width           =   1575
         End
         Begin VB.OptionButton optStandardPin 
            Caption         =   "Standard Pin"
            Height          =   255
            Left            =   600
            TabIndex        =   54
            Top             =   2040
            Width           =   1335
         End
         Begin VB.CheckBox chkPlow 
            Caption         =   "Plow"
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CheckBox chkBulldozer 
            Caption         =   "Bulldozer Blade (TL1)"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   840
            Width           =   2055
         End
         Begin VB.CheckBox chkRam 
            Caption         =   "Ram (TL1)"
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox cboConvertible 
            Height          =   315
            ItemData        =   "frmOptions.frx":000C
            Left            =   240
            List            =   "frmOptions.frx":0019
            TabIndex        =   39
            Text            =   "None"
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CheckBox chkHitch 
            Caption         =   "Hitch"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CheckBox chkPin 
            Caption         =   "Pin"
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Convertible Top"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   2640
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Force Field Grids"
         Height          =   3375
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   2535
         Begin VB.TextBox txtForceDR 
            Height          =   285
            Left            =   600
            TabIndex        =   56
            Text            =   "100"
            Top             =   1200
            Width           =   615
         End
         Begin VB.OptionButton optVariableScreen 
            Caption         =   "Variable (TL12)"
            Height          =   375
            Left            =   600
            TabIndex        =   35
            Top             =   1800
            Width           =   1455
         End
         Begin VB.OptionButton optStandardScreen 
            Caption         =   "Standard (TL11)"
            Height          =   255
            Left            =   600
            TabIndex        =   34
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CheckBox chkDeflector 
            Caption         =   "Deflector Field (TL11)"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   600
            Width           =   1935
         End
         Begin VB.CheckBox chkForce 
            Caption         =   "Force Screen"
            Height          =   375
            Left            =   240
            TabIndex        =   32
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Added DR"
            Height          =   195
            Left            =   1320
            TabIndex        =   57
            Top             =   1320
            Width           =   750
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   1
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   5655
      TabIndex        =   11
      Top             =   480
      Width           =   5655
      Begin VB.Frame Frame1 
         Caption         =   "Concealment"
         Height          =   3375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2535
         Begin VB.ComboBox cboEmissionCloaking 
            Height          =   315
            ItemData        =   "frmOptions.frx":0034
            Left            =   360
            List            =   "frmOptions.frx":0041
            TabIndex        =   59
            Text            =   "Combo1"
            Top             =   2040
            Width           =   1575
         End
         Begin VB.ComboBox cboSoundBaffling 
            Height          =   315
            ItemData        =   "frmOptions.frx":005B
            Left            =   360
            List            =   "frmOptions.frx":0068
            TabIndex        =   23
            Text            =   "Combo2"
            Top             =   720
            Width           =   1575
         End
         Begin VB.ComboBox cboInfraredCloaking 
            Height          =   315
            ItemData        =   "frmOptions.frx":0082
            Left            =   360
            List            =   "frmOptions.frx":008F
            TabIndex        =   21
            Text            =   "Combo1"
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CheckBox chkCamouflage 
            Caption         =   "Camouflage Paint"
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label Label8 
            Caption         =   "Emission Cloaking (TL8)"
            Height          =   255
            Left            =   360
            TabIndex        =   58
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Sound Baffling (TL7)"
            Height          =   195
            Left            =   360
            TabIndex        =   22
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Infrared Cloacking (TL7)"
            Height          =   195
            Left            =   360
            TabIndex        =   20
            Top             =   1200
            Width           =   1710
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stealth"
         Height          =   3375
         Left            =   2880
         TabIndex        =   24
         Top             =   240
         Width           =   2655
         Begin VB.ComboBox cboChameleon 
            Height          =   315
            ItemData        =   "frmOptions.frx":00A9
            Left            =   240
            List            =   "frmOptions.frx":00B9
            TabIndex        =   28
            Text            =   "Combo4"
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CheckBox chkPsiShielding 
            Caption         =   "Psi Shielding (TL8)"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   2640
            Width           =   1935
         End
         Begin VB.CheckBox chkLiquidCrystal 
            Caption         =   "Liquid Crystal Skin (TL8)"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   2280
            Width           =   2175
         End
         Begin VB.ComboBox cboStealth 
            Height          =   315
            ItemData        =   "frmOptions.frx":00DD
            Left            =   240
            List            =   "frmOptions.frx":00EA
            TabIndex        =   26
            Text            =   "Combo3"
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Chameleon Systems"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   1200
            Width           =   1425
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Stealth (TL7)"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   915
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   0
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   5655
      TabIndex        =   10
      Top             =   480
      Width           =   5655
      Begin VB.OptionButton optNoCatorTri 
         Caption         =   "No Options"
         Height          =   255
         Left            =   3360
         TabIndex        =   53
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optTrimaran 
         Caption         =   "Trimaran Option"
         Height          =   255
         Left            =   3360
         TabIndex        =   52
         Top             =   1800
         Width           =   1695
      End
      Begin VB.OptionButton optCatamaran 
         Caption         =   "Catamaran Option"
         Height          =   255
         Left            =   3360
         TabIndex        =   51
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Frame fraSample1 
         Caption         =   "Hull Features"
         Height          =   3405
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2535
         Begin VB.CheckBox chkSubmarine 
            Caption         =   "Submarine Lines (TL7)"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   2160
            Width           =   1935
         End
         Begin VB.CheckBox chksealed 
            Caption         =   "Sealed (TL5)"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CheckBox chkwaterproof 
            Caption         =   "Waterproof"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkLiftingbody 
            Caption         =   "Lifting Body (TL6)"
            Height          =   375
            Left            =   240
            TabIndex        =   47
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CheckBox chkFloatation 
            Caption         =   "Floatation Hull"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   480
            Width           =   1695
         End
         Begin VB.CheckBox chkSubmersible 
            Caption         =   "Submersible Hull (TL5)"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   1560
            Width           =   2175
         End
         Begin VB.CheckBox chkRollStabilizers 
            Caption         =   "Roll Stabilizers (TL7)"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   2520
            Width           =   1935
         End
      End
      Begin VB.ComboBox cboStreamlining 
         Height          =   315
         ItemData        =   "frmOptions.frx":0104
         Left            =   3000
         List            =   "frmOptions.frx":011D
         TabIndex        =   14
         Text            =   "None"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ComboBox cboLines 
         Height          =   315
         ItemData        =   "frmOptions.frx":015C
         Left            =   3000
         List            =   "frmOptions.frx":016F
         TabIndex        =   13
         Text            =   "None"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Streamlining"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Hydrodynamic Lines"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   345
         Width           =   1815
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   60
      Top             =   4515
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6429
            Text            =   "Added Cost:"
            TextSave        =   "Added Cost:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6429
            Text            =   "Added Weight"
            TextSave        =   "Added Weight"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MouseIcon       =   "frmOptions.frx":019D
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   7
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   9
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   8
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin ComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   327680
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Body Features"
            Key             =   "Body"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Set Options for Body Features"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Surface Features"
            Key             =   "Surface"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Set Options for Surface Features"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "External Features"
            Key             =   "External"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Set Options for External Features"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmOptions.frx":01B9
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboConvertible_Click()

If cboConvertible.Text = "Hardtop" Then
    MsgBox "Reminder: Only vehicles with Six or Fewer seats, excluding Exposed seats, may have a detachable hardtop.  The Vehicle may not have top mounted Superstructures or Turrets."
ElseIf cboConvertible.Text = "Ragtop" Then
    MsgBox "Reminder: Only Vehicles with Exposed seats may have a fabric top"
End If

End Sub

Private Sub cboStealth_Click()
Dim Element As Object
If cboStealth.Text = "Radical" Then
        'Radical stealth cannot be added to vehicles with OpenMounts
        ' Triplane or Biplane wings, masts, or exposed or cycle seats
    For Each Element In colVehicle
        If TypeOf Element Is clsMast Then
            MsgBox "Radical Stealth cannot be applied to Vehicles with Masts"
            cboStealth.ListIndex = 0
            Exit For
        ElseIf TypeOf Element Is clsSeatsandStandingRoom Then
            If Element.SubType = "Exposed" Then
                MsgBox "Radical Stealth cannot be applied to Vehicles with Exposed Seats"
                cboStealth.ListIndex = 0
            ElseIf Element.SubType = "Cycle" Then
                MsgBox "Radical Stealth cannot be applied to Vehicles with Cycle Seats"
                cboStealth.ListIndex = 0
            End If
        ElseIf TypeOf Element Is clsWing Then
            If Element.SubType = "Biplane Wings" Then
                MsgBox "Radical Stealth cannot be applied to Vehicles with Biplane Wings"
                cboStealth.ListIndex = 0
            ElseIf Element.SubType = "Triplane Wings" Then
                MsgBox "Radical Stealth cannot be applied to Vehicles with Triplane Wings"
                cboStealth.ListIndex = 0
            End If
        End If
    Next
End If
End Sub

Private Sub cboStreamlining_Click()
    If chkLiftingbody = 1 Then
        Select Case cboStreamlining.Text
            Case "Fair", "Good", "None"
                MsgBox "You have enabled the 'Lifting Body' option.  This requires 'Very Good' or better streamlining."
                cboStreamlining.ListIndex = 3
        End Select
    End If

End Sub

Private Sub chkForce_Click()
If chkForce.Value Then
        txtForceDR = 100
        txtForceDR.Enabled = True
        optStandardScreen.Value = True
        optStandardScreen.Enabled = True
        If gVehicleTL < 12 Then 'this stays disabled until TL 12 or higher
        Else
        optVariableScreen.Enabled = True
        End If
Else
    optVariableScreen.Value = False
    optStandardScreen.Value = False
    optVariableScreen.Enabled = False
    optStandardScreen.Enabled = False
    txtForceDR.Enabled = False
End If
End Sub

Private Sub chkLiftingbody_Click()
If chkLiftingbody.Value = 0 Then
Exit Sub
End If

Select Case cboStreamlining.Text
Case "Fair", "Good", "None"
    MsgBox "Vehicle must have 'Very Good' or better streamlining to use the 'Lifting Body' option."
    chkLiftingbody.Value = 0
Case Else
End Select
End Sub

Private Sub chkLiquid_Click()

End Sub

Private Sub chkPsi_Click()

End Sub

Private Sub chkPin_Click()
If chkPin.Value Then
        optStandardPin.Enabled = True
        optStandardPin.Value = True
        optExplosivePin.Enabled = True
Else
    optExplosivePin.Value = False
    optStandardPin.Value = False
    optExplosivePin.Enabled = False
    optStandardPin.Enabled = False
End If
End Sub

Private Sub chksealed_Click()
If chksealed.Value = 0 Then
    If chkSubmersible.Value = 0 Then
        chkwaterproof.Enabled = True
        Exit Sub
    End If
End If
chkwaterproof.Value = 1
chkwaterproof.Enabled = False
End Sub

Private Sub chkSubmarine_Click()
If chkSubmarine.Value = 0 Then
    cboLines.Enabled = True
    Exit Sub
End If

If chkSubmersible.Value <> 1 Then
    MsgBox "A vehicle with submarine lines must have Submersible Option enabled."
    chkSubmarine.Value = 0
    Exit Sub
End If

If cboLines.Text <> "None" Then
    MsgBox "Hydrodynamic Lines must be set to 'None' if you wish to use the 'Submarine Lines' option."
    chkSubmarine.Value = 0
Else
    cboLines.Enabled = False
End If
End Sub

Private Sub chkSubmersible_Click()
If chkSubmersible.Value = 0 Then
chkwaterproof.Enabled = True
chksealed.Enabled = True
    If chksealed.Value = 1 Then
        chkwaterproof.Enabled = False
    End If
Exit Sub
End If

chkwaterproof.Value = 1
chksealed.Value = 1
chksealed.Enabled = False
chkwaterproof.Enabled = False
End Sub

Private Sub cmdApply_Click()
SaveOptions
CalcOptionsWeightandCost
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If chkFloatation.Value = 1 Then
    If chkwaterproof.Value = 0 Then
        If chksealed.Value = 0 Then
            MsgBox "A vehicle with a floation hull must be sealed OR waterproofed. (Note: Sealed vehicles are waterproofed for free.)"
            Exit Sub
        End If
    End If
End If
SaveOptions
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
    frmOptions.Tag = "Loaded"
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    'gray out values based on TL of vehicle
    If gVehicleTL < 11 Then
        chkForce.Enabled = False
        chkDeflector.Enabled = False
    End If
    If gVehicleTL < 8 Then
        cboChameleon.Enabled = False
        cboEmissionCloaking.Enabled = False
        chkLiquidCrystal.Enabled = False
        chkPsiShielding.Enabled = False
    End If
    If gVehicleTL < 7 Then
        chkSubmarine.Enabled = False
        chkRollStabilizers.Enabled = False
        cboSoundBaffling.Enabled = False
        cboInfraredCloaking.Enabled = False
        cboStealth.Enabled = False
    End If
    If gVehicleTL < 6 Then
        chkLiftingbody.Enabled = False
    End If
    If gVehicleTL < 5 Then
        chksealed.Enabled = False
        chkSubmersible.Enabled = False
    End If
    If gVehicleTL < 1 Then
        chkRam.Enabled = False
        chkBulldozer.Enabled = False
    End If
    
    'load the values for all checkboxes, combos and option buttons
    With colVehicle.Item("1_")
        chkFloatation = Abs(.FloatationHull)
        chkwaterproof = Abs(.WaterProof)
        chksealed = Abs(.Sealed)
        chkSubmersible = Abs(.Submersible)
        chkSubmarine = Abs(.SubmarineLines)
        chkRollStabilizers = Abs(.RollStabilizers)
        cboLines = .HydrodynamicLines
        If .Catamaran = False Then
            If .Trimaran = False Then
                optNoCatorTri = 1
            Else
                optCatamaran = 0
                optTrimaran = 1
            End If
        Else
            optCatamaran = 1
        End If
        cboStreamlining = .StreamLining
        chkLiftingbody = Abs(.LiftingBody)
        cboSoundBaffling = .SoundBaffling
        cboInfraredCloaking = .InfraredCloaking
        chkCamouflage = Abs(.Camouflage)
        cboEmissionCloaking = .EmissionCloaking
        cboStealth = .Stealth
        cboChameleon = .Chameleon
        chkLiquidCrystal = Abs(.LiquidCrystal)
        chkPsiShielding = Abs(.PsiShielding)
        chkDeflector = Abs(.Deflector)
        If .ForceScreen = "None" Then
            chkForce = 0
            optStandardScreen.Enabled = False
            optVariableScreen.Enabled = False
            txtForceDR.Enabled = False
        ElseIf .ForceScreen = "Standard" Then
            chkForce = 1
            optStandardScreen = 1
            txtForceDR = .ForceDR
        Else
            chkForce = 1
            optVariableScreen = 1
            txtForceDR = .ForceDR
        End If
        If .Pin = "None" Then
            chkPin = 0
            optStandardPin.Enabled = False
            optExplosivePin.Enabled = False
        ElseIf .Pin = "Standard" Then
            chkPin = 1
            optStandardPin = 1
        Else
            chkPin = 1
            optExplosivePin = 1
        End If
        chkRam = Abs(.Ram)
        chkBulldozer = Abs(.Bulldozer)
        chkPlow = Abs(.Plow)
        chkHitch = Abs(.Hitch)
        cboConvertible = .Convertible
    End With
    CalcOptionsWeightandCost ' recalc these stats and update the statusbar since they are not hardcoded
End Sub

Private Sub tbsOptions_Click()
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            Picture1(i).Left = 210
            Picture1(i).Enabled = True
        Else
            Picture1(i).Left = -20000
            Picture1(i).Enabled = False
        End If
    Next
End Sub

