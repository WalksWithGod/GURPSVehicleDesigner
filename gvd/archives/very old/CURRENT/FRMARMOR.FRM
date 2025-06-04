VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmArmor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Armor"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7635
   Icon            =   "frmArmor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   52
      Top             =   3420
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3281
            Text            =   "Cost:"
            TextSave        =   "Cost:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3281
            Text            =   "Weight:"
            TextSave        =   "Weight:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3281
            Text            =   "Points:"
            TextSave        =   "Points:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "DR:"
            TextSave        =   "DR:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   423
            MinWidth        =   423
            Picture         =   "frmArmor.frx":000C
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   423
            MinWidth        =   423
            Picture         =   "frmArmor.frx":0106
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MouseIcon       =   "frmArmor.frx":0200
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   6000
      TabIndex        =   51
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox chkOption 
      Caption         =   "Complex Option"
      Height          =   375
      Left            =   5880
      TabIndex        =   50
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Frame Frame6 
      Caption         =   "Armor By Facing"
      Height          =   3135
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5655
      Begin VB.OptionButton Side 
         Caption         =   "Option3"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   43
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton Side 
         Caption         =   "Option3"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   42
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Side 
         Caption         =   "Option3"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   41
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton Side 
         Caption         =   "Option3"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   40
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Side 
         Caption         =   "Option3"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   39
         Top             =   1080
         Width           =   255
      End
      Begin VB.ComboBox Material 
         Height          =   315
         Index           =   0
         ItemData        =   "frmArmor.frx":021C
         Left            =   2040
         List            =   "frmArmor.frx":021E
         TabIndex        =   38
         Text            =   "None"
         Top             =   690
         Width           =   1335
      End
      Begin VB.ComboBox Quality 
         Height          =   315
         Index           =   0
         ItemData        =   "frmArmor.frx":0220
         Left            =   3480
         List            =   "frmArmor.frx":0222
         TabIndex        =   37
         Text            =   "None"
         Top             =   690
         Width           =   1335
      End
      Begin VB.TextBox DR 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   36
         Text            =   "1"
         Top             =   705
         Width           =   495
      End
      Begin VB.TextBox PD 
         Height          =   285
         Index           =   0
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "0"
         Top             =   705
         Width           =   255
      End
      Begin VB.ComboBox Material 
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   34
         Text            =   "None"
         Top             =   1050
         Width           =   1335
      End
      Begin VB.ComboBox Material 
         Height          =   315
         Index           =   2
         Left            =   2040
         TabIndex        =   33
         Text            =   "None"
         Top             =   1410
         Width           =   1335
      End
      Begin VB.ComboBox Material 
         Height          =   315
         Index           =   3
         Left            =   2040
         TabIndex        =   32
         Text            =   "None"
         Top             =   1770
         Width           =   1335
      End
      Begin VB.ComboBox Material 
         Height          =   315
         Index           =   4
         Left            =   2040
         TabIndex        =   31
         Text            =   "None"
         Top             =   2130
         Width           =   1335
      End
      Begin VB.ComboBox Material 
         Height          =   315
         Index           =   5
         Left            =   2040
         TabIndex        =   30
         Text            =   "None"
         Top             =   2490
         Width           =   1335
      End
      Begin VB.ComboBox Quality 
         Height          =   315
         Index           =   1
         Left            =   3480
         TabIndex        =   29
         Text            =   "None"
         Top             =   1050
         Width           =   1335
      End
      Begin VB.ComboBox Quality 
         Height          =   315
         Index           =   2
         Left            =   3480
         TabIndex        =   28
         Text            =   "None"
         Top             =   1410
         Width           =   1335
      End
      Begin VB.ComboBox Quality 
         Height          =   315
         Index           =   3
         Left            =   3480
         TabIndex        =   27
         Text            =   "None"
         Top             =   1770
         Width           =   1335
      End
      Begin VB.ComboBox Quality 
         Height          =   315
         Index           =   4
         Left            =   3480
         TabIndex        =   26
         Text            =   "None"
         Top             =   2130
         Width           =   1335
      End
      Begin VB.ComboBox Quality 
         Height          =   315
         Index           =   5
         Left            =   3480
         TabIndex        =   25
         Text            =   "None"
         Top             =   2490
         Width           =   1335
      End
      Begin VB.TextBox DR 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   24
         Text            =   "1"
         Top             =   1065
         Width           =   495
      End
      Begin VB.TextBox DR 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   23
         Text            =   "1"
         Top             =   1425
         Width           =   495
      End
      Begin VB.TextBox DR 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   22
         Text            =   "1"
         Top             =   1785
         Width           =   495
      End
      Begin VB.TextBox DR 
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   21
         Text            =   "1"
         Top             =   2145
         Width           =   495
      End
      Begin VB.TextBox DR 
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   20
         Text            =   "1"
         Top             =   2505
         Width           =   495
      End
      Begin VB.TextBox PD 
         Height          =   285
         Index           =   1
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0"
         Top             =   1065
         Width           =   255
      End
      Begin VB.TextBox PD 
         Height          =   285
         Index           =   2
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0"
         Top             =   1425
         Width           =   255
      End
      Begin VB.TextBox PD 
         Height          =   285
         Index           =   3
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0"
         Top             =   1785
         Width           =   255
      End
      Begin VB.TextBox PD 
         Height          =   285
         Index           =   4
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   2145
         Width           =   255
      End
      Begin VB.TextBox PD 
         Height          =   285
         Index           =   5
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   2505
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Front:"
         Height          =   195
         Left            =   495
         TabIndex        =   49
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Back:"
         Height          =   195
         Left            =   480
         TabIndex        =   48
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Top:"
         Height          =   195
         Left            =   525
         TabIndex        =   47
         Top             =   2160
         Width           =   330
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Bottom:"
         Height          =   195
         Left            =   360
         TabIndex        =   46
         Top             =   2550
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Left:"
         Height          =   195
         Left            =   540
         TabIndex        =   45
         Top             =   1470
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Right:"
         Height          =   195
         Left            =   480
         TabIndex        =   44
         Top             =   1830
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "PD"
         Height          =   255
         Left            =   4920
         TabIndex        =   14
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "DR:"
         Height          =   255
         Left            =   1500
         TabIndex        =   13
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Quality:"
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Material:"
         Height          =   255
         Left            =   2220
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "C"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Statistics"
      Height          =   1575
      Left            =   720
      TabIndex        =   2
      Top             =   4800
      Width           =   2415
      Begin VB.Label Label12 
         Caption         =   "$0.00"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "0.0lbs"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Cost:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Weight"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   510
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total PR:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   210
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmArmor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Qual(1 To 4) As Integer
Dim Tech(4 To 13) As Integer
Dim WeightMod() As Single



Private Sub chkOption_Click()
Dim j As Integer
Dim i As Integer
j = GetCurrentSide
If TypeName(colVehicle.Item((colVehicle.Item(frmArmor.Tag).Parent))) <> "clsBody" Then TurretOrSuper = True

If chkOption = False Then ' if we are doing basic option
    ' Configure the Status Bar
    With StatusBar1
       .Panels(4).Visible = True
       .Panels(5).Visible = True
       .Panels(6).Visible = True
       .Panels(3).Text = "DR Points: " & (UserDR * 6)
    End With
    ArmorSetting = "Basic"
' Gray out the appropriate boxes
        For i = 1 To 5
            If frmArmor.Side(i).Value = True Then
            Else
                frmArmor.Side(i).Value = False
                frmArmor.Side(i).Enabled = False
            End If
            frmArmor.Material(i).Text = Material(0).Text
            frmArmor.Material(i).Enabled = False
            frmArmor.Quality(i).Text = Quality(0).Text
            frmArmor.Quality(i).Enabled = False
            frmArmor.PD(i).Text = PD(0).Text
            frmArmor.PD(i).Enabled = False
            frmArmor.DR(i).Text = DR(0).Text
            'frmArmor.DR(i).Enabled = False
        Next
    If Not TurretOrSuper Then
    
        ' turn off the Side option buttons (user cant change connected)
        For i = 1 To 5
            frmArmor.Side(i).Enabled = False
            'frmArmor.DR(i).Enabled = True ' turn on the DR for Body
        Next
        Exit Sub
    Else
        ' Its a Super or Turret, disable the connected side
            'frmArmor.Side(j).Value = False
            'frmArmor.Side(j).Enabled = False
            frmArmor.Material(j).Text = "None"
            frmArmor.Material(j).Enabled = False
            frmArmor.Quality(j).Text = "None"
            frmArmor.Quality(j).Enabled = False
            frmArmor.PD(j).Text = "0"
            frmArmor.PD(j).Enabled = False
            frmArmor.DR(j).Text = "0"
            frmArmor.DR(j).Enabled = False
    End If
Else
        ' Configure the Status Bar
        With StatusBar1
            .Panels(4).Visible = False
            .Panels(5).Visible = False
            .Panels(6).Visible = False
            .Panels(3).Text = "DR: " & UserDR
        End With
        ArmorSetting = "Complex"
        For i = 1 To 5
            If Not TurretOrSuper Then
                ' turn off the Side option buttons (user cant change connected)
                frmArmor.Side(i).Enabled = False
            Else
                frmArmor.Side(i).Enabled = True
                
            End If
            frmArmor.Material(i).Enabled = True
            frmArmor.Quality(i).Enabled = True
            frmArmor.PD(i).Enabled = True
            'frmArmor.DR(i).Enabled = True
        Next
       
        If TurretOrSuper Then
        ' Its a Super or Turret, disable the connected side
            frmArmor.Side(j).Value = True
            frmArmor.Side(j).Enabled = True
            frmArmor.Material(j).Text = "None"
            frmArmor.Material(j).Enabled = False
            frmArmor.Quality(j).Text = "None"
            frmArmor.Quality(j).Enabled = False
            frmArmor.PD(j).Text = "0"
            frmArmor.PD(j).Enabled = False
            frmArmor.DR(j).Text = "0"
            frmArmor.DR(j).Enabled = False
        End If
End If
End Sub

Private Sub cmdApply_Click()
SaveArmor frmArmor.Tag

End Sub



Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdOK_Click()

SaveArmor frmArmor.Tag
Unload Me

End Sub

Private Sub Form_Load()
Dim i As Integer ' counter 1
    
    'update the Temporary TechLevel
    TempTL = colVehicle.Item("1_").TL
    If TempTL < 4 Then TempTL = 4
    
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
            

End Sub


Private Sub Material_Click(Index As Integer)
FillMaterialCombo (Index)
End Sub


Sub Quality_Click(Index As Integer)

FillQualityCombo (Index)
End Sub


Private Sub Side_Click(Index As Integer)
Dim j As Integer

    'enable the former side
    For j = 1 To 5
        frmArmor.Side(j).Enabled = True
        frmArmor.Material(j).Enabled = True
        frmArmor.Quality(j).Enabled = True
        frmArmor.PD(j).Text = "1"
        frmArmor.PD(j).Enabled = True
        frmArmor.DR(j).Enabled = True
        frmArmor.DR(j).Text = "1"
    Next
    'disable the new side
        frmArmor.Material(Index).Text = "None"
        frmArmor.Material(Index).Enabled = False
        frmArmor.Quality(Index).Text = "None"
        frmArmor.Quality(Index).Enabled = False
        frmArmor.PD(Index).Text = "0"
        frmArmor.PD(Index).Enabled = False
        frmArmor.DR(Index).Text = "0"
        frmArmor.DR(Index).Enabled = False
End Sub



Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)
If StatusBar1.Panels(4).Visible Then
    If Panel.Index = 5 Then
        UserDR = UserDR + 1
    ElseIf Panel.Index = 6 Then
        If UserDR > 1 Then
        UserDR = UserDR - 1
        End If
    End If
    If TurretOrSuper Then
    DRPoints = UserDR * 5
    StatusBar1.Panels(3).Text = "DR Points: " & (DRPoints)
    Else
    DRPoints = UserDR * 6
    StatusBar1.Panels(3).Text = "DR Points: " & (DRPoints)
    End If
StatusBar1.Panels(4).Text = "DR: " & UserDR
End If
End Sub


