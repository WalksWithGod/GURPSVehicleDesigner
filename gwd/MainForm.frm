VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox NameBox 
      Height          =   375
      Left            =   240
      TabIndex        =   37
      Text            =   "Gun Name"
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox TypeBox 
      Height          =   375
      Index           =   15
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox DamBox 
      Height          =   375
      Index           =   14
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox SSBox 
      Height          =   375
      Index           =   13
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox AccBox 
      Height          =   375
      Index           =   12
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox HalfDBox 
      Height          =   375
      Index           =   11
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox MaxBox 
      Height          =   375
      Index           =   10
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox WhtBox 
      Height          =   375
      Index           =   9
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox RofBox 
      Height          =   375
      Index           =   8
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox ShotsBox 
      Height          =   375
      Index           =   7
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox StBox 
      Height          =   375
      Index           =   6
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox RclBox 
      Height          =   375
      Index           =   5
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox CostBox 
      Height          =   375
      Index           =   4
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox LCBox 
      Height          =   375
      Index           =   3
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox HldBox 
      Height          =   375
      Index           =   2
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox TlBox 
      Height          =   375
      Index           =   1
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox MalfBox 
      Height          =   375
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   6360
      Width           =   375
   End
   Begin VB.ComboBox GunType 
      Height          =   315
      ItemData        =   "MainForm.frx":0000
      Left            =   2520
      List            =   "MainForm.frx":0007
      TabIndex        =   4
      Text            =   "CPR"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox TLChooser 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Text            =   "7"
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox MagChooser 
      Height          =   315
      Index           =   2
      ItemData        =   "MainForm.frx":0010
      Left            =   120
      List            =   "MainForm.frx":0020
      TabIndex        =   2
      Text            =   "Magazine Type"
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox ActionChooser 
      Height          =   315
      Index           =   1
      ItemData        =   "MainForm.frx":0042
      Left            =   120
      List            =   "MainForm.frx":0052
      TabIndex        =   1
      Text            =   "Barrel Type"
      Top             =   120
      Width           =   2055
   End
   Begin VB.ComboBox StockChooser 
      Height          =   315
      Index           =   0
      ItemData        =   "MainForm.frx":0091
      Left            =   120
      List            =   "MainForm.frx":00A1
      TabIndex        =   0
      Text            =   "Grip Type"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label StatROFLabel 
      Caption         =   "ROF"
      Height          =   255
      Index           =   15
      Left            =   4080
      TabIndex        =   36
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatSSlLabel 
      Caption         =   "SS"
      Height          =   255
      Index           =   14
      Left            =   1680
      TabIndex        =   34
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatHalfDLabel 
      Caption         =   "1/2 D"
      Height          =   255
      Index           =   13
      Left            =   2640
      TabIndex        =   32
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatMaxLabel 
      Caption         =   "Max"
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   30
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatTypelLabel 
      Caption         =   "Type"
      Height          =   255
      Index           =   11
      Left            =   720
      TabIndex        =   28
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatDamLabel 
      Caption         =   "Damage"
      Height          =   255
      Index           =   10
      Left            =   1200
      TabIndex        =   26
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatAcclLabel 
      Caption         =   "Acc"
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   24
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatWhtlLabel 
      Caption         =   "Wht"
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   22
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatShotsLabel 
      Caption         =   "Shots"
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   20
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatStLabel 
      Caption         =   "St"
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   18
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatRclLabel 
      Caption         =   "Rcl"
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   16
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatCostLabel 
      Caption         =   "Cost"
      Height          =   255
      Index           =   4
      Left            =   6000
      TabIndex        =   14
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatLCLabel 
      Caption         =   "LC"
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   12
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatHldLabel 
      Caption         =   "Hld"
      Height          =   255
      Index           =   2
      Left            =   6960
      TabIndex        =   10
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatTLlLabel 
      Caption         =   "TL"
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   8
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label StatMalLabel 
      Caption         =   "Malf"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   6000
      Width           =   375
   End
   Begin VB.Image MagImage 
      Height          =   1335
      Left            =   3600
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image BarrelImage 
      Height          =   1215
      Left            =   3360
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Image StockImage 
      Height          =   1575
      Left            =   720
      Top             =   1920
      Width           =   2415
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GunTl As Variant

Private Sub BarrelImage_Click()
    BarrelData.Show
End Sub

Private Sub Form_GotFocus()
    TlBox(1).Text = GunTl
    HldBox(2).Text = "0"
    LCBox(3).Text = "0"
    CostBox(4).Text = clsGun.Cost
    RclBox(5).Text = "0"
    StBox(6).Text = "0"
    ShotsBox(7).Text = clsGun.Shots
    RofBox(8).Text = clsGun.sRoF
    WhtBox(9).Text = clsGun.Weight
    MaxBox(10).Text = clsGun.MaxRange
    HalfDBox(11).Text = clsGun.halfDamage
    AccBox(12).Text = clsGun.Accuracy
    SSBox(13).Text = clsGun.SnapShot
    DamBox(14).Text = clsGun.Damage
    TypeBox(15).Text = "0"
    MalfBox(0).Text = clsGun.Malfunction
    
    
End Sub

Private Sub Form_Load()
   GunTl = 7
   TLChooser.Text = GunTl
   TlBox(1).Text = GunTl
   Load MagData
   Load GripData
   Load BarrelData
   clsGun.Class_Initialize
End Sub


Private Sub Form_Terminate()
clsGun.Class_Terminate
End Sub

Private Sub MagImage_Click()
MagData.Show
End Sub

Private Sub StockImage_Click()
GripData.Show
End Sub

Private Sub TLChooser_Change()
 GunTl = TLChooser.Text
 TlBox(1).Text = GunTl
End Sub
