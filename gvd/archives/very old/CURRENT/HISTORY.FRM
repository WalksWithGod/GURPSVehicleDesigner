VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vehicle Settings"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "ok"
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   0
      Left            =   240
      ScaleHeight     =   4815
      ScaleWidth      =   4815
      TabIndex        =   1
      Top             =   600
      Width           =   4815
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "history.frx":0000
         Left            =   720
         List            =   "history.frx":0043
         TabIndex        =   24
         Text            =   "9"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "125 x 125 Bitmap"
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Text            =   "Terminator"
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtModel 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Text            =   "T-1000"
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Text            =   "Cyberdyne Systems"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Text            =   "Cyborg"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   2175
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2640
         Width           =   4815
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   1695
         Left            =   2880
         Picture         =   "history.frx":0091
         ScaleHeight     =   125
         ScaleMode       =   0  'User
         ScaleWidth      =   125
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Nick:"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Model:"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Manfctr:"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Class:"
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "TL:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Description:"
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   2400
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   1
      Left            =   240
      ScaleHeight     =   4815
      ScaleWidth      =   4815
      TabIndex        =   2
      Top             =   600
      Width           =   4815
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   32
         Text            =   "http://"
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   31
         Text            =   "me@here.com"
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox Text5 
         Height          =   1095
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   28
         Text            =   "history.frx":3E92
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox Text11 
         Height          =   1095
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "history.frx":3F2F
         Top             =   3480
         Width           =   3495
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Text            =   "1.0"
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "1998"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label13 
         Caption         =   "URL:"
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Email:"
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Footer:"
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Header:"
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Version:"
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Author:"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Copyright:"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   26
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   25
      Top             =   5640
      Width           =   1095
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9551
      _Version        =   327680
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Settings"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Info"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      MouseIcon       =   "history.frx":3F73
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
' make sure the required field "Nick" has been filled in
If txtName = "" Then
    MsgBox "Please enter a Nickname for your vehicle."
    Exit Sub
End If
' make sure the require field "Model" has been filled in
If txtModel = "" Then
    MsgBox "Please enter a Model for your vehicle."
    Exit Sub
End If

If frmSettings.Caption <> "New Vehicle Settings" Then
    frmDesigner.Label2.Caption = " " + txtModel.Text + " " + txtName.Text
    'save all the settings
    With colVehicle.Item("1_")
        .TL = Combo1.Text
    End With
Else
    ' Destroy any object variables that might already exist in the collection
    Set colVehicle = Nothing
    ' set the Forms new title
    frmDesigner.Caption = "untitled - Vehicle Designer"
    ' set the treeviews label2 to the Model and Name of the vehicle
    frmDesigner.Label2.Caption = "  " & frmSettings.txtModel & " " & frmSettings.txtName
    ' Clear the treeview of any nodes that might already exist
    frmDesigner.TreeView1.Nodes.Clear
    ' get unique key number
    NewKey = GetNextKey
    'Add the Root node according to the type of vehicle being created
    frmDesigner.TreeView1.Nodes.Add , , NewKey, "Body", Body
    ' add the new object to the collection
    Set objBody = New clsBody
    colVehicle.Add objBody, NewKey
    ' save all the settings
    With colVehicle.Item("1_")
        .TL = Combo1.Text
    End With
    ' new code to update the object's properties
    With colVehicle.Item(NewKey)
        .Key = frmDesigner.TreeView1.Nodes.Item(NewKey).Key
        '.Parent = frmdesigner.TreeView1.Nodes.Item(NewKey).Parent.Key ' root node doesnt have a parent right?
        .component = frmDesigner.TreeView1.Nodes.Item(NewKey).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(NewKey).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(NewKey).SelectedImage
    End With
    ' select the first node
    frmDesigner.TreeView1.Nodes(1).Selected = True 'Select the first Node
    ' Update the properties window
    frmProperties.Caption = "Properties - " & frmDesigner.TreeView1.SelectedItem.Text
    Properties_Show frmDesigner.TreeView1.SelectedItem.Key
    ' Set flags and states
    bLoadedFlag = True  ' there is a vehicle file loaded
    bSavedFlag = False  ' the vehicle has not been saved yet
    With frmDesigner
        .mnuProperties.Enabled = True
        .mnuPerformance.Enabled = True
        .mnuSave = True
        .mnuSaveAs = True
        .Toolbar1.Buttons.Item(3).Enabled = True ' save button
        .Toolbar1.Buttons.Item(5).Enabled = True ' print button
        .Toolbar1.Buttons.Item(6).Enabled = True ' print preview
    End With
End If
Unload Me
End Sub

Private Sub TabStrip1_Click()
  Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To TabStrip1.Tabs.Count - 1
        If i = TabStrip1.SelectedItem.Index - 1 Then
            Picture1(i).Left = 210
            Picture1(i).Enabled = True
        Else
            Picture1(i).Left = -20000
            Picture1(i).Enabled = False
        End If
    Next
End Sub
