VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmNewProfile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Profile"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   2070
      TabIndex        =   3
      Top             =   5055
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1275
      TabIndex        =   2
      Top             =   5055
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   2625
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4275
      Left            =   30
      TabIndex        =   4
      Top             =   690
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   7541
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   105
      Top             =   4860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNewProfile.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNewProfile.frx":031A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   465
   End
End
Attribute VB_Name = "frmNewProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'store the item selected in the ItemClick event
'for use in the DblClick event
Private ItemClicked As ListItem
Private ItemKey As String 'holds the key of that item

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
FillNewProfilesList ' add items to the Listview

'other views will work, but will allow additional
'behavior over a standard listbox
ListView1.View = lvwList
'display full label
ListView1.LabelWrap = False
End Sub
   
   Private Sub ListView1_DblClick()
   Dim Element As ListItem
       'toggle icon between checked and unchecked
       'use the private ItemClicked variable set in the
       'ListView1_ItemClick event
       For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).SmallIcon = 1
        Next
       With ItemClicked
        If .SmallIcon = 1 Then
           .SmallIcon = 2
           ItemKey = .Key
        End If
       End With
    End Sub
   
   Private Sub ListView1_ItemClick(ByVal Item As ListItem)
       'since there is no ItemDblClick event, save the item
       'clicked on for use in the ListView's DblClick Event
       Set ItemClicked = Item
    End Sub

Private Sub Command1_Click()
Dim NewKey As String

' check to make sure a name has been entered in the textbox
If Text1.Text = "" Then
    MsgBox "You must first enter a name for this Profile"
    Exit Sub
End If

'check to make sure the Name(aka Key) is not already used
On Error GoTo ErrorHandler

' add the object
NewKey = Text1.Text
Set objPerformance = New clsPerformance
colVehicle.Add objPerformance, NewKey
colVehicle(NewKey).Key = NewKey
colVehicle(NewKey).PerformanceType = ItemKey
FillProfilesCombo ' repopulate the combo box for the profiles dialog
FillProfilesList ' add items to the Listview
Unload Me ' unload the form
Exit Sub
ErrorHandler:
If Err.Number = 457 Then  ' key already used
    MsgBox "Profile name already exist.  Please enter a unique name for this profile"
Else
    MsgBox Err.Number
End If
End Sub

