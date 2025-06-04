VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form frmProfiles 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Peformance Profiles"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView ListView1 
      Height          =   2955
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   5212
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Done"
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   3960
      Width           =   675
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Profile"
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Top             =   480
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Profile"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   180
      Top             =   3960
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
            Picture         =   "frmProfiles.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfiles.frx":031A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'store the item selected in the ItemClick event
'for use in the DblClick event
Private ItemClicked As ListItem


Private Sub Combo1_Click()
FillProfilesList ' add items to the Listview
End Sub

Private Sub Command1_Click()
frmNewProfile.Show vbModal, frmProfiles
End Sub

Private Sub Command3_Click()
If Combo1.Text = "" Then
Else
    If MsgBox("This action cannot be undone.  Are you sure you want to delete this profile?", vbYesNo) = vbYes Then
    
        colVehicle.Remove Combo1.Text ' remove the item from the collection
        Combo1.Text = ""
        FillProfilesCombo ' repopulate the profiles combo box
        FillProfilesList ' add items to the Listview
    End If
End If
End Sub

Private Sub Form_Load()
FillProfilesCombo ' repopulate the profiles combo box
FillProfilesList ' add items to the Listview

'other views will work, but will allow additional
'behavior over a standard listbox
ListView1.View = lvwList
'display full label
ListView1.LabelWrap = False
End Sub
   
   Private Sub ListView1_DblClick()
       'toggle icon between checked and unchecked
       'use the private ItemClicked variable set in the
       'ListView1_ItemClick event
       With ItemClicked
           If .SmallIcon = 1 Then
              .SmallIcon = 2
              colVehicle.Item(Combo1.Text).CheckItem .Key
            Else
               .SmallIcon = 1
               colVehicle.Item(Combo1.Text).UnCheckItem .Key
           End If
       End With
    End Sub
   
   Private Sub ListView1_ItemClick(ByVal Item As ListItem)
       'since there is no ItemDblClick event, save the item
       'clicked on for use in the ListView's DblClick Event
       Set ItemClicked = Item
    End Sub




