VERSION 5.00
Begin VB.Form frmHardpoints 
   Caption         =   "Add Hardpoints"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   480
      TabIndex        =   9
      Text            =   "Combo2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblWeight 
      Caption         =   "Weight:"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblCost 
      Caption         =   "Cost:"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Load Capacity (in pounds)"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Location"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   $"frmHardpoints.frx":0000
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   3375
   End
End
Attribute VB_Name = "frmHardpoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Dim newkey As String

' check to see that a valid Load Capacity has been entered
If Text1.Text = "" Then
    MsgBox "You must enter a Load Capacity value"
Exit Sub
End If
' check to see that the max load for that wing/arm/body has not been exceeded

' now add the item to the collection
newkey = GetNextKey ' get a new unique key.
Set objHardpoint = New clsHardPoint
colVehicle.Add objHardpoint, newkey  ' add the new object to the collection
colVehicle(newkey).Parent = Combo1.Tag
frmDesigner.TreeView1.Nodes.Add Combo1.Tag, Child, newkey, "Hardpoint", 1, 1
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' allow user to only enter numbers
Select Case KeyCode
Case 48 To 57
    If Len(Text1) > 8 Then
        KeyCode = 0
    Else
        Text1 = Text1 + Chr(KeyCode)
    End If
Case vbKeyBack
    If Text1 <> "" Then
        Length = Len(Text1)
        Text1 = Left(Text1, Length - 1)
    End If
Case Else
    KeyCode = 0
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Form_Load()
Text1.Text = ""
Combo1.Text = "Body"
Combo2.Text = "Body"
Combo1.Tag = "1_"
'Combo2.Sorted = False 'use the invisible combo to track the key of the arms
'Combo1.Sorted = False ' make sure it doesnt reorder the list

'Fill the combo list
Combo1.AddItem "Body"
For i = 1 To colVehicle.Count
    If TypeName(colVehicle.Item(i)) = "clsWing" Then
        Combo1.AddItem "Wing"
        Combo2.AddItem "Wing"
    ElseIf TypeName(colVehicle.Item(i)) = "clsArm" Then
        Combo1.AddItem colVehicle.Item(i).Key & "Arm"
        Combo2.AddItem colVehicle.Item(i).Key
    End If
Next
End Sub


Private Sub Text1_Change()
Dim Load As Long
Dim MaxLoad As Double

Load = Val(Text1)
lblCost = "Cost:  $" & Str(0.1 * Load)
lblWeight = "Weight:  " & Str(0.05 * Load) & "lbs"


End Sub
