VERSION 5.00
Begin VB.Form frmNotes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Notes"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5715
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3450
      TabIndex        =   1
      Top             =   2700
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   330
      Left            =   4575
      Picture         =   "frmNotes.frx":0000
      TabIndex        =   0
      Top             =   2700
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      Height          =   2550
      Left            =   45
      MaxLength       =   600
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   5610
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p_bChangedFlag As Boolean    ' JAW 2000.05.07
Option Explicit

Private Sub Command1_Click()
Dim sKey As String
Dim lngDatatype As Long

    sKey = frmDesigner.treeVehicle.ItemData(frmDesigner.treeVehicle.Selection)
    
    
    If Me.Tag = "crewstation" Then
        'm_oCurrentVeh.Components(sKey).StationFunction = Text1.Text
    Else
        'm_oCurrentVeh.Components(sKey).comment = RTrim(Text1.Text)
    End If
    
    'm_oCurrentVeh.Components(sKey).StatsUpdate
    Call DisplayPrintOutput
    Me.Tag = "" 'must clear the tag here or else everything will be a crewstation assignment
    Unload Me
End Sub

Private Sub Command2_Click()
    Me.Tag = "" 'must clear the tag here or else everything will be a crewstation assignment
    Unload Me
End Sub

Private Sub Form_Activate()
Dim sKey As String
    sKey = frmDesigner.treeVehicle.ItemData(frmDesigner.treeVehicle.Selection)
    
    If Me.Tag = "crewstation" Then
        Me.Caption = "Crew station assignment"
        'Text1.Text = m_oCurrentVeh.Components(sKey).StationFunction
    Else
        Me.Caption = "Notes - " '& m_oCurrentVeh.Components(sKey).CustomDescription
        'Text1.Text = m_oCurrentVeh.Components(sKey).comment
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Me.Tag = "" 'must clear the tag here or else everything will be a crewstation assignment
                p_bChangedFlag = False ' JAW 2000.05.07
            Unload Me
            Exit Sub
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If p_bChangedFlag = True Then    ' JAW 2000.05.07
        p_bChangedFlag = True
    End If

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not IsValidKeyCode(KeyAscii) Then
        KeyAscii = 0
        Exit Sub
    End If
    p_bChangedFlag = True  ' JAW 2000.05.07

End Sub
