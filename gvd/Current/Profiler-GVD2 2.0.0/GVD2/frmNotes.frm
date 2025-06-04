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
vbwProfiler.vbwProcIn 33
Dim sKey As String
Dim lngDatatype As Long

vbwProfiler.vbwExecuteLine 1015
    sKey = frmDesigner.treeVehicle.ItemData(frmDesigner.treeVehicle.Selection)


vbwProfiler.vbwExecuteLine 1016
    If Me.Tag = "crewstation" Then
        'm_oCurrentVeh.Components(sKey).StationFunction = Text1.Text
    Else
vbwProfiler.vbwExecuteLine 1017 'B
        'm_oCurrentVeh.Components(sKey).comment = RTrim(Text1.Text)
    End If
vbwProfiler.vbwExecuteLine 1018 'B

    'm_oCurrentVeh.Components(sKey).StatsUpdate
vbwProfiler.vbwExecuteLine 1019
    Call DisplayPrintOutput
vbwProfiler.vbwExecuteLine 1020
    Me.Tag = "" 'must clear the tag here or else everything will be a crewstation assignment
vbwProfiler.vbwExecuteLine 1021
    Unload Me
vbwProfiler.vbwProcOut 33
vbwProfiler.vbwExecuteLine 1022
End Sub

Private Sub Command2_Click()
vbwProfiler.vbwProcIn 34
vbwProfiler.vbwExecuteLine 1023
    Me.Tag = "" 'must clear the tag here or else everything will be a crewstation assignment
vbwProfiler.vbwExecuteLine 1024
    Unload Me
vbwProfiler.vbwProcOut 34
vbwProfiler.vbwExecuteLine 1025
End Sub

Private Sub Form_Activate()
vbwProfiler.vbwProcIn 35
Dim sKey As String
vbwProfiler.vbwExecuteLine 1026
    sKey = frmDesigner.treeVehicle.ItemData(frmDesigner.treeVehicle.Selection)

vbwProfiler.vbwExecuteLine 1027
    If Me.Tag = "crewstation" Then
vbwProfiler.vbwExecuteLine 1028
        Me.Caption = "Crew station assignment"
        'Text1.Text = m_oCurrentVeh.Components(sKey).StationFunction
    Else
vbwProfiler.vbwExecuteLine 1029 'B
vbwProfiler.vbwExecuteLine 1030
        Me.Caption = "Notes - " '& m_oCurrentVeh.Components(sKey).CustomDescription
        'Text1.Text = m_oCurrentVeh.Components(sKey).comment
    End If
vbwProfiler.vbwExecuteLine 1031 'B
vbwProfiler.vbwProcOut 35
vbwProfiler.vbwExecuteLine 1032
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
vbwProfiler.vbwProcIn 36
vbwProfiler.vbwExecuteLine 1033
    Select Case KeyCode
'vbwLine 1034:        Case vbKeyEscape
        Case IIf(vbwProfiler.vbwExecuteLine(1034), VBWPROFILER_EMPTY, _
        vbKeyEscape)
vbwProfiler.vbwExecuteLine 1035
            Me.Tag = "" 'must clear the tag here or else everything will be a crewstation assignment
vbwProfiler.vbwExecuteLine 1036
                p_bChangedFlag = False ' JAW 2000.05.07
vbwProfiler.vbwExecuteLine 1037
            Unload Me
vbwProfiler.vbwProcOut 36
vbwProfiler.vbwExecuteLine 1038
            Exit Sub
    End Select
vbwProfiler.vbwExecuteLine 1039 'B
vbwProfiler.vbwProcOut 36
vbwProfiler.vbwExecuteLine 1040
End Sub

Private Sub Form_Unload(Cancel As Integer)
vbwProfiler.vbwProcIn 37
vbwProfiler.vbwExecuteLine 1041
    If p_bChangedFlag = True Then    ' JAW 2000.05.07
vbwProfiler.vbwExecuteLine 1042
        p_bChangedFlag = True
    End If
vbwProfiler.vbwExecuteLine 1043 'B

vbwProfiler.vbwProcOut 37
vbwProfiler.vbwExecuteLine 1044
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
vbwProfiler.vbwProcIn 38
vbwProfiler.vbwExecuteLine 1045
    If Not IsValidKeyCode(KeyAscii) Then
vbwProfiler.vbwExecuteLine 1046
        KeyAscii = 0
vbwProfiler.vbwProcOut 38
vbwProfiler.vbwExecuteLine 1047
        Exit Sub
    End If
vbwProfiler.vbwExecuteLine 1048 'B
vbwProfiler.vbwExecuteLine 1049
    p_bChangedFlag = True  ' JAW 2000.05.07

vbwProfiler.vbwProcOut 38
vbwProfiler.vbwExecuteLine 1050
End Sub

