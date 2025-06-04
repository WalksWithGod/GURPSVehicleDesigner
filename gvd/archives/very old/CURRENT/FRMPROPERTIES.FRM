VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProperties 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Properties"
   ClientHeight    =   5625
   ClientLeft      =   10395
   ClientTop       =   3480
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3435
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ComboBox combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmProperties.frx":0442
      Left            =   1860
      List            =   "frmProperties.frx":0444
      TabIndex        =   2
      Text            =   "combo1"
      Top             =   3630
      Visible         =   0   'False
      Width           =   1230
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5580
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   9843
      _Version        =   65541
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2145
      TabIndex        =   3
      Top             =   2955
      Width           =   720
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' SetWindowPos Flags
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOZORDER = &H4
Const SWP_NOREDRAW = &H8
Const SWP_NOACTIVATE = &H10
Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOCOPYBITS = &H100
Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
' SetWindowPos() hwndInsertAfter values
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Dim iCurrentGridRow As Integer  ' hold the value of the current row

Private Sub Combo1_Click()
MSFlexGrid1.Row = iCurrentGridRow ' make sure we are on the same row the user clicked on
colVehicle.Item(Text1.Tag).PropUpdate MSFlexGrid1.Row, Combo1.Text
MSFlexGrid1 = Combo1
're-calc the vehicles statistics
CalcStats
' re-calc vehicle performance figures only
If frmPerformance.Combo1.Text <> "" Then colVehicle.Item(frmPerformance.Combo1.Text).CalcPerformance
' update the Performance window
Performance_Update
End Sub

Private Sub combo1_KeyDown(KeyCode As Integer, Shift As Integer)
' I could add code to accept only up and down arrow keys so that the user can
' scroll through the combo box
KeyCode = 0
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub combo1_LostFocus()
    Combo1.Visible = False
    Properties_Show Text1.Tag
    're-calc the vehicles statistics
    'CalcStats
    ' re-calc vehicle performance figures only
    'If frmPerformance.Combo1.Text <> "" Then colVehicle.Item(frmPerformance.Combo1.Text).CalcPerformance
    ' update the Performance window
    'Performance_Update
End Sub

Private Sub Command1_Click()
    LoadArmor Text1.Tag ' load the armor dialog using the key of the current object
End Sub

Private Sub Form_Load()

    Dim i
    Dim mydata As String
    ' This uses the API to keep the window on top
    i = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    ' set the # of rows in the grid
    MSFlexGrid1.Rows = 0
    ' Set the width of the first column
    MSFlexGrid1.ColWidth(0) = 1900
    ' set the width of the second column so that it fills up the rest of the grids width
    MSFlexGrid1.ColWidth(1) = MSFlexGrid1.Width - MSFlexGrid1.ColWidth(0) - 60
    ' set the alignment of the second colum
    MSFlexGrid1.ColAlignment(1) = 0 ' flexalignright is the constant
End Sub

Sub GridEdit(KeyAscii As Integer)
Const Combobox = 1
Const Textbox = 2
Const Command = 3
Dim Selection As Integer

Selection = colVehicle.Item(Text1.Tag).PropQuery(MSFlexGrid1.Row)
Select Case Selection
    Case Combobox
      ' Adjust the column position so that the user always edits the second colum
        If MSFlexGrid1.Col = 0 Then
            MSFlexGrid1.Col = 1
        End If
        'use correct font
        Combo1.FontName = MSFlexGrid1.FontName
        Combo1.FontSize = MSFlexGrid1.FontSize
        Select Case KeyAscii
            Case 0 To Asc(" ")
                Combo1 = MSFlexGrid1
                Combo1.SelStart = 1000
            Case Else
                Combo1 = Chr(KeyAscii)
                Combo1.SelStart = 1
        End Select
        'position the edit box
        Combo1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
        Combo1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        Combo1.Width = MSFlexGrid1.CellWidth
        'combo1.Height = MSFlexGrid1.CellHeight
        Combo1.Visible = True
        'combo1.SetFocus
    Case Textbox
        ' Adjust the column position so that the user always edits the second colum
        If MSFlexGrid1.Col = 0 Then
        MSFlexGrid1.Col = 1
        End If
        'use correct font
        Text1.FontName = MSFlexGrid1.FontName
        Text1.FontSize = MSFlexGrid1.FontSize
        Select Case KeyAscii
            Case 0 To Asc(" ")
                Text1 = MSFlexGrid1
                Text1.SelStart = 1000
            Case Else
                Text1 = Chr(KeyAscii)
                Text1.SelStart = 1
        End Select
        'position the edit box
        Text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
        Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        Text1.Width = MSFlexGrid1.CellWidth
        Text1.Height = MSFlexGrid1.CellHeight
        Text1.Visible = True
        Text1.SetFocus
    Case Command
        ' Adjust the column position so that the user always edits the second colum
        If MSFlexGrid1.Col = 0 Then
        MSFlexGrid1.Col = 1
        End If
        'position the edit box
        Command1.Caption = "Customize"
        Command1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
        Command1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        Command1.Width = MSFlexGrid1.CellWidth
        Command1.Height = MSFlexGrid1.CellHeight
        Command1.Visible = True
        Command1.SetFocus
    Case Else
End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' if the user clicked the X in the control box, hide the form instead of unloading
If UnloadMode = vbFormControlMenu Then
    ' We aren't really closing the form.. just making it invisible
    Me.Visible = False
    Cancel = True
    ' Uncheck the Properties menu item
    frmDesigner.mnuProperties.Checked = False
End If
End Sub

Private Sub MSFlexGrid1_Click()

iCurrentGridRow = MSFlexGrid1.Row ' store the current row number
' when the user tries to edit single clicking in the cell

GridEdit Asc(" ")
End Sub

Private Sub MSFlexGrid1_DblClick()
' Debug Should probably delete this if not allowing dblclick functionality
'iCurrentGridRow = MSFlexGrid1.Row ' store the current row number
' when the user tries to edit by first double clicking in the cell
'GridEdit Asc(" ")
End Sub


Private Sub MSFlexGrid1_GotFocus()
' house keeping
If Text1.Visible Then
    'MSFlexGrid1.Row = iCurrentGridRow ' make sure we are on the same row the user clicked on
    'MSFlexGrid1 = Text1
    Text1.Visible = False
    ' update the object
'    colVehicle.Item(Text1.Tag).PropUpdate MSFlexGrid1.Row, Text1.Text
ElseIf Combo1.Visible Then
        Combo1.Visible = False
Else
        Command1.Visible = False
End If
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
iCurrentGridRow = MSFlexGrid1.Row ' store the current row number
' when the user tries to edit by just starting to type instantly
GridEdit KeyAscii

End Sub

Private Sub MSFlexGrid1_LeaveCell()
' house keeping
If Text1.Visible Then
    'MSFlexGrid1.Row = iCurrentGridRow ' make sure we are on the same row the user clicked on
'    MSFlexGrid1 = Text1
    Text1.Visible = False
    ' update the object with the new property value of the cell
    colVehicle.Item(Text1.Tag).PropUpdate MSFlexGrid1.Row, Text1
ElseIf Combo1.Visible Then
        'MSFlexGrid1.Row = iCurrentGridRow ' make sure we are on the same row the user clicked on
'        MSFlexGrid1 = combo1
        Combo1.Visible = False
Else
        Command1.Visible = False
End If
End Sub


Private Sub MSFlexGrid1_LostFocus()
    're-calc the vehicles statistics
    CalcStats
    ' re-calc vehicle performance figures only
    If frmPerformance.Combo1.Text <> "" Then colVehicle.Item(frmPerformance.Combo1.Text).CalcPerformance
    ' update the Performance window
    Performance_Update
    CalcStructuralStats ' update the Vehicle's Total Structural Stats
End Sub

Private Sub Text1_Change()
MSFlexGrid1.Row = iCurrentGridRow ' make sure we are on the same row the user clicked on
MSFlexGrid1 = Text1
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

' once the text box is visible and editable, check for the following special keystrokes
Select Case KeyCode
Case vbKeyEscape
    Text1.Visible = False
    Combo1.Visible = False
    Command1.Visible = False
    MSFlexGrid1.SetFocus
Case vbKeyReturn
    MSFlexGrid1.SetFocus
Case vbKeyDown
    MSFlexGrid1.SetFocus
    DoEvents
    If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
    End If
Case vbKeyUp
    MSFlexGrid1.SetFocus
    DoEvents
    If MSFlexGrid1.Row > MSFlexGrid1.FixedRows Then
        MSFlexGrid1.Row = MSFlexGrid1.Row - 1
    End If
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    'noise suppression for when user hits Enter in a textbox
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text1_LostFocus()
    'MSFlexGrid1.Row = iCurrentGridRow ' make sure we are on the same row the user clicked on
    'MSFlexGrid1 = Text1
    Text1.Visible = False ' hide the textbox
    're-calc the vehicles statistics
    CalcStats
    ' re-calc vehicle performance figures only
    If frmPerformance.Combo1.Text <> "" Then colVehicle.Item(frmPerformance.Combo1.Text).CalcPerformance
    ' update the Performance window
    Performance_Update
    Properties_Show Text1.Tag
End Sub
