VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPerformance 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Performance"
   ClientHeight    =   2985
   ClientLeft      =   4620
   ClientTop       =   4125
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      Height          =   330
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2475
      Left            =   30
      TabIndex        =   0
      Top             =   495
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   4366
      _Version        =   65541
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmPerformance"
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


Private Sub Form_Load()
    Dim i
    Dim mydata As String
    
 ' This uses the API to keep the window on top
  i = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

    ' Set the default toolbar button to its mixed state
    'Toolbar1.Buttons.Item(1).MixedState = True
    

    ' set the # of rows in the grid
    MSFlexGrid1.Rows = 0
    
   ' set the alignment of the second colum
    MSFlexGrid1.ColAlignment(1) = 0 ' flexalignright is the constant
    
    ' set the width of the second column so that it fills up the rest of the grids width
    MSFlexGrid1.ColWidth(1) = (Screen.TwipsPerPixelX * MSFlexGrid1.Width) - MSFlexGrid1.ColWidth(0)
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' if the user clicked the X in the control box, hide the form instead of unloading
If UnloadMode = vbFormControlMenu Then
    ' We aren't really closing the form.. just making it invisible
    Me.Visible = False
    Cancel = True
    
    ' Uncheck the Performance menu item
    frmDesigner.mnuPerformance.Checked = False
End If

End Sub


