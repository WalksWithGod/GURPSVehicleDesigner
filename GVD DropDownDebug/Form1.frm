VERSION 5.00
Begin VB.Form frmDesigner 
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboComponents 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   390
      Width           =   2235
   End
End
Attribute VB_Name = "frmDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents m_oCBO  As clsCompList
Attribute m_oCBO.VB_VarHelpID = -1


Sub ShowCustomDropDown()
    If m_oCBO Is Nothing Then
        Set m_oCBO = New clsCompList
        ' todo: Move this string into the configuration dialog... maybe even use a hidden dialog to
    ' make this configureably by me, but not by users.  Or maybe users will want to
    ' have their own versions of this text
        Call m_oCBO.SetFileName(App.Path & "\parts.txt")
    End If
    Call m_oCBO.ShowDropDown
End Sub


Private Sub cboComponents_Click()
Call ShowCustomDropDown
End Sub

Private Sub cboComponents_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cboComponents_KeyPress(KeyAscii As Integer)
    ' dont allow the user to manually edit this box
    KeyAscii = 0
End Sub

Private Sub Form_Load()

 ' subclass the combo box so we can create our own drop down list
 Call SetHook(cboComponents.hWnd, True)

cboComponents.Text = "Components"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SetHook(cboComponents.hWnd, False)
End Sub

