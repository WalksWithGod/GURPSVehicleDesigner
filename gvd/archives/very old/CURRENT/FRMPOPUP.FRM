VERSION 5.00
Begin VB.Form frmPopup 
   Caption         =   "Form1"
   ClientHeight    =   150
   ClientLeft      =   450
   ClientTop       =   8130
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "Haettenschweiler"
      Size            =   9.75
      Charset         =   0
      Weight          =   500
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   150
   ScaleWidth      =   7185
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuDelete_Click()
' call procedure to remove nodes and collection items
RemoveNode frmDesigner.TreeView1.SelectedItem.Key

End Sub

Private Sub mnuProperties_Click()
   ' check the menu item
   frmDesigner.mnuProperties.Checked = True
   
   If frmProperties.Visible = False Then
        ' Show the properties for the selected node
        frmProperties.Visible = True
    End If
    
End Sub
