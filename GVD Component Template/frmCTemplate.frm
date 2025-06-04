VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCTemplate 
   Caption         =   "GVD Component Behavior Template Builder"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmCTemplate.frx":0000
      Top             =   1290
      Width           =   3345
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8085
      Left            =   30
      TabIndex        =   0
      Top             =   660
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   14261
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "frmCTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub
