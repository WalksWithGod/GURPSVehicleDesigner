VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   975
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "Form2.frx":0000
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Text            =   "Location:"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2640
      TabIndex        =   19
      Text            =   "Responsibilities"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Text            =   "Name"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   4335
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   15
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Warning Check Dialog:"
      Height          =   5895
      Left            =   5760
      TabIndex        =   0
      Top             =   2160
      Width           =   2655
      Begin VB.Label Label13 
         Caption         =   "Total Required Crew Size:"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   5280
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "Service Crew:"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   4680
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Medic:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Scientists:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Stokers:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Auxiliary Vehicles Crew"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Paying Passengers:"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Mechanics:"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Riggers:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Loaders:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Rower:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Crew Stations:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label Label16 
      Caption         =   "Secondary Responsibility:"
      Height          =   375
      Left            =   2640
      TabIndex        =   24
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label15 
      Caption         =   "Primary Responsibility:"
      Height          =   375
      Left            =   2640
      TabIndex        =   23
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label14 
      Caption         =   "Title:"
      Height          =   255
      Left            =   2640
      TabIndex        =   22
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "perhaps use list on right so user can add ""default"" types of crew and when they add more than one, it will simply be incremented."
      Height          =   615
      Left            =   600
      TabIndex        =   17
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
