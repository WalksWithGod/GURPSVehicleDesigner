VERSION 5.00
Begin VB.Form frmSaveComponent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Component Category"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   3840
      TabIndex        =   3
      Top             =   960
      Width           =   795
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "frmSaveComponent.frx":0000
      Left            =   180
      List            =   "frmSaveComponent.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1020
      Width           =   3360
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmSaveComponent.frx":0004
      Left            =   180
      List            =   "frmSaveComponent.frx":001D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   510
      Width           =   3360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   405
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   345
      Left            =   210
      TabIndex        =   4
      Top             =   60
      Width           =   4425
   End
End
Attribute VB_Name = "frmSaveComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_click()
vbwProfiler.vbwProcIn 124
vbwProfiler.vbwExecuteLine 3365
    Label1 = "...\components\" & Combo1 & "\" & frmDesigner.treeVehicle.SelectedItem.Text & ".cmp"
vbwProfiler.vbwProcOut 124
vbwProfiler.vbwExecuteLine 3366
End Sub

Private Sub Combo2_Click()
vbwProfiler.vbwProcIn 125
    Dim sEntry As String
    Dim sFirstEntry As String
    Dim sFileName As String
    Dim i As Integer

vbwProfiler.vbwExecuteLine 3367
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 3368
    DoEvents 'make sure the combo drop down repaints

    'Clear the Combo1 items
vbwProfiler.vbwExecuteLine 3369
    Combo1.Clear
vbwProfiler.vbwExecuteLine 3370
    sFileName = Combo2.Text

    'make sure we are back in the program's install path
vbwProfiler.vbwExecuteLine 3371
    ChDir App.Path
    ' Load the combo2 with the names of the components within the selected List file
vbwProfiler.vbwExecuteLine 3372
    Open App.Path & "\lists\" & sFileName & ".txt" For Input As #1 ' Open file for input.
vbwProfiler.vbwExecuteLine 3373
    i = 1 ' intialize the counter
'vbwLine 3374:    Do While Not EOF(1) ' Loop until end of file.
    Do While vbwProfiler.vbwExecuteLine(3374) Or Not EOF(1) ' Loop until end of file.
vbwProfiler.vbwExecuteLine 3375
        Input #1, sEntry
vbwProfiler.vbwExecuteLine 3376
        If i = 1 Then
vbwProfiler.vbwExecuteLine 3377
             sFirstEntry = sEntry
        End If
vbwProfiler.vbwExecuteLine 3378 'B
vbwProfiler.vbwExecuteLine 3379
        i = i + 1
vbwProfiler.vbwExecuteLine 3380
        With Combo1
vbwProfiler.vbwExecuteLine 3381
            .AddItem sEntry
vbwProfiler.vbwExecuteLine 3382
        End With
vbwProfiler.vbwExecuteLine 3383
    Loop
vbwProfiler.vbwExecuteLine 3384
    Combo1 = sFirstEntry
vbwProfiler.vbwExecuteLine 3385
    Close #1    ' Close file.

vbwProfiler.vbwProcOut 125
vbwProfiler.vbwExecuteLine 3386
    Exit Sub
errorhandler:
vbwProfiler.vbwExecuteLine 3387
    If err.Number = 53 Then
vbwProfiler.vbwExecuteLine 3388
        InfoPrint 1, "Can't find " & Combo1.Text & " listing"
    Else
vbwProfiler.vbwExecuteLine 3389 'B
vbwProfiler.vbwExecuteLine 3390
        InfoPrint 1, "Err in combo2_click: " + err.Description
    End If
vbwProfiler.vbwExecuteLine 3391 'B
vbwProfiler.vbwExecuteLine 3392
    Close #1
vbwProfiler.vbwProcOut 125
vbwProfiler.vbwExecuteLine 3393
    Exit Sub
vbwProfiler.vbwProcOut 125
vbwProfiler.vbwExecuteLine 3394
End Sub

Private Sub Command1_Click()
'    Dim sFileName As String
'    Dim sKey As String
'
'    On Error GoTo errorhandler
'
'    If Combo1.Text = "" Then
'        MsgBox "You must select a category in both drop down boxes"
'    Else
'        '//determine the filepath based on the datatype
'        sFileName = App.Path & "\components\" & Combo1.Text & "\" & frmDesigner.treeVehicle.SelectedItem.Text & ".cmp"
'        '//make sure the path exists
'        MkDir App.Path & "\components"
'        MkDir App.Path & "\components\" & Combo1.Text
'
'        '//check that the file doesnt already exist
'        If Dir(sFileName) <> "" Then
'            If MsgBox("File already exists.  Overwrite?", vbYesNo) = vbNo Then
'                Exit Sub
'            End If
'        End If
'
'
'        Call SaveComponent(frmDesigner.treeVehicle.SelectedItem.Key, sFileName)
'
'        LoadListView frmDesigner.cboComponents
vbwProfiler.vbwProcIn 126
vbwProfiler.vbwExecuteLine 3395
        Unload Me
'    End If
'    Exit Sub
'
'errorhandler:
'    If err.Number = 75 Then Resume Next
'    Debug.Print "frmSaveComponent:Click() -- Error " & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 126
vbwProfiler.vbwExecuteLine 3396
End Sub

Private Sub Command2_Click()
vbwProfiler.vbwProcIn 127
vbwProfiler.vbwExecuteLine 3397
    Unload Me
vbwProfiler.vbwProcOut 127
vbwProfiler.vbwExecuteLine 3398
End Sub


