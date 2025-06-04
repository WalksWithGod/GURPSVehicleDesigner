VERSION 5.00
Object = "{FF047D84-C3F1-11D2-877E-0040055C08D9}#1.0#0"; "TreeX.OCX"
Begin VB.Form frmWizard 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rules Wizard"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Formatting Options"
      Height          =   3195
      Left            =   135
      TabIndex        =   12
      Top             =   480
      Width           =   5850
      Begin VB.CheckBox chkUseThousands 
         Caption         =   "Use thousand seperators (e.g 1,000,000.00)"
         Height          =   240
         Left            =   330
         TabIndex        =   19
         Top             =   2235
         Width           =   5145
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Display as scientific notation"
         Height          =   210
         Index           =   2
         Left            =   330
         TabIndex        =   18
         Top             =   1050
         Width           =   4680
      End
      Begin VB.CheckBox chkAppendPrefix 
         Caption         =   "Append suffix to result (e.g. lbs, sq ft, mph)"
         Height          =   255
         Left            =   330
         TabIndex        =   17
         Top             =   2625
         Width           =   4965
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Do not format"
         Height          =   210
         Index           =   0
         Left            =   345
         TabIndex        =   16
         Top             =   1515
         Width           =   1875
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   2010
         TabIndex        =   14
         Text            =   "2"
         Top             =   555
         Width           =   390
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Round result to "
         Height          =   210
         Index           =   1
         Left            =   330
         TabIndex        =   13
         Top             =   615
         Width           =   1650
      End
      Begin VB.Line Line1 
         X1              =   255
         X2              =   5640
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "places beyond the decimal"
         Height          =   225
         Index           =   0
         Left            =   2460
         TabIndex        =   15
         Top             =   600
         Width           =   2985
      End
   End
   Begin TreeXLibCtl.TreeX treeConvert 
      Height          =   3180
      Left            =   180
      TabIndex        =   9
      Top             =   435
      Width           =   5835
      _cx             =   1368926260
      _cy             =   1368921577
      BorderStyle     =   5
      BackColor       =   16777215
      ForeColor       =   0
      PicturePosition =   17
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      AutoHScroll     =   -1  'True
      AutoVScroll     =   -1  'True
   End
   Begin VB.TextBox txtRuleName 
      Height          =   375
      Left            =   165
      TabIndex        =   1
      Top             =   495
      Width           =   5775
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H80000005&
      Height          =   1575
      Left            =   180
      MouseIcon       =   "frmWizard.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   5775
      TabIndex        =   8
      Top             =   4080
      Width           =   5835
      Begin VB.VScrollBar VScroll1 
         Height          =   1575
         Left            =   5535
         TabIndex        =   11
         Top             =   -15
         Width           =   255
      End
   End
   Begin VB.ListBox lstCheckList 
      Height          =   2760
      Left            =   135
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   495
      Width           =   5835
   End
   Begin VB.CommandButton cmdWizard 
      Caption         =   "Done"
      Height          =   375
      Index           =   3
      Left            =   4785
      TabIndex        =   4
      Top             =   5850
      Width           =   1215
   End
   Begin VB.CommandButton cmdWizard 
      Caption         =   "Next  >"
      Height          =   375
      Index           =   2
      Left            =   2955
      TabIndex        =   3
      Top             =   5850
      Width           =   1215
   End
   Begin VB.CommandButton cmdWizard 
      Caption         =   "<  Back"
      Height          =   375
      Index           =   1
      Left            =   1755
      TabIndex        =   2
      Top             =   5850
      Width           =   1215
   End
   Begin VB.CommandButton cmdWizard 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   5850
      Width           =   1170
   End
   Begin VB.ListBox lstConvertTo 
      Height          =   3180
      Left            =   165
      TabIndex        =   10
      Top             =   465
      Width           =   5835
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rule description (click on the underlined value to modify it):"
      Height          =   240
      Index           =   1
      Left            =   165
      TabIndex        =   6
      Top             =   3840
      Width           =   5100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Which 'Vehicles' datatype do you want to create a rule for?"
      Height          =   240
      Index           =   0
      Left            =   165
      TabIndex        =   5
      Top             =   240
      Width           =   5085
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Enum WIZARD_MODE
    MODE_CONVERT_FROM = 0
    MODE_CONVERT_TO = 1
    MODE_CONDITIONS = 2
    MODE_EXCEPTIONS = 3
    MODE_ROUND = 4
    MODE_NAME_RULE = 5
End Enum

Const CMD_CANCEL = 0
Const CMD_BACK = 1
Const CMD_NEXT = 2
Const CMD_FINISH = 3

Private WithEvents m_oLB As cCustomListBox
Attribute m_oLB.VB_VarHelpID = -1
Private m_oRule As cRule
Private m_ptrNode As Long
Private m_lngCurrentMode As WIZARD_MODE


Private Sub Form_Unload(Cancel As Integer)
vbwProfiler.vbwProcIn 495
vbwProfiler.vbwExecuteLine 10433
    Set m_oRule = Nothing
vbwProfiler.vbwExecuteLine 10434
    Set m_oLB = Nothing
vbwProfiler.vbwProcOut 495
vbwProfiler.vbwExecuteLine 10435
End Sub

Private Sub Form_Load()
vbwProfiler.vbwProcIn 496
vbwProfiler.vbwExecuteLine 10436
    Set m_oLB = New cCustomListBox
vbwProfiler.vbwExecuteLine 10437
    m_oLB.initDisplay picDisplay.hwnd, 0, 0
vbwProfiler.vbwExecuteLine 10438
    m_oLB.TextColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
vbwProfiler.vbwExecuteLine 10439
    m_oLB.BackColor = GetSysColor(COLOR_DESKTOP)

vbwProfiler.vbwExecuteLine 10440
    lstCheckList.BackColor = GetSysColor(COLOR_DESKTOP)
vbwProfiler.vbwExecuteLine 10441
    lstCheckList.ForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
vbwProfiler.vbwExecuteLine 10442
    treeConvert.BackColor = GetSysColor(COLOR_DESKTOP)
vbwProfiler.vbwExecuteLine 10443
    treeConvert.ForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
vbwProfiler.vbwExecuteLine 10444
    picDisplay.BackColor = GetSysColor(COLOR_DESKTOP)
vbwProfiler.vbwExecuteLine 10445
    lstConvertTo.BackColor = GetSysColor(COLOR_DESKTOP)
vbwProfiler.vbwExecuteLine 10446
    lstConvertTo.ForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)

    ' the frames too
vbwProfiler.vbwExecuteLine 10447
    Frame1.BackColor = GetSysColor(COLOR_DESKTOP)
vbwProfiler.vbwExecuteLine 10448
    Frame1.ForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)

vbwProfiler.vbwExecuteLine 10449
    Line1.BorderColor = GetSysColor(COLOR_HIGHLIGHTTEXT)

    ' the optiosn and checkboxes
vbwProfiler.vbwExecuteLine 10450
    Option1(0).BackColor = GetSysColor(COLOR_DESKTOP)
vbwProfiler.vbwExecuteLine 10451
    Option1(0).ForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
vbwProfiler.vbwExecuteLine 10452
    Option1(1).BackColor = GetSysColor(COLOR_DESKTOP)
vbwProfiler.vbwExecuteLine 10453
    Option1(1).ForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
vbwProfiler.vbwExecuteLine 10454
    Option1(2).BackColor = GetSysColor(COLOR_DESKTOP)
vbwProfiler.vbwExecuteLine 10455
    Option1(2).ForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)

vbwProfiler.vbwExecuteLine 10456
    chkUseThousands.BackColor = GetSysColor(COLOR_DESKTOP)
vbwProfiler.vbwExecuteLine 10457
    chkUseThousands.ForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
vbwProfiler.vbwExecuteLine 10458
    chkAppendPrefix.BackColor = GetSysColor(COLOR_DESKTOP)
vbwProfiler.vbwExecuteLine 10459
    chkAppendPrefix.ForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
vbwProfiler.vbwExecuteLine 10460
    Label2(0).BackColor = GetSysColor(COLOR_DESKTOP)
vbwProfiler.vbwExecuteLine 10461
    Label2(0).ForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)


    ' position and size all our windows
vbwProfiler.vbwExecuteLine 10462
    lstCheckList.Top = 480
vbwProfiler.vbwExecuteLine 10463
    lstCheckList.Left = 155

vbwProfiler.vbwExecuteLine 10464
    treeConvert.Top = 480
vbwProfiler.vbwExecuteLine 10465
    treeConvert.Left = 155

vbwProfiler.vbwExecuteLine 10466
    lstConvertTo.Top = 480
vbwProfiler.vbwExecuteLine 10467
    lstConvertTo.Left = 155

vbwProfiler.vbwExecuteLine 10468
    Frame1.Top = Label1(0).Top
vbwProfiler.vbwExecuteLine 10469
    Frame1.Left = Label1(0).Left
vbwProfiler.vbwExecuteLine 10470
    Frame1.Visible = False
vbwProfiler.vbwProcOut 496
vbwProfiler.vbwExecuteLine 10471
End Sub

Private Sub Form_Activate()
    ' there are only two modes when we enter this form.  1) New rule 2) edit of existing.
vbwProfiler.vbwProcIn 497
    Dim oNode As cINode
    Dim bNewNode As Boolean

vbwProfiler.vbwExecuteLine 10472
    m_ptrNode = Me.Tag
vbwProfiler.vbwExecuteLine 10473
    CopyMemory oNode, m_ptrNode, 4
vbwProfiler.vbwExecuteLine 10474
    Set m_oRule = oNode

vbwProfiler.vbwExecuteLine 10475
    If oNode.Description = "" Then ' its a new node
vbwProfiler.vbwExecuteLine 10476
        bNewNode = True
    End If
vbwProfiler.vbwExecuteLine 10477 'B
vbwProfiler.vbwExecuteLine 10478
    CopyMemory oNode, 0&, 4

vbwProfiler.vbwExecuteLine 10479
    If bNewNode Then
vbwProfiler.vbwExecuteLine 10480
        setMode MODE_CONVERT_FROM
    Else
vbwProfiler.vbwExecuteLine 10481 'B
vbwProfiler.vbwExecuteLine 10482
        setMode MODE_CONVERT_TO
vbwProfiler.vbwExecuteLine 10483
        renderRule m_oRule, m_oLB
    End If
vbwProfiler.vbwExecuteLine 10484 'B
vbwProfiler.vbwProcOut 497
vbwProfiler.vbwExecuteLine 10485
End Sub
Private Sub cmdWizard_Click(index As Integer)
vbwProfiler.vbwProcIn 498
vbwProfiler.vbwExecuteLine 10486
    Select Case index
'vbwLine 10487:        Case CMD_CANCEL
        Case IIf(vbwProfiler.vbwExecuteLine(10487), VBWPROFILER_EMPTY, _
        CMD_CANCEL)
vbwProfiler.vbwExecuteLine 10488
            Set m_oRule = Nothing
vbwProfiler.vbwExecuteLine 10489
            Unload Me
'vbwLine 10490:        Case CMD_BACK
        Case IIf(vbwProfiler.vbwExecuteLine(10490), VBWPROFILER_EMPTY, _
        CMD_BACK)
vbwProfiler.vbwExecuteLine 10491
            setMode m_lngCurrentMode - 1
'vbwLine 10492:        Case CMD_NEXT
        Case IIf(vbwProfiler.vbwExecuteLine(10492), VBWPROFILER_EMPTY, _
        CMD_NEXT)
vbwProfiler.vbwExecuteLine 10493
            setMode m_lngCurrentMode + 1

'vbwLine 10494:        Case CMD_FINISH
        Case IIf(vbwProfiler.vbwExecuteLine(10494), VBWPROFILER_EMPTY, _
        CMD_FINISH)
            ' now verify validity of file name and no reserved name is used
vbwProfiler.vbwExecuteLine 10495
            If IsValidFilename(txtRuleName.Text) Then
vbwProfiler.vbwExecuteLine 10496
                If IsNotReservedName(txtRuleName.Text) Then
vbwProfiler.vbwExecuteLine 10497
                     m_oRule.Name = txtRuleName.Text
vbwProfiler.vbwExecuteLine 10498
                    Set m_oRule = Nothing
vbwProfiler.vbwExecuteLine 10499
                    Unload Me
                Else
vbwProfiler.vbwExecuteLine 10500 'B
vbwProfiler.vbwExecuteLine 10501
                    MsgBox "Please enter a different name:", vbInformation, "Reserved name"
                End If
vbwProfiler.vbwExecuteLine 10502 'B
            Else
vbwProfiler.vbwExecuteLine 10503 'B
vbwProfiler.vbwExecuteLine 10504
                MsgBox "Please enter a different name:", vbInformation, "Invalid name"
            End If
vbwProfiler.vbwExecuteLine 10505 'B

        Case Else
vbwProfiler.vbwExecuteLine 10506 'B
vbwProfiler.vbwExecuteLine 10507
            Debug.Assert (index <= 3) And (index >= 0)
    End Select
vbwProfiler.vbwExecuteLine 10508 'B
vbwProfiler.vbwExecuteLine 10509
    renderRule m_oRule, m_oLB
vbwProfiler.vbwProcOut 498
vbwProfiler.vbwExecuteLine 10510
End Sub
Private Sub setMode(ByVal iMode As WIZARD_MODE)
vbwProfiler.vbwProcIn 499
    Dim hFile As Long
    Dim sFile As String
    Dim s() As String
    Dim sDescription As String
    Dim sCategory As String
    Dim sLine As String
    Dim hItem As Long
    Dim hParent As Long
    Dim lngID As Long
    Dim i As Long
    Dim j As Long

    Dim lngBaseID As Long
    Dim lCount As Long
    Dim oConvert As cUnitConverter

vbwProfiler.vbwExecuteLine 10511
    m_lngCurrentMode = iMode

vbwProfiler.vbwExecuteLine 10512
    Select Case m_lngCurrentMode
'vbwLine 10513:        Case MODE_CONVERT_FROM
        Case IIf(vbwProfiler.vbwExecuteLine(10513), VBWPROFILER_EMPTY, _
        MODE_CONVERT_FROM)
vbwProfiler.vbwExecuteLine 10514
            treeConvert.Visible = True
vbwProfiler.vbwExecuteLine 10515
            lstCheckList.Visible = False
vbwProfiler.vbwExecuteLine 10516
            lstConvertTo.Visible = False
vbwProfiler.vbwExecuteLine 10517
            picDisplay.Visible = True
vbwProfiler.vbwExecuteLine 10518
            Label1(0).Visible = True
vbwProfiler.vbwExecuteLine 10519
            Label1(1).Visible = False
vbwProfiler.vbwExecuteLine 10520
            Label1(0).Caption = "Which 'Vehicles' unit type do you want to create a rule for?"
vbwProfiler.vbwExecuteLine 10521
            Label1(1).Caption = "Rule description (click on an underlined value to edit it):"
vbwProfiler.vbwExecuteLine 10522
            txtRuleName.Visible = False

            ' enable back/next/cancel/finish buttons
vbwProfiler.vbwExecuteLine 10523
            cmdWizard(CMD_NEXT).Enabled = True
vbwProfiler.vbwExecuteLine 10524
            cmdWizard(CMD_BACK).Enabled = False
vbwProfiler.vbwExecuteLine 10525
            cmdWizard(CMD_FINISH).Enabled = False

            ' reset the rule stat code and clear the list
vbwProfiler.vbwExecuteLine 10526
            m_oRule.convertFrom = -1
vbwProfiler.vbwExecuteLine 10527
            treeConvert.RemoveAllItems

            ' fill our list of gvd unit types which the user may convert "FROM" into our tree
            ' note: only show the ones with actual unit types beneath the main category
            ' (i.e. so if we have a unit category "pressure" but no actual units in it, no need to display it
vbwProfiler.vbwExecuteLine 10528
            Set oConvert = New cUnitConverter

vbwProfiler.vbwExecuteLine 10529
            lCount = oConvert.tableCount
vbwProfiler.vbwExecuteLine 10530
            For i = 0 To lCount - 1
vbwProfiler.vbwExecuteLine 10531
                lngBaseID = i * 1000
vbwProfiler.vbwExecuteLine 10532
                sCategory = oConvert.tableName(i)

                ' first add the category node
vbwProfiler.vbwExecuteLine 10533
                hParent = treeConvert.AddItem(sCategory)
vbwProfiler.vbwExecuteLine 10534
                treeConvert.ItemData(hParent) = lngBaseID

                ' now add the intrinsic types
vbwProfiler.vbwExecuteLine 10535
                For j = 0 To oConvert.getUnitCountInTable(lngBaseID) - 1
vbwProfiler.vbwExecuteLine 10536
                    lngID = lngBaseID + j
vbwProfiler.vbwExecuteLine 10537
                    If oConvert.isIntrinsicUnit(lngID) Then
vbwProfiler.vbwExecuteLine 10538
                        sDescription = oConvert.unitDescription(lngID)
vbwProfiler.vbwExecuteLine 10539
                        hItem = treeConvert.AddItem(sDescription, hParent)
vbwProfiler.vbwExecuteLine 10540
                        treeConvert.ItemData(hItem) = lngID
                    End If
vbwProfiler.vbwExecuteLine 10541 'B
vbwProfiler.vbwExecuteLine 10542
                Next
                ' if the number of children under this category node is 0, remove it since its not useable anyway
                ' if the very last parent doesnt have kids, delete it
vbwProfiler.vbwExecuteLine 10543
                hItem = treeConvert.ItemChild(hParent)
vbwProfiler.vbwExecuteLine 10544
                If hItem = 0 Then
vbwProfiler.vbwExecuteLine 10545
                     treeConvert.RemoveItem (hParent)
                End If
vbwProfiler.vbwExecuteLine 10546 'B
vbwProfiler.vbwExecuteLine 10547
            Next
vbwProfiler.vbwExecuteLine 10548
            Set oConvert = Nothing

'vbwLine 10549:        Case MODE_CONVERT_TO
        Case IIf(vbwProfiler.vbwExecuteLine(10549), VBWPROFILER_EMPTY, _
        MODE_CONVERT_TO)
vbwProfiler.vbwExecuteLine 10550
            If m_oRule.convertFrom > -1 Then
                ' hide and show relevant windows and resize them to fit properly
vbwProfiler.vbwExecuteLine 10551
                treeConvert.Visible = False
vbwProfiler.vbwExecuteLine 10552
                lstCheckList.Visible = False
vbwProfiler.vbwExecuteLine 10553
                lstConvertTo.Visible = True
vbwProfiler.vbwExecuteLine 10554
                picDisplay.Visible = True
vbwProfiler.vbwExecuteLine 10555
                Label1(0).Visible = True
vbwProfiler.vbwExecuteLine 10556
                Label1(1).Visible = False
vbwProfiler.vbwExecuteLine 10557
                Label1(0).Caption = "Convert to which unit type?"
vbwProfiler.vbwExecuteLine 10558
                Label1(1).Caption = "Rule description (click on an underlined value to edit it):"
vbwProfiler.vbwExecuteLine 10559
                txtRuleName.Visible = False
vbwProfiler.vbwExecuteLine 10560
                Frame1.Visible = False

                ' enable back/next/cancel/finish buttons
vbwProfiler.vbwExecuteLine 10561
                cmdWizard(CMD_NEXT).Enabled = True
vbwProfiler.vbwExecuteLine 10562
                cmdWizard(CMD_BACK).Enabled = True
vbwProfiler.vbwExecuteLine 10563
                cmdWizard(CMD_FINISH).Enabled = False

                ' based on what we want to convert FROM, retreive list of
                ' possible convert TO units
vbwProfiler.vbwExecuteLine 10564
                Set oConvert = New cUnitConverter
vbwProfiler.vbwExecuteLine 10565
                lngBaseID = (m_oRule.convertFrom \ 1000) * 1000
vbwProfiler.vbwExecuteLine 10566
                lCount = oConvert.getUnitCountInTable(lngBaseID)  '<-- would need to call indexfromunit code internall to get the correct table
vbwProfiler.vbwExecuteLine 10567
                For i = 0 To lCount - 1
vbwProfiler.vbwExecuteLine 10568
                    sDescription = oConvert.unitDescription(lngBaseID + i)
vbwProfiler.vbwExecuteLine 10569
                    lstConvertTo.AddItem sDescription, i
vbwProfiler.vbwExecuteLine 10570
                    lstConvertTo.ItemData(i) = lngBaseID + i
vbwProfiler.vbwExecuteLine 10571
                Next
vbwProfiler.vbwExecuteLine 10572
                Set oConvert = Nothing
            Else
vbwProfiler.vbwExecuteLine 10573 'B
vbwProfiler.vbwExecuteLine 10574
                MsgBox "No 'Vehicles' unit type selected to convert..."
vbwProfiler.vbwExecuteLine 10575
                m_lngCurrentMode = m_lngCurrentMode - 1
vbwProfiler.vbwProcOut 499
vbwProfiler.vbwExecuteLine 10576
                Exit Sub
            End If
vbwProfiler.vbwExecuteLine 10577 'B

'vbwLine 10578:        Case MODE_CONDITIONS, MODE_EXCEPTIONS
        Case IIf(vbwProfiler.vbwExecuteLine(10578), VBWPROFILER_EMPTY, _
        MODE_CONDITIONS), MODE_EXCEPTIONS
vbwProfiler.vbwExecuteLine 10579
            If m_oRule.convertTo >= 0 Then
                ' hide and show relevant windows and resize them to fit properly
vbwProfiler.vbwExecuteLine 10580
                treeConvert.Visible = False
vbwProfiler.vbwExecuteLine 10581
                lstCheckList.Visible = True
vbwProfiler.vbwExecuteLine 10582
                lstConvertTo.Visible = False
vbwProfiler.vbwExecuteLine 10583
                picDisplay.Visible = True
vbwProfiler.vbwExecuteLine 10584
                Label1(0).Visible = True
vbwProfiler.vbwExecuteLine 10585
                Label1(1).Visible = True
vbwProfiler.vbwExecuteLine 10586
                Label1(1).Caption = "Rule description (click on an underlined value to edit it):"
vbwProfiler.vbwExecuteLine 10587
                txtRuleName.Visible = False
vbwProfiler.vbwExecuteLine 10588
                Frame1.Visible = False

vbwProfiler.vbwExecuteLine 10589
                If iMode = MODE_CONDITIONS Then
vbwProfiler.vbwExecuteLine 10590
                    Label1(0).Caption = "Select conditions..."
'vbwLine 10591:                ElseIf iMode = MODE_EXCEPTIONS Then
                ElseIf vbwProfiler.vbwExecuteLine(10591) Or iMode = MODE_EXCEPTIONS Then
vbwProfiler.vbwExecuteLine 10592
                    Label1(0).Caption = "Select exceptions..."
                Else
vbwProfiler.vbwExecuteLine 10593 'B
vbwProfiler.vbwExecuteLine 10594
                    MsgBox "Invalid expression mode..."
                End If
vbwProfiler.vbwExecuteLine 10595 'B

                ' enable back/next/cancel/finish buttons
vbwProfiler.vbwExecuteLine 10596
                cmdWizard(CMD_NEXT).Enabled = True
vbwProfiler.vbwExecuteLine 10597
                cmdWizard(CMD_BACK).Enabled = True
vbwProfiler.vbwExecuteLine 10598
                cmdWizard(CMD_FINISH).Enabled = False

vbwProfiler.vbwExecuteLine 10599
                With lstCheckList
vbwProfiler.vbwExecuteLine 10600
                    .Clear
vbwProfiler.vbwExecuteLine 10601
                    .AddItem EXPRESSION_EQUAL_TO, 0
vbwProfiler.vbwExecuteLine 10602
                    .ItemData(0) = EQUAL_TO
vbwProfiler.vbwExecuteLine 10603
                    .AddItem EXPRESSION_GREATER_THAN, 1
vbwProfiler.vbwExecuteLine 10604
                    .ItemData(1) = GREATER_THAN
vbwProfiler.vbwExecuteLine 10605
                    .AddItem EXPRESSION_GREATER_THAN_OR_EQUAL_TO, 2
vbwProfiler.vbwExecuteLine 10606
                    .ItemData(2) = GREATER_THAN_OR_EQUAL_TO
vbwProfiler.vbwExecuteLine 10607
                    .AddItem EXPRESSION_LESS_THAN, 3
vbwProfiler.vbwExecuteLine 10608
                    .ItemData(3) = LESS_THAN
vbwProfiler.vbwExecuteLine 10609
                    .AddItem EXPRESSION_LESS_THAN_OR_EQUAL_TO, 4
vbwProfiler.vbwExecuteLine 10610
                    .ItemData(4) = LESS_THAN_EQUAL_TO
vbwProfiler.vbwExecuteLine 10611
                    .AddItem EXPRESSION_NOT_EQUAL_TO, 5
vbwProfiler.vbwExecuteLine 10612
                    .ItemData(5) = NOT_EQUAL_TO
vbwProfiler.vbwExecuteLine 10613
                End With

                ' checkmark those "expressions" which the user has already
                ' added to the rule.  We're assisted by the fact that an expression of
                ' a specific TYPE (either condition or exception) cannot be added
                ' twice.
                Dim uExpression As uExpression
                Dim iExpressionType As EXPRESSION_TYPE
                Dim h As Long

                Dim lRet As Long
vbwProfiler.vbwExecuteLine 10614
                If iMode = MODE_CONDITIONS Then
vbwProfiler.vbwExecuteLine 10615
                    iExpressionType = CONDITION
                Else
vbwProfiler.vbwExecuteLine 10616 'B
vbwProfiler.vbwExecuteLine 10617
                    iExpressionType = EXCEPTION
                End If
vbwProfiler.vbwExecuteLine 10618 'B

vbwProfiler.vbwExecuteLine 10619
                For i = 0 To m_oRule.expressionCount - 1
vbwProfiler.vbwExecuteLine 10620
                    h = m_oRule.getExpression(i)
vbwProfiler.vbwExecuteLine 10621
                    If h Then
vbwProfiler.vbwExecuteLine 10622
                        CopyMemory uExpression, ByVal h, LenB(uExpression)
vbwProfiler.vbwExecuteLine 10623
                        If uExpression.type = iExpressionType Then
                            ' find the checklist item with an ItemData that equals our Evalutor and select it
vbwProfiler.vbwExecuteLine 10624
                            For j = 0 To lstCheckList.ListCount - 1
vbwProfiler.vbwExecuteLine 10625
                                If lstCheckList.ItemData(j) = uExpression.evaluator Then
vbwProfiler.vbwExecuteLine 10626
                                    lstCheckList.Selected(j) = True
vbwProfiler.vbwExecuteLine 10627
                                    Exit For
                                End If
vbwProfiler.vbwExecuteLine 10628 'B
vbwProfiler.vbwExecuteLine 10629
                            Next
                        End If
vbwProfiler.vbwExecuteLine 10630 'B
                    End If
vbwProfiler.vbwExecuteLine 10631 'B
vbwProfiler.vbwExecuteLine 10632
                Next
            Else
vbwProfiler.vbwExecuteLine 10633 'B
vbwProfiler.vbwExecuteLine 10634
                MsgBox "No conversion unit selected"
vbwProfiler.vbwExecuteLine 10635
                m_lngCurrentMode = m_lngCurrentMode - 1
vbwProfiler.vbwProcOut 499
vbwProfiler.vbwExecuteLine 10636
                Exit Sub
            End If
vbwProfiler.vbwExecuteLine 10637 'B

'vbwLine 10638:        Case MODE_ROUND
        Case IIf(vbwProfiler.vbwExecuteLine(10638), VBWPROFILER_EMPTY, _
        MODE_ROUND)
            ' hide and show relevant windows and resize them to fit properly
vbwProfiler.vbwExecuteLine 10639
                treeConvert.Visible = False
vbwProfiler.vbwExecuteLine 10640
                lstCheckList.Visible = False
vbwProfiler.vbwExecuteLine 10641
                lstConvertTo.Visible = False
vbwProfiler.vbwExecuteLine 10642
                picDisplay.Visible = True
vbwProfiler.vbwExecuteLine 10643
                Label1(0).Visible = False
vbwProfiler.vbwExecuteLine 10644
                Label1(1).Visible = True
vbwProfiler.vbwExecuteLine 10645
                Frame1.Caption = "Select Formatting Options"
                'Label1(0).Caption = "Select formatting options:"
vbwProfiler.vbwExecuteLine 10646
                Label1(1).Caption = "Rule description (click on an underlined value to edit it):"
vbwProfiler.vbwExecuteLine 10647
                txtRuleName.Visible = False
vbwProfiler.vbwExecuteLine 10648
                Frame1.Visible = True

                ' enable back/next/cancel/finish buttons
vbwProfiler.vbwExecuteLine 10649
                cmdWizard(CMD_NEXT).Enabled = True
vbwProfiler.vbwExecuteLine 10650
                cmdWizard(CMD_BACK).Enabled = True
vbwProfiler.vbwExecuteLine 10651
                cmdWizard(CMD_FINISH).Enabled = False

                ' check the options, checkboxes already checked
                Dim index As Long
vbwProfiler.vbwExecuteLine 10652
                index = m_oRule.RoundType
vbwProfiler.vbwExecuteLine 10653
                Option1(index).value = True
vbwProfiler.vbwExecuteLine 10654
                If m_oRule.useThousandSeperators Then
vbwProfiler.vbwExecuteLine 10655
                    chkUseThousands.value = 1
                End If
vbwProfiler.vbwExecuteLine 10656 'B
vbwProfiler.vbwExecuteLine 10657
                If m_oRule.appendPostfix Then
vbwProfiler.vbwExecuteLine 10658
                    chkAppendPrefix.value = 1
                End If
vbwProfiler.vbwExecuteLine 10659 'B

'vbwLine 10660:        Case MODE_NAME_RULE
        Case IIf(vbwProfiler.vbwExecuteLine(10660), VBWPROFILER_EMPTY, _
        MODE_NAME_RULE)
vbwProfiler.vbwExecuteLine 10661
                treeConvert.Visible = False
vbwProfiler.vbwExecuteLine 10662
                lstCheckList.Visible = False
vbwProfiler.vbwExecuteLine 10663
                lstConvertTo.Visible = False
vbwProfiler.vbwExecuteLine 10664
                picDisplay.Visible = True
vbwProfiler.vbwExecuteLine 10665
                Label1(0).Visible = True
vbwProfiler.vbwExecuteLine 10666
                Label1(1).Visible = True
vbwProfiler.vbwExecuteLine 10667
                Label1(0).Caption = "Please specify a name for this rule:"
vbwProfiler.vbwExecuteLine 10668
                Label1(1).Caption = "Rule description (click on an underlined value to edit it):"
vbwProfiler.vbwExecuteLine 10669
                txtRuleName.Visible = True
vbwProfiler.vbwExecuteLine 10670
                txtRuleName.Text = m_oRule.Name
vbwProfiler.vbwExecuteLine 10671
                Frame1.Visible = False

                ' enable back/next/cancel/finish buttons
vbwProfiler.vbwExecuteLine 10672
                cmdWizard(CMD_NEXT).Enabled = False
vbwProfiler.vbwExecuteLine 10673
                cmdWizard(CMD_BACK).Enabled = True
vbwProfiler.vbwExecuteLine 10674
                cmdWizard(CMD_FINISH).Enabled = True

                ' new created rule will have default name of "FromUnit" to "ToUnit"
vbwProfiler.vbwExecuteLine 10675
                If m_oRule.Name = "" Then
vbwProfiler.vbwExecuteLine 10676
                    Set oConvert = New cUnitConverter
vbwProfiler.vbwExecuteLine 10677
                    txtRuleName.Text = oConvert.unitDescription(m_oRule.convertFrom) & " to " & oConvert.unitDescription(m_oRule.convertTo)
vbwProfiler.vbwExecuteLine 10678
                    Set oConvert = Nothing
                End If
vbwProfiler.vbwExecuteLine 10679 'B

                ' set focus to the txtRuleName and move the cursor to the end
vbwProfiler.vbwExecuteLine 10680
                txtRuleName.SetFocus
vbwProfiler.vbwExecuteLine 10681
                txtRuleName.SelStart = Len(txtRuleName.Text)

        Case Else
vbwProfiler.vbwExecuteLine 10682 'B
vbwProfiler.vbwExecuteLine 10683
            If iMode >= MODE_NAME_RULE Then ' cant go any higher
vbwProfiler.vbwExecuteLine 10684
                 iMode = MODE_NAME_RULE
            End If
vbwProfiler.vbwExecuteLine 10685 'B
vbwProfiler.vbwExecuteLine 10686
            If iMode <= MODE_CONVERT_FROM Then ' cant go any lower
vbwProfiler.vbwExecuteLine 10687
                 iMode = MODE_CONVERT_FROM
            End If
vbwProfiler.vbwExecuteLine 10688 'B
vbwProfiler.vbwExecuteLine 10689
            m_lngCurrentMode = iMode
    End Select
vbwProfiler.vbwExecuteLine 10690 'B
vbwProfiler.vbwProcOut 499
vbwProfiler.vbwExecuteLine 10691
End Sub


Private Sub treeConvert_ItemSelect(ByVal hItem As Long)
vbwProfiler.vbwProcIn 500
vbwProfiler.vbwExecuteLine 10692
    If treeConvert.ItemData(hItem) > 0 Then
vbwProfiler.vbwExecuteLine 10693
        m_oRule.convertFrom = treeConvert.ItemData(hItem)
vbwProfiler.vbwExecuteLine 10694
        m_oRule.Category = treeConvert.ItemText(treeConvert.ItemParent(hItem))
    Else
vbwProfiler.vbwExecuteLine 10695 'B
vbwProfiler.vbwExecuteLine 10696
        m_oRule.convertFrom = -1 'treeConvert.ItemData (hItem)
vbwProfiler.vbwExecuteLine 10697
        m_oRule.Category = ""
    End If
vbwProfiler.vbwExecuteLine 10698 'B
vbwProfiler.vbwExecuteLine 10699
    renderRule m_oRule, m_oLB
vbwProfiler.vbwProcOut 500
vbwProfiler.vbwExecuteLine 10700
End Sub

Private Sub lstCheckList_ItemCheck(Item As Integer)
    ' problem... when navigating the back/next buttons, and loading the
    ' items already checked, this procedure gets called and adds those items
    ' to the rule AGAIN!  I guess sorta a simple hack is to just check that
    ' the expression doesnt already exist before adding it
vbwProfiler.vbwProcIn 501
    Dim lngType As EXPRESSION_TYPE

vbwProfiler.vbwExecuteLine 10701
    If Not m_oRule Is Nothing Then
vbwProfiler.vbwExecuteLine 10702
        If m_lngCurrentMode = MODE_CONDITIONS Then
vbwProfiler.vbwExecuteLine 10703
            lngType = CONDITION
'vbwLine 10704:        ElseIf m_lngCurrentMode = MODE_EXCEPTIONS Then
        ElseIf vbwProfiler.vbwExecuteLine(10704) Or m_lngCurrentMode = MODE_EXCEPTIONS Then
vbwProfiler.vbwExecuteLine 10705
            lngType = EXCEPTION
        Else
vbwProfiler.vbwExecuteLine 10706 'B
vbwProfiler.vbwExecuteLine 10707
            MsgBox "lstCheckList_ItemCheck() -- Invalid wizard mode."
vbwProfiler.vbwProcOut 501
vbwProfiler.vbwExecuteLine 10708
            Exit Sub
        End If
vbwProfiler.vbwExecuteLine 10709 'B

vbwProfiler.vbwExecuteLine 10710
        If lstCheckList.Selected(Item) Then
            ' if it doesnt exist, we can add it (dont want any duplicates)
vbwProfiler.vbwExecuteLine 10711
            If Not m_oRule.expressionExists(lngType, lstCheckList.ItemData(Item)) Then
vbwProfiler.vbwExecuteLine 10712
                m_oRule.addExpression 0, lngType, lstCheckList.ItemData(Item)
             End If
vbwProfiler.vbwExecuteLine 10713 'B
        Else
vbwProfiler.vbwExecuteLine 10714 'B
            ' if it exists, delete it
vbwProfiler.vbwExecuteLine 10715
            If m_oRule.expressionExists(lngType, lstCheckList.ItemData(Item)) Then
vbwProfiler.vbwExecuteLine 10716
                m_oRule.removeExpression lngType, lstCheckList.ItemData(Item)
            End If
vbwProfiler.vbwExecuteLine 10717 'B
        End If
vbwProfiler.vbwExecuteLine 10718 'B
    End If
vbwProfiler.vbwExecuteLine 10719 'B
vbwProfiler.vbwExecuteLine 10720
    renderRule m_oRule, m_oLB
vbwProfiler.vbwProcOut 501
vbwProfiler.vbwExecuteLine 10721
End Sub

Private Sub lstConvertTo_Click()
vbwProfiler.vbwProcIn 502
vbwProfiler.vbwExecuteLine 10722
    If m_lngCurrentMode = MODE_CONVERT_TO Then
vbwProfiler.vbwExecuteLine 10723
        m_oRule.convertTo = lstConvertTo.ItemData(lstConvertTo.ListIndex)
    Else
vbwProfiler.vbwExecuteLine 10724 'B
vbwProfiler.vbwExecuteLine 10725
        MsgBox "Invalid wizard mode."
    End If
vbwProfiler.vbwExecuteLine 10726 'B
vbwProfiler.vbwExecuteLine 10727
    renderRule m_oRule, m_oLB
vbwProfiler.vbwProcOut 502
vbwProfiler.vbwExecuteLine 10728
End Sub

Private Sub Option1_Click(index As Integer)
vbwProfiler.vbwProcIn 503
    Dim id As ROUND_OPTION
    'IMPORTANT:  Note that the option1 index array matches the Enum values
vbwProfiler.vbwExecuteLine 10729
    id = index
    ' if its selected
vbwProfiler.vbwExecuteLine 10730
    If Option1(id).value = True Then
        ' update the roundType in the rule
vbwProfiler.vbwExecuteLine 10731
        m_oRule.RoundType = id

vbwProfiler.vbwExecuteLine 10732
        Select Case id
'vbwLine 10733:            Case DECIMAL_PLACES
            Case IIf(vbwProfiler.vbwExecuteLine(10733), VBWPROFILER_EMPTY, _
        DECIMAL_PLACES)
vbwProfiler.vbwExecuteLine 10734
                m_oRule.roundDigits = Val(Text1(0).Text)
vbwProfiler.vbwExecuteLine 10735
                Text1(0).Enabled = True

'vbwLine 10736:            Case SCIENTIFIC
            Case IIf(vbwProfiler.vbwExecuteLine(10736), VBWPROFILER_EMPTY, _
        SCIENTIFIC)
vbwProfiler.vbwExecuteLine 10737
                Text1(0).Enabled = False

'vbwLine 10738:            Case ROUND_NONE
            Case IIf(vbwProfiler.vbwExecuteLine(10738), VBWPROFILER_EMPTY, _
        ROUND_NONE)
vbwProfiler.vbwExecuteLine 10739
                Text1(0).Enabled = False

            Case Else
vbwProfiler.vbwExecuteLine 10740 'B
vbwProfiler.vbwExecuteLine 10741
                Debug.Print "frmWizard:Option1_Click() -- Index not a valid Enum"
        End Select
vbwProfiler.vbwExecuteLine 10742 'B
vbwProfiler.vbwExecuteLine 10743
        renderRule m_oRule, m_oLB
    End If
vbwProfiler.vbwExecuteLine 10744 'B
vbwProfiler.vbwProcOut 503
vbwProfiler.vbwExecuteLine 10745
End Sub

Private Sub chkAppendPrefix_Click()
vbwProfiler.vbwProcIn 504
vbwProfiler.vbwExecuteLine 10746
    If chkAppendPrefix.value Then
vbwProfiler.vbwExecuteLine 10747
        m_oRule.appendPostfix = True
    Else
vbwProfiler.vbwExecuteLine 10748 'B
vbwProfiler.vbwExecuteLine 10749
        m_oRule.appendPostfix = False
    End If
vbwProfiler.vbwExecuteLine 10750 'B
vbwProfiler.vbwExecuteLine 10751
    renderRule m_oRule, m_oLB
vbwProfiler.vbwProcOut 504
vbwProfiler.vbwExecuteLine 10752
End Sub

Private Sub chkUseThousands_Click()
vbwProfiler.vbwProcIn 505
vbwProfiler.vbwExecuteLine 10753
    If chkUseThousands.value Then
vbwProfiler.vbwExecuteLine 10754
        m_oRule.useThousandSeperators = True
    Else
vbwProfiler.vbwExecuteLine 10755 'B
vbwProfiler.vbwExecuteLine 10756
        m_oRule.useThousandSeperators = False
    End If
vbwProfiler.vbwExecuteLine 10757 'B
vbwProfiler.vbwExecuteLine 10758
    renderRule m_oRule, m_oLB
vbwProfiler.vbwProcOut 505
vbwProfiler.vbwExecuteLine 10759
End Sub

Private Sub picDisplay_Paint()
vbwProfiler.vbwProcIn 506
vbwProfiler.vbwExecuteLine 10760
    If Not m_oLB Is Nothing Then
vbwProfiler.vbwExecuteLine 10761
        m_oLB.Paint
    End If
vbwProfiler.vbwExecuteLine 10762 'B
vbwProfiler.vbwProcOut 506
vbwProfiler.vbwExecuteLine 10763
End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 507
    Dim lngX As Long
    Dim lngY As Long
    Dim lRet As Long
    Dim hTreeNode As Long
    Dim hObject As Long

vbwProfiler.vbwExecuteLine 10764
    lngX = x \ Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 10765
    lngY = y \ Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 10766
    lRet = m_oLB.PointInHotSpot(lngX, lngY)

vbwProfiler.vbwExecuteLine 10767
    If lRet > 0 Then
vbwProfiler.vbwExecuteLine 10768
         Call displayItemClick(lRet, m_ptrNode, m_oLB)
    End If
vbwProfiler.vbwExecuteLine 10769 'B
vbwProfiler.vbwProcOut 507
vbwProfiler.vbwExecuteLine 10770
End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 508
    Dim lngX As Long
    Dim lngY As Long
vbwProfiler.vbwExecuteLine 10771
    lngX = x \ Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 10772
    lngY = y \ Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 10773
    If m_oLB.PointInHotSpot(lngX, lngY) > 0 Then
vbwProfiler.vbwExecuteLine 10774
        picDisplay.MousePointer = vbCustom
    Else
vbwProfiler.vbwExecuteLine 10775 'B
vbwProfiler.vbwExecuteLine 10776
        picDisplay.MousePointer = 1
    End If
vbwProfiler.vbwExecuteLine 10777 'B
vbwProfiler.vbwProcOut 508
vbwProfiler.vbwExecuteLine 10778
End Sub

Private Sub m_oLB_ItemAdded(ByVal lngItemCount As Long, ByVal lngMaxVisible As Long)
  ' configure the splitter bounds based on the num of rows and max visible
    ' an event should trigger this action
vbwProfiler.vbwProcIn 509
    Dim i As Long

vbwProfiler.vbwExecuteLine 10779
    i = lngItemCount
vbwProfiler.vbwExecuteLine 10780
    If lngMaxVisible < i Then
vbwProfiler.vbwExecuteLine 10781
        VScroll1.Max = i - lngMaxVisible
    Else
vbwProfiler.vbwExecuteLine 10782 'B
vbwProfiler.vbwExecuteLine 10783
        VScroll1.Max = 0
    End If
vbwProfiler.vbwExecuteLine 10784 'B
vbwProfiler.vbwExecuteLine 10785
    VScroll1.Min = 0
vbwProfiler.vbwExecuteLine 10786
   m_oLB.scrollPosition = VScroll1.value
vbwProfiler.vbwProcOut 509
vbwProfiler.vbwExecuteLine 10787
End Sub

Private Sub txtRuleName_KeyPress(KeyAscii As Integer)
vbwProfiler.vbwProcIn 510
vbwProfiler.vbwExecuteLine 10788
    If KeyAscii = vbKeyReturn Then
vbwProfiler.vbwExecuteLine 10789
        cmdWizard_Click CMD_FINISH
    End If
vbwProfiler.vbwExecuteLine 10790 'B
vbwProfiler.vbwProcOut 510
vbwProfiler.vbwExecuteLine 10791
End Sub

Private Sub VScroll1_Change()
vbwProfiler.vbwProcIn 511
vbwProfiler.vbwExecuteLine 10792
    m_oLB.scrollPosition = VScroll1.value
vbwProfiler.vbwExecuteLine 10793
    m_oLB.RenderText
vbwProfiler.vbwProcOut 511
vbwProfiler.vbwExecuteLine 10794
End Sub

Private Sub VScroll1_Scroll()
vbwProfiler.vbwProcIn 512
vbwProfiler.vbwExecuteLine 10795
    m_oLB.scrollPosition = VScroll1.value
vbwProfiler.vbwExecuteLine 10796
    m_oLB.RenderText
vbwProfiler.vbwProcOut 512
vbwProfiler.vbwExecuteLine 10797
End Sub


