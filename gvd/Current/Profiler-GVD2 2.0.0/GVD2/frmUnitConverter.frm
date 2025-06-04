VERSION 5.00
Object = "{FF047D84-C3F1-11D2-877E-0040055C08D9}#1.0#0"; "TreeX.OCX"
Begin VB.Form frmUnitConverter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit Converter"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   ClipControls    =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H8000000D&
   Icon            =   "frmUnitConverter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmUnitConverter.frx":0442
   ScaleHeight     =   6240
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   1935
      Left            =   5715
      TabIndex        =   9
      Top             =   3735
      Width           =   255
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H80000001&
      FillColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   195
      MouseIcon       =   "frmUnitConverter.frx":0884
      ScaleHeight     =   1905
      ScaleWidth      =   5745
      TabIndex        =   8
      Top             =   3720
      Width           =   5805
   End
   Begin TreeXLibCtl.TreeX treeRules 
      Height          =   2895
      Left            =   195
      TabIndex        =   7
      Top             =   360
      Width           =   4245
      _cx             =   1368923456
      _cy             =   1368921074
      BorderStyle     =   4
      BackColor       =   -2147483647
      ForeColor       =   0
      PicturePosition =   0
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
   Begin VB.CommandButton cmdRules 
      Caption         =   "&Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4605
      TabIndex        =   6
      Top             =   5790
      Width           =   1335
   End
   Begin VB.CommandButton cmdRules 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4605
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdRules 
      Caption         =   "Rename"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4605
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdRules 
      Caption         =   "Modify ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4605
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdRules 
      Caption         =   "New ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4605
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Existing Rules:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   1
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Rule description (click on an underlined value to edit it):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   3480
      Width           =   4785
   End
End
Attribute VB_Name = "frmUnitConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal lngLength As Long)

 ' note: these are private, we only need pass reference to a single rule object to the Wizard
'       it definetly does not need access to the entire object store or collection
Private WithEvents m_OS As ObjectStore
Attribute m_OS.VB_VarHelpID = -1
Private WithEvents m_oLB As cCustomListBox
Attribute m_oLB.VB_VarHelpID = -1
Private m_oGroups As cINode
Private m_hTreeRootNode As Long

Private Sub Form_Load()
vbwProfiler.vbwProcIn 483
    Dim sFileName As String
    Dim sXPath As String
    Dim lptr As Long

vbwProfiler.vbwExecuteLine 10246
    sFileName = "D:\visualb\projects\UnitConverter\" & "test.xml" '"rules.xml"
vbwProfiler.vbwExecuteLine 10247
    sXPath = "object"

    ' load the user's previously saved wizard rules from file
vbwProfiler.vbwExecuteLine 10248
    Set m_oLB = New cCustomListBox
vbwProfiler.vbwExecuteLine 10249
    m_oLB.initDisplay picDisplay.hwnd, 0, 0
vbwProfiler.vbwExecuteLine 10250
    m_oLB.TextColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
vbwProfiler.vbwExecuteLine 10251
    m_oLB.BackColor = GetSysColor(COLOR_DESKTOP)
vbwProfiler.vbwExecuteLine 10252
    treeRules.ForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
vbwProfiler.vbwExecuteLine 10253
    treeRules.BackColor = GetSysColor(COLOR_DESKTOP)

vbwProfiler.vbwExecuteLine 10254
    Set m_OS = New ObjectStore
vbwProfiler.vbwExecuteLine 10255
    Set m_oGroups = m_OS.Deserialize(sFileName, 0, sXPath)
vbwProfiler.vbwExecuteLine 10256
    lptr = m_oGroups.Handle
vbwProfiler.vbwExecuteLine 10257
    Debug.Assert lptr = ObjPtr(m_oGroups)
    ' build our tree of existing rules from our loaded objects
vbwProfiler.vbwExecuteLine 10258
    m_hTreeRootNode = GraphVehicle(treeRules, 0, lptr)
vbwProfiler.vbwProcOut 483
vbwProfiler.vbwExecuteLine 10259
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
vbwProfiler.vbwProcIn 484
vbwProfiler.vbwExecuteLine 10260
    cmdRules_Click 4  ' done button
vbwProfiler.vbwProcOut 484
vbwProfiler.vbwExecuteLine 10261
End Sub

Private Sub Form_Unload(Cancel As Integer)
vbwProfiler.vbwProcIn 485
vbwProfiler.vbwExecuteLine 10262
    Set m_oGroups = Nothing
vbwProfiler.vbwExecuteLine 10263
    Set m_OS = Nothing
vbwProfiler.vbwExecuteLine 10264
    Set m_oLB = Nothing
vbwProfiler.vbwProcOut 485
vbwProfiler.vbwExecuteLine 10265
End Sub


Private Sub m_os_RequestObject(ByVal sClassname As String, ByVal DefPath As String, ByVal Def_ID As String, newObject As PersistenceManager.cIPersist)
vbwProfiler.vbwProcIn 486
    Dim oNode As cINode
    Dim lngPtr As Long

    ' based on the classname, create an instance of the object (this is a simple object factory)
vbwProfiler.vbwExecuteLine 10266
    Select Case sClassname
'vbwLine 10267:        Case "cRule"
        Case IIf(vbwProfiler.vbwExecuteLine(10267), VBWPROFILER_EMPTY, _
        "cRule")
vbwProfiler.vbwExecuteLine 10268
            Set oNode = New cRule

'vbwLine 10269:        Case "cGroupCollection"
        Case IIf(vbwProfiler.vbwExecuteLine(10269), VBWPROFILER_EMPTY, _
        "cGroupCollection")
vbwProfiler.vbwExecuteLine 10270
            Set oNode = New cGroupCollection

'vbwLine 10271:        Case "cRuleGroup"
        Case IIf(vbwProfiler.vbwExecuteLine(10271), VBWPROFILER_EMPTY, _
        "cRuleGroup")
vbwProfiler.vbwExecuteLine 10272
            Set oNode = New cRuleGroup
        Case Else
vbwProfiler.vbwExecuteLine 10273 'B
vbwProfiler.vbwExecuteLine 10274
            Debug.Print "frmUnitConverter:m_os_RequestObject() -- unknown object type"
    End Select
vbwProfiler.vbwExecuteLine 10275 'B

vbwProfiler.vbwExecuteLine 10276
    lngPtr = ObjPtr(oNode)
vbwProfiler.vbwExecuteLine 10277
    oNode.Handle = lngPtr
vbwProfiler.vbwExecuteLine 10278
    Debug.Assert oNode.Handle > 0
vbwProfiler.vbwExecuteLine 10279
    Debug.Print "frmUnitConverter:m_os_RequestObject() -- Successfully instanced object '" & sClassname & "'  Handle = " & lngPtr & ". "
vbwProfiler.vbwExecuteLine 10280
    Set newObject = oNode
vbwProfiler.vbwExecuteLine 10281
    Set oNode = Nothing
vbwProfiler.vbwProcOut 486
vbwProfiler.vbwExecuteLine 10282
End Sub

Private Sub cmdRules_Click(index As Integer)
vbwProfiler.vbwProcIn 487
vbwProfiler.vbwExecuteLine 10283
    Const CMD_NEW = 0
vbwProfiler.vbwExecuteLine 10284
    Const CMD_MODIFY = 1
vbwProfiler.vbwExecuteLine 10285
    Const CMD_RENAME = 2
vbwProfiler.vbwExecuteLine 10286
    Const CMD_DELETE = 3
vbwProfiler.vbwExecuteLine 10287
    Const CMD_DONE = 4

    Dim hObject As Long
    Dim hParent As Long
    Dim hTreeNode As Long

    Dim oNode As cINode
    Dim oParent As cINode
    Dim oRule As cRule
    Dim i As Long
    Dim sCategory As String

vbwProfiler.vbwExecuteLine 10288
    Select Case index
'vbwLine 10289:        Case CMD_NEW
        Case IIf(vbwProfiler.vbwExecuteLine(10289), VBWPROFILER_EMPTY, _
        CMD_NEW)

vbwProfiler.vbwExecuteLine 10290
            Set oNode = New cRule
vbwProfiler.vbwExecuteLine 10291
            hObject = ObjPtr(oNode)
vbwProfiler.vbwExecuteLine 10292
            oNode.Handle = hObject

vbwProfiler.vbwExecuteLine 10293
            Load frmWizard
vbwProfiler.vbwExecuteLine 10294
            frmWizard.Tag = hObject
vbwProfiler.vbwExecuteLine 10295
            frmWizard.Show vbModal, Me

            ' if the rule was given a "name" (description" then it was completed, else just bail
vbwProfiler.vbwExecuteLine 10296
            If oNode.Description <> "" Then
vbwProfiler.vbwExecuteLine 10297
                Set oRule = oNode
vbwProfiler.vbwExecuteLine 10298
                sCategory = oRule.Category
vbwProfiler.vbwExecuteLine 10299
                Set oRule = Nothing

                ' there is no possibility of finding a false match UNLESS the
                ' user alters the names in the XML directly by hand, otherwise
                ' the GUI wont allow names for rule nodes to be used that are
                ' reserved as "category" node names
vbwProfiler.vbwExecuteLine 10300
                hTreeNode = treeRules.FindItem(sCategory)
vbwProfiler.vbwExecuteLine 10301
                If hTreeNode = 0 Then
                    ' we have to create the cRuleGroup node, then add the new one and graph
                    ' it starting back at the root tree node
vbwProfiler.vbwExecuteLine 10302
                    Set oParent = New cRuleGroup
vbwProfiler.vbwExecuteLine 10303
                    hParent = ObjPtr(oParent)
vbwProfiler.vbwExecuteLine 10304
                    oParent.Handle = hParent
vbwProfiler.vbwExecuteLine 10305
                    oParent.Description = sCategory
vbwProfiler.vbwExecuteLine 10306
                    oParent.addChild oNode
vbwProfiler.vbwExecuteLine 10307
                    m_oGroups.addChild oParent
vbwProfiler.vbwExecuteLine 10308
                    GraphVehicle treeRules, m_hTreeRootNode, hParent
vbwProfiler.vbwExecuteLine 10309
                    Set oParent = Nothing
                Else
vbwProfiler.vbwExecuteLine 10310 'B
                    ' cRuleGroup already exists, just add the child
vbwProfiler.vbwExecuteLine 10311
                    hParent = treeRules.ItemData(hTreeNode)
vbwProfiler.vbwExecuteLine 10312
                    CopyMemory oParent, hParent, 4
vbwProfiler.vbwExecuteLine 10313
                    oParent.addChild oNode
vbwProfiler.vbwExecuteLine 10314
                    CopyMemory oParent, 0&, 4
vbwProfiler.vbwExecuteLine 10315
                    hObject = oNode.Handle
vbwProfiler.vbwExecuteLine 10316
                    GraphVehicle treeRules, hTreeNode, hObject
                End If
vbwProfiler.vbwExecuteLine 10317 'B
vbwProfiler.vbwExecuteLine 10318
                Set oNode = Nothing
            End If
vbwProfiler.vbwExecuteLine 10319 'B


'vbwLine 10320:        Case CMD_MODIFY
        Case IIf(vbwProfiler.vbwExecuteLine(10320), VBWPROFILER_EMPTY, _
        CMD_MODIFY)
            ' get the handle to the selected cInode
vbwProfiler.vbwExecuteLine 10321
            hTreeNode = treeRules.Selection
vbwProfiler.vbwExecuteLine 10322
            hObject = treeRules.ItemData(hTreeNode)
vbwProfiler.vbwExecuteLine 10323
            CopyMemory oNode, hObject, 4
vbwProfiler.vbwExecuteLine 10324
            If oNode.Classname = "cRule" Then

vbwProfiler.vbwExecuteLine 10325
                Load frmWizard
vbwProfiler.vbwExecuteLine 10326
                frmWizard.Tag = hObject
vbwProfiler.vbwExecuteLine 10327
                frmWizard.Show vbModal, Me

                ' if the user renamed it, change the node in the tree
vbwProfiler.vbwExecuteLine 10328
                treeRules.ItemText(hTreeNode) = oNode.Description
                ' re-select it (but de-select it first so that ItemClick event is triggered)
vbwProfiler.vbwExecuteLine 10329
                treeRules.SelectItem(hTreeNode) = False
vbwProfiler.vbwExecuteLine 10330
                treeRules.SelectItem(hTreeNode) = True
            Else
vbwProfiler.vbwExecuteLine 10331 'B
vbwProfiler.vbwExecuteLine 10332
                MsgBox "You must first select a rule node in the tree", vbInformation, "No rule node selected..."
            End If
vbwProfiler.vbwExecuteLine 10333 'B
vbwProfiler.vbwExecuteLine 10334
            CopyMemory oNode, 0&, 4

'vbwLine 10335:        Case CMD_RENAME
        Case IIf(vbwProfiler.vbwExecuteLine(10335), VBWPROFILER_EMPTY, _
        CMD_RENAME)
            ' todo: user can only rename cRule nodes and not rule groups or the group collection.
            '       remember that we search for the rule group names based on their baseID (they are
            '       named after the worksheet names)
vbwProfiler.vbwExecuteLine 10336
            hTreeNode = treeRules.Selection
vbwProfiler.vbwExecuteLine 10337
            hObject = treeRules.ItemData(hTreeNode)
vbwProfiler.vbwExecuteLine 10338
            CopyMemory oNode, hObject, 4
vbwProfiler.vbwExecuteLine 10339
            If oNode.Classname = "cRule" Then
                Dim sNewName As String
vbwProfiler.vbwExecuteLine 10340
                sNewName = InputBox("Enter a new name for this rule:", "Rename Rule...", treeRules.ItemText(hTreeNode))
vbwProfiler.vbwExecuteLine 10341
                If IsValidFilename(sNewName) Then
vbwProfiler.vbwExecuteLine 10342
                    If IsNotReservedName(sNewName) Then
vbwProfiler.vbwExecuteLine 10343
                        oNode.Description = sNewName
vbwProfiler.vbwExecuteLine 10344
                        treeRules.ItemText(hTreeNode) = sNewName
                    Else
vbwProfiler.vbwExecuteLine 10345 'B
vbwProfiler.vbwExecuteLine 10346
                        MsgBox "Please enter a different name:", vbInformation, "Reserved name"
                    End If
vbwProfiler.vbwExecuteLine 10347 'B
                Else
vbwProfiler.vbwExecuteLine 10348 'B
vbwProfiler.vbwExecuteLine 10349
                    MsgBox "Please enter a different name:", vbInformation, "Invalid name"
                End If
vbwProfiler.vbwExecuteLine 10350 'B
            End If
vbwProfiler.vbwExecuteLine 10351 'B
vbwProfiler.vbwExecuteLine 10352
            CopyMemory oNode, 0&, 4

'vbwLine 10353:        Case CMD_DELETE
        Case IIf(vbwProfiler.vbwExecuteLine(10353), VBWPROFILER_EMPTY, _
        CMD_DELETE)
            ' user can delete individual rules or an entire branch! (except for the root)
            ' we might want to try and add to the modTree since GVD can use it too
vbwProfiler.vbwExecuteLine 10354
            If treeRules.Selection > 0 Then
vbwProfiler.vbwExecuteLine 10355
                If treeRules.Selection <> m_hTreeRootNode Then
                    ' get a handle to the parent node object so we can tell it to delete the child
vbwProfiler.vbwExecuteLine 10356
                    hTreeNode = treeRules.ItemParent(treeRules.Selection)
vbwProfiler.vbwExecuteLine 10357
                    hObject = treeRules.ItemData(hTreeNode)
vbwProfiler.vbwExecuteLine 10358
                    CopyMemory oNode, hObject, 4
vbwProfiler.vbwExecuteLine 10359
                    i = oNode.getChildIndexByHandle(treeRules.ItemData(treeRules.Selection))
vbwProfiler.vbwExecuteLine 10360
                    If i >= 0 Then
vbwProfiler.vbwExecuteLine 10361
                        oNode.RemoveChild i
vbwProfiler.vbwExecuteLine 10362
                        treeRules.RemoveItemRec treeRules.Selection
                    End If
vbwProfiler.vbwExecuteLine 10363 'B
vbwProfiler.vbwExecuteLine 10364
                    CopyMemory oNode, 0&, 4
                Else
vbwProfiler.vbwExecuteLine 10365 'B
vbwProfiler.vbwExecuteLine 10366
                    MsgBox "The root node cannot be deleted.", vbInformation, "Access Denied."
                End If
vbwProfiler.vbwExecuteLine 10367 'B
            Else
vbwProfiler.vbwExecuteLine 10368 'B
vbwProfiler.vbwExecuteLine 10369
                MsgBox "You must first select a node in the tree", vbInformation, "No node selected."
            End If
vbwProfiler.vbwExecuteLine 10370 'B
'vbwLine 10371:        Case CMD_DONE
        Case IIf(vbwProfiler.vbwExecuteLine(10371), VBWPROFILER_EMPTY, _
        CMD_DONE)
            Dim sXML As String
            ' save it all
vbwProfiler.vbwExecuteLine 10372
             sXML = m_OS.Serialize(m_oGroups, False)
vbwProfiler.vbwExecuteLine 10373
             Debug.Print sXML

vbwProfiler.vbwExecuteLine 10374
            Unload Me
    End Select
vbwProfiler.vbwExecuteLine 10375 'B
vbwProfiler.vbwProcOut 487
vbwProfiler.vbwExecuteLine 10376
End Sub

Private Sub picDisplay_Paint()
vbwProfiler.vbwProcIn 488
vbwProfiler.vbwExecuteLine 10377
    If Not m_oLB Is Nothing Then
vbwProfiler.vbwExecuteLine 10378
        m_oLB.Paint
    End If
vbwProfiler.vbwExecuteLine 10379 'B
vbwProfiler.vbwProcOut 488
vbwProfiler.vbwExecuteLine 10380
End Sub

Private Sub treeRules_ItemSelect(ByVal hItem As Long)
vbwProfiler.vbwProcIn 489
    Dim oNode As cINode
    Dim oRule As cRule
vbwProfiler.vbwExecuteLine 10381
    Const LNG_LENGTH = 4
    Dim s As String
    Dim hNode As Long

vbwProfiler.vbwExecuteLine 10382
    Debug.Print "ITEM VALUE = " & treeRules.ItemText(hItem)
vbwProfiler.vbwExecuteLine 10383
    hNode = treeRules.ItemData(hItem)
    'Debug.Assert hNode = ObjPtr(m_oTESTNODE)
vbwProfiler.vbwExecuteLine 10384
    CopyMemory oNode, hNode, LNG_LENGTH
vbwProfiler.vbwExecuteLine 10385
    s = oNode.Classname

    ' if the node is a cRule, then render it
vbwProfiler.vbwExecuteLine 10386
    Select Case s
'vbwLine 10387:        Case "cRule" 'todo: use constants?
        Case IIf(vbwProfiler.vbwExecuteLine(10387), VBWPROFILER_EMPTY, _
        "cRule" )'todo: use constants?
vbwProfiler.vbwExecuteLine 10388
            Set oRule = oNode
vbwProfiler.vbwExecuteLine 10389
            renderRule oRule, m_oLB
'vbwLine 10390:        Case "cRuleGroup"
        Case IIf(vbwProfiler.vbwExecuteLine(10390), VBWPROFILER_EMPTY, _
        "cRuleGroup")
vbwProfiler.vbwExecuteLine 10391
            m_oLB.Clear
vbwProfiler.vbwExecuteLine 10392
            m_oLB.Paint
'vbwLine 10393:        Case "cGroupCollection"
        Case IIf(vbwProfiler.vbwExecuteLine(10393), VBWPROFILER_EMPTY, _
        "cGroupCollection")
vbwProfiler.vbwExecuteLine 10394
            m_oLB.Clear
vbwProfiler.vbwExecuteLine 10395
            m_oLB.Paint
        Case Else
vbwProfiler.vbwExecuteLine 10396 'B
vbwProfiler.vbwExecuteLine 10397
            Debug.Print "Else.. = " & s
    End Select
vbwProfiler.vbwExecuteLine 10398 'B
vbwProfiler.vbwExecuteLine 10399
    CopyMemory oNode, 0&, LNG_LENGTH
vbwProfiler.vbwProcOut 489
vbwProfiler.vbwExecuteLine 10400
End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 490
    Dim lngX As Long
    Dim lngY As Long
    Dim lRet As Long
    Dim hTreeNode As Long
    Dim hObject As Long

vbwProfiler.vbwExecuteLine 10401
    lngX = x \ Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 10402
    lngY = y \ Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 10403
    lRet = m_oLB.PointInHotSpot(lngX, lngY)

vbwProfiler.vbwExecuteLine 10404
    hTreeNode = treeRules.Selection
vbwProfiler.vbwExecuteLine 10405
    hObject = treeRules.ItemData(hTreeNode)

vbwProfiler.vbwExecuteLine 10406
    If lRet > 0 Then
vbwProfiler.vbwExecuteLine 10407
         Call displayItemClick(lRet, hObject, m_oLB)
    End If
vbwProfiler.vbwExecuteLine 10408 'B
vbwProfiler.vbwProcOut 490
vbwProfiler.vbwExecuteLine 10409
End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 491
    Dim lngX As Long
    Dim lngY As Long
vbwProfiler.vbwExecuteLine 10410
    lngX = x \ Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 10411
    lngY = y \ Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 10412
    If m_oLB.PointInHotSpot(lngX, lngY) > 0 Then
vbwProfiler.vbwExecuteLine 10413
        picDisplay.MousePointer = vbCustom
    Else
vbwProfiler.vbwExecuteLine 10414 'B
vbwProfiler.vbwExecuteLine 10415
        picDisplay.MousePointer = 1
    End If
vbwProfiler.vbwExecuteLine 10416 'B
vbwProfiler.vbwProcOut 491
vbwProfiler.vbwExecuteLine 10417
End Sub

Private Sub m_oLB_ItemAdded(ByVal lngItemCount As Long, ByVal lngMaxVisible As Long)
  ' configure the splitter bounds based on the num of rows and max visible
    ' an event should trigger this action
vbwProfiler.vbwProcIn 492
    Dim i As Long

vbwProfiler.vbwExecuteLine 10418
    i = lngItemCount
vbwProfiler.vbwExecuteLine 10419
    If lngMaxVisible < i Then
vbwProfiler.vbwExecuteLine 10420
        VScroll1.Max = i - lngMaxVisible
    Else
vbwProfiler.vbwExecuteLine 10421 'B
vbwProfiler.vbwExecuteLine 10422
        VScroll1.Max = 0
    End If
vbwProfiler.vbwExecuteLine 10423 'B
vbwProfiler.vbwExecuteLine 10424
    VScroll1.Min = 0
vbwProfiler.vbwExecuteLine 10425
   m_oLB.scrollPosition = VScroll1.value
vbwProfiler.vbwProcOut 492
vbwProfiler.vbwExecuteLine 10426
End Sub

Private Sub VScroll1_Change()
vbwProfiler.vbwProcIn 493
vbwProfiler.vbwExecuteLine 10427
    m_oLB.scrollPosition = VScroll1.value
vbwProfiler.vbwExecuteLine 10428
    m_oLB.RenderText
vbwProfiler.vbwProcOut 493
vbwProfiler.vbwExecuteLine 10429
End Sub

Private Sub VScroll1_Scroll()
vbwProfiler.vbwProcIn 494
vbwProfiler.vbwExecuteLine 10430
    m_oLB.scrollPosition = VScroll1.value
vbwProfiler.vbwExecuteLine 10431
    m_oLB.RenderText
vbwProfiler.vbwProcOut 494
vbwProfiler.vbwExecuteLine 10432
End Sub

