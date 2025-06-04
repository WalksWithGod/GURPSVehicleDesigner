VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCompList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        x As Long
        y As Long
End Type

'Private Type m_Cell
'    r As RECT
'    text As String
'    iMore As Long   ' array index for extended drop down list
'    iSubCount As Long ' number of subitems
'End Type

'-------------------
' CONVERT TO HEAP
Private Type uNode
    bString(0 To 30) As Byte
    pBranch As Long 'pointer to start of sublist heap
    pNext As Long   ' pointer to next node
    bSelectedState As Byte
End Type

Const NODE_SELECTED = 1
Const NODE_DESELECTED = 0


Private m_hHeap As Long
Private m_hStart As Long ' pointer to start of primary list
Private m_hHeapSub As Long
Private m_oCurrentNode As uNode

Const PAGE_SIZE = 4096   'only on Alphas, the page size is 8196
Const HEAP_SIZE = 8192
Const MAX_HEAP_SIZE = 16384

Const MAXIMUM_POLL_RATE = 25

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
'-------------------
'Private m_Items() As m_Cell
'Private m_SubItems() As String ' two dimensional array.  First subscript is index to m_Items, 2nd subscript contains actual data
Private m_ArrowPoints(0 To 2) As POINTAPI



Private Const PT_MOVETO = &H6
Private Const PT_LINETO = &H2
Private Const PT_CLOSEFIGURE = &H1
Private Const PT_BEZIERTO = &H4

Private m_lngMaxWidth As Long
Private m_lngMaxHeight As Long
Private m_lngCellHeight As Long
Private m_lngCellCount As Long
Private m_lngCellWidth As Long
Private m_sFileName As String

Private m_lngTwipsX As Long
Private m_lngTwipsY As Long

Private m_crColorWindowText As Long
Private m_crColorHighlightText As Long
Private m_hDefaultBrush As Long
Private m_hFocusRectBrush As Long  ' brush used for solid rect around selected item
Private m_hOriginalBrush As Long  ' brush for non selected item rectangle
Private m_hHighlightArrowBrush As Long ' brush for the arrow so that it matches the color of the highlight text
Private m_hArrowBrush As Long
Private m_hPen As Long
Private m_hOriginalPen As Long

Private WithEvents m_oPic As PictureBox
Attribute m_oPic.VB_VarHelpID = -1
Private WithEvents m_oSubPic As PictureBox
Attribute m_oSubPic.VB_VarHelpID = -1

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long

#If DEBUG_MODE Then
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function GetFocus Lib "user32" () As Long
#End If
Public Event Click(ByRef s As String)

Const LEFT_CELL_PADDING = 2
Const TOP_CELL_PADDING = 3
Const MAX_SUBITEMS = 32
Const UNIQUE_KEY = "DROPDOWN"
Const UNIQUE_KEY_2 = "SUBDROPDOWN"

Const ARROWHEIGHT = 6
Const ARROWWIDTH = 5
Const ARROWMID = 3
    

' returns TRUE if successful
Public Function SetFileName(ByRef s As String) As Long
    Dim lRet As Long
    If s <> "" Then ' todo: should maybe call routine to check if IsFile or something
        m_sFileName = s
        SetFileName = True
        
        ' load lists
        If Not Load_Components Then
            MsgBox "Error loading vehicle components list."
            Call Class_Terminate
        End If
    Else
        SetFileName = False
    End If
End Function

' This class handles the reading and listing of components
Private Sub Class_Initialize()
    Dim lRet As Long
    Dim lngTempHeapSize As Long

    On Error GoTo errorhandler

    Set m_oPic = frmDesigner.Controls.Add("VB.PictureBox", UNIQUE_KEY)
    Set m_oSubPic = frmDesigner.Controls.Add("VB.PictureBox", UNIQUE_KEY_2)
    With m_oPic
        .Font = frmDesigner.treeVehicle.Font
        .ZOrder
        .ScaleMode = vbPixels
        .BorderStyle = vbFixedSingle
        .Appearance = vbFlat
        .Visible = False
    End With
    
    With m_oSubPic
        .Font = frmDesigner.treeVehicle.Font
        .ZOrder
        .ScaleMode = vbPixels
        .BorderStyle = vbFixedSingle
        .Appearance = vbFlat
        .Visible = False
    End With
    
    ' create two brushes which we use to draw our selected rect
    m_crColorWindowText = GetSysColor(COLOR_WINDOWTEXT)
    m_crColorHighlightText = GetSysColor(COLOR_HIGHLIGHTTEXT)
    
    m_hDefaultBrush = CreateSolidBrush(GetSysColor(COLOR_WINDOW))
    m_hOriginalBrush = SelectObject(m_oPic.hdc, m_hDefaultBrush)  'todo: im still not loading a brush for the sub pic, note that the arrow we draw using the polygon function, uses the currently loaded brush for its color and NOT textcolor
    m_hFocusRectBrush = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
    m_hHighlightArrowBrush = CreateSolidBrush(m_crColorHighlightText)
    m_hArrowBrush = CreateSolidBrush(m_crColorWindowText)
    
    ' try these pens to see if they fix problems we are having drawing our fucking arrows
    m_hPen = CreatePen(PS_SOLID, 1, m_crColorHighlightText)
    m_hOriginalPen = SelectObject(m_oPic.hdc, m_hPen)
    ' get the font height for determining the cell height
    m_lngMaxHeight = m_oPic.TextHeight("W")
    m_lngCellHeight = m_lngMaxHeight + 4
    
    ' create our primary heap
    m_hHeap = HeapCreate(HEAP_NO_SERIALIZE Or HEAP_GENERATE_EXCEPTIONS, HEAP_SIZE, 0)
    Debug.Print "Heap Created @ " & m_hHeap
    m_hStart = HeapAlloc(ByVal m_hHeap, HEAP_ZERO_MEMORY, HEAP_SIZE)
    Debug.Print "Heap Allocated @ " & m_hStart
    lngTempHeapSize = HeapSize(ByVal m_hHeap, 0, ByVal m_hStart)
    Debug.Print "Heap Size = " & lngTempHeapSize
            
    m_hHeapSub = HeapCreate(HEAP_NO_SERIALIZE Or HEAP_GENERATE_EXCEPTIONS, MAX_HEAP_SIZE, 0)
    If m_hStart <= 0 Then GoTo errorhandler
    Exit Sub
    
errorhandler:
    MsgBox "Could not ceate custom drop down menu"
    Call Class_Terminate
    Exit Sub
End Sub


Private Sub Class_Terminate()
    ' erase arrays and destroy object
    Dim l As Long
    Dim b As Boolean
    
    'todo: the way im handling original pens and brushes really only works for the main pic and not the sub pic.
    ' ideally id like to make it so i can have an arbitrary number of sublists.  This would also mean each
    ' subllist would use the same code and cleanup would be automatic
    l = SelectObject(m_oPic.hdc, m_hOriginalBrush)
    l = DeleteObject(m_hDefaultBrush)
    l = DeleteObject(m_hFocusRectBrush)
    
    Set m_oPic = Nothing
    Set m_oSubPic = Nothing
    
    b = DestroySubHeaps
    Debug.Print "DestroySubHeaps Returns " & b
    
    HeapFree ByVal m_hHeap, 0, m_hStart
    HeapDestroy m_hHeap
    HeapDestroy m_hHeapSub
    
    #If DEBUG_MODE Then
        frmDesigner.txtInfo.Text = "clsCompList:Terminate() -- Successfully terminated custom dropdown menu control"
    #End If
End Sub

Private Function DestroySubHeaps() As Boolean
    On Error GoTo err
      
    Dim i As Long
    Dim oNode As uNode
    Dim hAddress As Long
    Dim lngNodeSize As Long
    Dim l As Long
    
    ' goes through each parent node in the primary heap, finds the offset (pBranch) in the secondary heap
    ' which is the pointer to the specific subheap and frees it.
    lngNodeSize = LenB(oNode)
    hAddress = m_hStart
    
    For i = 0 To m_lngCellCount
        CopyMemory oNode, ByVal hAddress, lngNodeSize
        l = HeapFree(ByVal m_hHeapSub, 0, oNode.pBranch)
        Debug.Print "clsCompList:DestroySubHeaps -- Itteration = " & i & " HeapFree Returning " & l
        hAddress = oNode.pNext
    Next
    DestroySubHeaps = True
    Exit Function
err:
    DestroySubHeaps = False
End Function

Public Function ShowDropDown() As Long
    Dim i As Long

    On Error GoTo errorhandler
        
    m_oPic.Visible = True
    m_oPic.Top = frmDesigner.cboComponents.Top + frmDesigner.cboComponents.Height
    m_oPic.Left = frmDesigner.cboComponents.Left
    m_oPic.Height = VB.Screen.TwipsPerPixelX * (m_lngCellHeight * (m_lngCellCount + 1))
    m_oPic.Width = 2500  ' m_lngCellWidth
    m_oPic.ZOrder
    m_oSubPic.ZOrder
    
    ' temp: paint dropdown
    #If DEBUG_MODE Then
        InfoPrint 1, timeGetTime & " clsCompList.ShowDropDown -- Stage 0: Menu Handle = " & m_oPic.hwnd
    #End If
    
    #If DEBUG_MODE Then
        InfoPrint 1, timeGetTime & " clsCompList.ShowDropDown -- Stage 1: Focus Window Handle = " & GetFocus
    #End If
    
    Call DrawItems(m_oPic, ByVal m_hStart)
    SetForegroundWindow m_oPic.hwnd
    
    #If DEBUG_MODE Then
        InfoPrint 1, timeGetTime & " clsCompList.ShowDropDown -- Stage 2: Focus Window Handle = " & GetFocus
    #End If
    
    SetFocus m_oPic.hwnd
    
    #If DEBUG_MODE Then
        InfoPrint 1, timeGetTime & " clsCompList.ShowDropDown -- Stage 3: Focus Window Handle = " & GetFocus
    #End If
    ShowDropDown = True
    Exit Function
errorhandler:
    ShowDropDown = False
End Function

Private Sub DrawItems(ByRef oPic As PictureBox, ByVal hAddress As Long)
    Dim i As Long
    Dim l As Long
    Dim s As String
    Dim yPos As Long
    Dim r As Rect
    Dim oNode As uNode
    Dim lngNodeSize As Long
    Dim hMem As Long
    Dim j As Long
    Dim lRet As Long
    Dim hOriginalPen As Long
    
    oPic.DrawMode = 13  '  <-- damnit, I had this draw mode set to 1=BLACKNESS and couldnt figure out why the arrows wouldnt draw same color as pen/brush! argh
    
    hMem = hAddress
    lngNodeSize = LenB(oNode)
    m_lngCellWidth = oPic.Width
    
    Do While hMem > 0
        CopyMemory oNode, ByVal hMem, lngNodeSize
        hMem = oNode.pNext
        
        s = NodeBytesToString(oNode.bString)
        l = Len(s)
        yPos = i * m_lngCellHeight
        
        With r
            .Top = yPos - 1
            .Bottom = yPos + m_lngCellHeight + 1
            .Left = 0
            .Right = m_lngCellWidth
        End With
          
        If oNode.bSelectedState = NODE_SELECTED Then
            Call FillRect(ByVal oPic.hdc, r, m_hFocusRectBrush)
            SetTextColor oPic.hdc, m_crColorHighlightText
            SelectObject oPic.hdc, m_hHighlightArrowBrush
            SelectObject oPic.hdc, m_hPen
        Else
            Call FillRect(ByVal oPic.hdc, r, m_hDefaultBrush)
            SetTextColor oPic.hdc, m_crColorWindowText
            SelectObject oPic.hdc, m_hArrowBrush
            SelectObject oPic.hdc, m_hOriginalPen
        End If
          
        ' only draw arrow if there is a branch off this node
        If oNode.pBranch > 0 Then
            Call NextArrow(i)
            Call Polygon(ByVal oPic.hdc, m_ArrowPoints(0), 3)
        End If
        Call TextOut(ByVal oPic.hdc, LEFT_CELL_PADDING, yPos, s, l)
    i = i + 1
    Loop
End Sub

Private Sub NextArrow(ByVal iIndex As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngOffset As Long
    
    lngLeft = 150
    If iIndex = 0 Then
        lngTop = 3
    Else
        lngOffset = (m_lngCellHeight - ARROWHEIGHT) \ 2
        lngTop = (iIndex * m_lngCellHeight) + lngOffset
    End If
    m_ArrowPoints(0).x = lngLeft
    m_ArrowPoints(0).y = lngTop
    m_ArrowPoints(1).x = lngLeft + ARROWWIDTH
    m_ArrowPoints(1).y = lngTop + ARROWMID
    m_ArrowPoints(2).x = lngLeft
    m_ArrowPoints(2).y = lngTop + ARROWHEIGHT
End Sub

Private Function Load_Components() As Long
    Dim sSections() As String 'main sections
    Dim sKeys() As String     'sub categories for each section
    Dim sKeyNames() As String
    Dim lKeyCount As Long
    Dim i As Integer
    Dim j As Long
    Dim oINI As cINI
    Dim oParent As uNode
    Dim oChild As uNode
    Dim hAddress As Long
    Dim lngNodeSize As Long
    Dim lngBranchSize As Long
    Dim hBranchAddress As Long
    Dim oTemp As uNode
    Dim lngTempHeapSize As Long
    Dim b() As Byte
    
    On Error GoTo errorhandler
       
    'make sure we are back in the program's install path
    ChDir App.Path
    Set oINI = New cINI
    oINI.FileName = m_sFileName
    sSections = oINI.RetreiveSectionNames
    
    ' check for zero array
    m_lngCellCount = UBound(sSections)
    If (m_lngCellCount = 0) And (sSections(0) = "") Then GoTo errorhandler
    
    ' all is good, lets retrieve them and store them
    lngNodeSize = LenB(oParent)
    hAddress = m_hStart
    
    'todo:!!!! Check that our heap size is big enuf to contain LenB(uNode) * m_lngCellCount
    For i = 0 To m_lngCellCount
        b = StrConv(sSections(i), vbFromUnicode)
        Erase oParent.bString
        CopyMemory oParent.bString(0), b(0), Minimum(30, UBound(b) + 1)
                   
        ' also retreive the subcategories for each main category
        sKeyNames = oINI.ReadSection(sSections(i))
        lKeyCount = UBound(sKeyNames)
        If (lKeyCount = 0) And (sKeyNames(0) = "") Then GoTo errorhandler
                 
        ' create a heap that will hold this branch
        lngBranchSize = (lKeyCount + 1) * lngNodeSize
        hBranchAddress = HeapAlloc(m_hHeapSub, HEAP_ZERO_MEMORY, lngBranchSize)
        oParent.pBranch = hBranchAddress
        If i = m_lngCellCount Then
            oParent.pNext = 0
        Else
            oParent.pNext = hAddress + lngNodeSize
        End If
        'todo: check hBranch value for error in heap creation
    
        ' copy the uParent to our primary list before proceeding to branch copy
        CopyMemory ByVal hAddress, oParent, lngNodeSize
        CopyMemory oTemp, ByVal hAddress, lngNodeSize
        hAddress = hAddress + lngNodeSize
        
        ' copy branch nodes subHeap
        For j = 0 To lKeyCount
            b = StrConv(oINI.ReadString(sSections(i), sKeyNames(j)), vbFromUnicode)
            Erase oChild.bString
            CopyMemory oChild.bString(0), b(0), Minimum(30, UBound(b) + 1)
            oChild.pBranch = 0
            If j = lKeyCount Then
                oChild.pNext = 0
            Else
                oChild.pNext = hBranchAddress + lngNodeSize
            End If
            CopyMemory ByVal hBranchAddress, oChild, lngNodeSize
            CopyMemory oTemp, ByVal hBranchAddress, lngNodeSize
            hBranchAddress = hBranchAddress + lngNodeSize
        Next
    Next
    Load_Components = True
    Exit Function
    
errorhandler:
    Load_Components = False
End Function

Private Sub m_oPic_LostFocus()
    m_oPic.Visible = False
    m_oSubPic.Visible = False
End Sub

Private Sub m_oPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngNodeSize As Long
    Dim i As Long
    Dim hAddress As Long
    Dim hBranchStart As Long
    Dim lngMouseOver As Long
    Static lngLastMouseOver As Long
    Dim oTempNode As uNode
    Static lngLastTime As Long
    
    'todo: this returns a negative number if computer is running more than 23 days.
    '      queryperformancecounter doesnt work on all cpu's or old versions of windows95 either i think.
    '      need a better timer.  i could code around it by checking for neg values and in that case, subtract last from current
    'If timeGetTime - lngLastTime < MAXIMUM_POLL_RATE Then Exit Sub
    lngLastTime = timeGetTime
    
    ' if the time between moves is too fast, exit sub. We dont want this to be
    ' too sensitive
    #If DEBUG_MODE Then
        InfoPrint 1, timeGetTime & " clsCompList:MouseMove -- mouse movement in main menu OK"
    #End If
    
    m_lngTwipsX = VB.Screen.TwipsPerPixelX
    lngNodeSize = LenB(m_oCurrentNode)
    ' determine which cell we are currently over
    lngMouseOver = GetMouseOverCell(x, y)
    
    If lngLastMouseOver <> lngMouseOver Then
        With m_oSubPic
        ' reposition the sub list
            .Visible = False
            .Top = m_oPic.Top + m_lngTwipsX * (m_lngCellHeight * lngMouseOver)
            '.Height = VB.Screen.TwipsPerPixelY * ((m_Items(m_lngMouseOver).iSubCount + 1) * m_lngCellHeight)
            .Left = m_oPic.Left + m_oPic.Width
            .Width = m_oPic.Width
            .Visible = True
        End With
       
        ' update the selected state of the previous node in the heap
        hAddress = m_hStart + (lngLastMouseOver * lngNodeSize)
        CopyMemory m_oCurrentNode, ByVal hAddress, lngNodeSize
        m_oCurrentNode.bSelectedState = NODE_DESELECTED
        CopyMemory ByVal hAddress, m_oCurrentNode, lngNodeSize
    End If
    lngLastMouseOver = lngMouseOver
    
    ' get the address for the subcategory to draw
    hAddress = m_hStart + (lngMouseOver * lngNodeSize)
    CopyMemory m_oCurrentNode, ByVal hAddress, lngNodeSize
    ' update the selected state of the node in the heap
    m_oCurrentNode.bSelectedState = NODE_SELECTED
    m_oPic.Tag = NodeBytesToString(m_oCurrentNode.bString)
    CopyMemory ByVal hAddress, m_oCurrentNode, lngNodeSize

    hBranchStart = m_oCurrentNode.pBranch
    hAddress = hBranchStart
    
    ' determine how many nodes are in this branch so we can determine picbox height
    Do While hAddress > 0
        CopyMemory oTempNode, ByVal hAddress, lngNodeSize
        hAddress = oTempNode.pNext
        i = i + 1
    Loop
    With m_oSubPic
        .Height = m_lngTwipsX * ((i) * m_lngCellHeight)
        .Tag = hBranchStart
    End With
    ' set our mem pointer back to start of the selected branch and draw it
    Call DrawItems(m_oSubPic, hBranchStart)
    ' draw both dropdowns
    Call DrawItems(m_oPic, m_hStart)
End Sub

Private Function GetMouseOverCell(x, y) As Long
    Dim l As Long
    Dim i As Long
    
    l = CLng(y)
    l = l \ m_lngCellHeight
    GetMouseOverCell = l
#If DEBUG_MODE Then
    InfoPrint 1, timeGetTime & "clsCompList:GetMouseOverCell -- detected cell = " & l
#End If
End Function

Private Sub m_oPic_Paint()
    Call DrawItems(m_oPic, m_hStart)
End Sub

Private Function NodeBytesToString(b() As Byte) As String
    Dim s As String
    Dim l As Long
    
    s = Trim(StrConv(b, vbUnicode))
    l = InStr(1, s, Chr(0))
    s = Left(s, l - 1)
    NodeBytesToString = s
End Function

Private Sub m_oSubPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim l As Long
    Dim oNode As uNode
    Dim lngNodeSize As Long
    Dim s As String
    Dim hAddress As Long
    
    lngNodeSize = LenB(oNode)
    l = GetMouseOverCell(x, y)
    hAddress = CLng(m_oSubPic.Tag) + (l * lngNodeSize)
    
    CopyMemory oNode, ByVal hAddress, lngNodeSize
    s = NodeBytesToString(oNode.bString)
    
    m_oPic.Visible = False
    m_oSubPic.Visible = False
    RaiseEvent Click(m_oPic.Tag & "\" & s)
End Sub

Private Sub m_oSubPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngMouseOver As Long
    Static lngLastMouseOver As Long
    Dim hAddress As Long
    Dim lngNodeSize As Long
    Dim oTempNode As uNode
    Static lngLastTime As Long
    
    'todo: this will not work on machines that have been running long enuf for timegettime to wrap around
    ' we will get negative numbers and so,
    'If timeGetTime - lngLastTime < MAXIMUM_POLL_RATE Then Exit Sub
    lngLastTime = timeGetTime
    
    lngNodeSize = LenB(m_oCurrentNode)
    ' determine which cell we are currently over
    lngMouseOver = GetMouseOverCell(x, y)
    
    If m_oCurrentNode.pBranch <= 0 Then
        Debug.Print "clsComplist:m_oSubPic_MouseMove() -- Branch node address not set.  Exiting Sub"
        Exit Sub
    End If
    
    #If DEBUG_MODE Then
        InfoPrint 1, timeGetTime & " clsCompList:SubMenuMouseMove -- mouse movement detected OK"
    #End If

    ' update the selected state of the previous node in the heap
    hAddress = m_oCurrentNode.pBranch + (lngLastMouseOver * lngNodeSize)
    If hAddress >= 0 Then
        CopyMemory oTempNode, ByVal hAddress, lngNodeSize
        oTempNode.bSelectedState = NODE_DESELECTED
        CopyMemory ByVal hAddress, oTempNode, lngNodeSize
    End If

    lngLastMouseOver = lngMouseOver
        
    ' get the address for the subcategory to draw
    hAddress = m_oCurrentNode.pBranch + (lngMouseOver * lngNodeSize)
    CopyMemory oTempNode, ByVal hAddress, lngNodeSize
    ' update the selected state of the node in the heap
    oTempNode.bSelectedState = NODE_SELECTED
    CopyMemory ByVal hAddress, oTempNode, lngNodeSize
         
    ' set our mem pointer back to start of the selected branch and draw it
    Call DrawItems(m_oSubPic, m_oCurrentNode.pBranch)
    
    ' deselect the item for next time
    CopyMemory oTempNode, ByVal hAddress, lngNodeSize
    oTempNode.bSelectedState = NODE_DESELECTED
    CopyMemory ByVal hAddress, oTempNode, lngNodeSize
End Sub
