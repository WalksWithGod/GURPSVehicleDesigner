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
vbwProfiler.vbwProcIn 209
    Dim lRet As Long
vbwProfiler.vbwExecuteLine 4128
    If s <> "" Then ' todo: should maybe call routine to check if IsFile or something
vbwProfiler.vbwExecuteLine 4129
        m_sFileName = s
vbwProfiler.vbwExecuteLine 4130
        SetFileName = True

        ' load lists
vbwProfiler.vbwExecuteLine 4131
        If Not Load_Components Then
vbwProfiler.vbwExecuteLine 4132
            MsgBox "Error loading vehicle components list."
vbwProfiler.vbwExecuteLine 4133
            Call Class_Terminate
        End If
vbwProfiler.vbwExecuteLine 4134 'B
    Else
vbwProfiler.vbwExecuteLine 4135 'B
vbwProfiler.vbwExecuteLine 4136
        SetFileName = False
    End If
vbwProfiler.vbwExecuteLine 4137 'B
vbwProfiler.vbwProcOut 209
vbwProfiler.vbwExecuteLine 4138
End Function

' This class handles the reading and listing of components
Private Sub Class_Initialize()
vbwProfiler.vbwProcIn 210
    Dim lRet As Long
    Dim lngTempHeapSize As Long

vbwProfiler.vbwExecuteLine 4139
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 4140
    Set m_oPic = frmDesigner.Controls.Add("VB.PictureBox", UNIQUE_KEY)
vbwProfiler.vbwExecuteLine 4141
    Set m_oSubPic = frmDesigner.Controls.Add("VB.PictureBox", UNIQUE_KEY_2)
vbwProfiler.vbwExecuteLine 4142
    With m_oPic
vbwProfiler.vbwExecuteLine 4143
        .Font = frmDesigner.treeVehicle.Font
vbwProfiler.vbwExecuteLine 4144
        .ZOrder
vbwProfiler.vbwExecuteLine 4145
        .ScaleMode = vbPixels
vbwProfiler.vbwExecuteLine 4146
        .BorderStyle = vbFixedSingle
vbwProfiler.vbwExecuteLine 4147
        .Appearance = vbFlat
vbwProfiler.vbwExecuteLine 4148
        .Visible = False
vbwProfiler.vbwExecuteLine 4149
    End With

vbwProfiler.vbwExecuteLine 4150
    With m_oSubPic
vbwProfiler.vbwExecuteLine 4151
        .Font = frmDesigner.treeVehicle.Font
vbwProfiler.vbwExecuteLine 4152
        .ZOrder
vbwProfiler.vbwExecuteLine 4153
        .ScaleMode = vbPixels
vbwProfiler.vbwExecuteLine 4154
        .BorderStyle = vbFixedSingle
vbwProfiler.vbwExecuteLine 4155
        .Appearance = vbFlat
vbwProfiler.vbwExecuteLine 4156
        .Visible = False
vbwProfiler.vbwExecuteLine 4157
    End With

    ' create two brushes which we use to draw our selected rect
vbwProfiler.vbwExecuteLine 4158
    m_crColorWindowText = GetSysColor(COLOR_WINDOWTEXT)
vbwProfiler.vbwExecuteLine 4159
    m_crColorHighlightText = GetSysColor(COLOR_HIGHLIGHTTEXT)

vbwProfiler.vbwExecuteLine 4160
    m_hDefaultBrush = CreateSolidBrush(GetSysColor(COLOR_WINDOW))
vbwProfiler.vbwExecuteLine 4161
    m_hOriginalBrush = SelectObject(m_oPic.hdc, m_hDefaultBrush)  'todo: im still not loading a brush for the sub pic, note that the arrow we draw using the polygon function, uses the currently loaded brush for its color and NOT textcolor
vbwProfiler.vbwExecuteLine 4162
    m_hFocusRectBrush = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
vbwProfiler.vbwExecuteLine 4163
    m_hHighlightArrowBrush = CreateSolidBrush(m_crColorHighlightText)
vbwProfiler.vbwExecuteLine 4164
    m_hArrowBrush = CreateSolidBrush(m_crColorWindowText)

    ' try these pens to see if they fix problems we are having drawing our fucking arrows
vbwProfiler.vbwExecuteLine 4165
    m_hPen = CreatePen(PS_SOLID, 1, m_crColorHighlightText)
vbwProfiler.vbwExecuteLine 4166
    m_hOriginalPen = SelectObject(m_oPic.hdc, m_hPen)
    ' get the font height for determining the cell height
vbwProfiler.vbwExecuteLine 4167
    m_lngMaxHeight = m_oPic.TextHeight("W")
vbwProfiler.vbwExecuteLine 4168
    m_lngCellHeight = m_lngMaxHeight + 4

    ' create our primary heap
vbwProfiler.vbwExecuteLine 4169
    m_hHeap = HeapCreate(HEAP_NO_SERIALIZE Or HEAP_GENERATE_EXCEPTIONS, HEAP_SIZE, 0)
vbwProfiler.vbwExecuteLine 4170
    Debug.Print "Heap Created @ " & m_hHeap
vbwProfiler.vbwExecuteLine 4171
    m_hStart = HeapAlloc(ByVal m_hHeap, HEAP_ZERO_MEMORY, HEAP_SIZE)
vbwProfiler.vbwExecuteLine 4172
    Debug.Print "Heap Allocated @ " & m_hStart
vbwProfiler.vbwExecuteLine 4173
    lngTempHeapSize = HeapSize(ByVal m_hHeap, 0, ByVal m_hStart)
vbwProfiler.vbwExecuteLine 4174
    Debug.Print "Heap Size = " & lngTempHeapSize

vbwProfiler.vbwExecuteLine 4175
    m_hHeapSub = HeapCreate(HEAP_NO_SERIALIZE Or HEAP_GENERATE_EXCEPTIONS, MAX_HEAP_SIZE, 0)
vbwProfiler.vbwExecuteLine 4176
    If m_hStart <= 0 Then
vbwProfiler.vbwExecuteLine 4177
         GoTo errorhandler
    End If
vbwProfiler.vbwExecuteLine 4178 'B
vbwProfiler.vbwProcOut 210
vbwProfiler.vbwExecuteLine 4179
    Exit Sub

errorhandler:
vbwProfiler.vbwExecuteLine 4180
    MsgBox "Could not ceate custom drop down menu"
vbwProfiler.vbwExecuteLine 4181
    Call Class_Terminate
vbwProfiler.vbwProcOut 210
vbwProfiler.vbwExecuteLine 4182
    Exit Sub
vbwProfiler.vbwProcOut 210
vbwProfiler.vbwExecuteLine 4183
End Sub


Private Sub Class_Terminate()
    ' erase arrays and destroy object
vbwProfiler.vbwProcIn 211
    Dim l As Long
    Dim b As Boolean

    'todo: the way im handling original pens and brushes really only works for the main pic and not the sub pic.
    ' ideally id like to make it so i can have an arbitrary number of sublists.  This would also mean each
    ' subllist would use the same code and cleanup would be automatic
vbwProfiler.vbwExecuteLine 4184
    l = SelectObject(m_oPic.hdc, m_hOriginalBrush)
vbwProfiler.vbwExecuteLine 4185
    l = DeleteObject(m_hDefaultBrush)
vbwProfiler.vbwExecuteLine 4186
    l = DeleteObject(m_hFocusRectBrush)

vbwProfiler.vbwExecuteLine 4187
    Set m_oPic = Nothing
vbwProfiler.vbwExecuteLine 4188
    Set m_oSubPic = Nothing

vbwProfiler.vbwExecuteLine 4189
    b = DestroySubHeaps
vbwProfiler.vbwExecuteLine 4190
    Debug.Print "DestroySubHeaps Returns " & b

vbwProfiler.vbwExecuteLine 4191
    HeapFree ByVal m_hHeap, 0, m_hStart
vbwProfiler.vbwExecuteLine 4192
    HeapDestroy m_hHeap
vbwProfiler.vbwExecuteLine 4193
    HeapDestroy m_hHeapSub

    #If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 4194
        frmDesigner.txtInfo.Text = "clsCompList:Terminate() -- Successfully terminated custom dropdown menu control"
    #End If
vbwProfiler.vbwProcOut 211
vbwProfiler.vbwExecuteLine 4195
End Sub

Private Function DestroySubHeaps() As Boolean
vbwProfiler.vbwProcIn 212
vbwProfiler.vbwExecuteLine 4196
    On Error GoTo err

    Dim i As Long
    Dim oNode As uNode
    Dim hAddress As Long
    Dim lngNodeSize As Long
    Dim l As Long

    ' goes through each parent node in the primary heap, finds the offset (pBranch) in the secondary heap
    ' which is the pointer to the specific subheap and frees it.
vbwProfiler.vbwExecuteLine 4197
    lngNodeSize = LenB(oNode)
vbwProfiler.vbwExecuteLine 4198
    hAddress = m_hStart

vbwProfiler.vbwExecuteLine 4199
    For i = 0 To m_lngCellCount
vbwProfiler.vbwExecuteLine 4200
        CopyMemory oNode, ByVal hAddress, lngNodeSize
vbwProfiler.vbwExecuteLine 4201
        l = HeapFree(ByVal m_hHeapSub, 0, oNode.pBranch)
vbwProfiler.vbwExecuteLine 4202
        Debug.Print "clsCompList:DestroySubHeaps -- Itteration = " & i & " HeapFree Returning " & l
vbwProfiler.vbwExecuteLine 4203
        hAddress = oNode.pNext
vbwProfiler.vbwExecuteLine 4204
    Next
vbwProfiler.vbwExecuteLine 4205
    DestroySubHeaps = True
vbwProfiler.vbwProcOut 212
vbwProfiler.vbwExecuteLine 4206
    Exit Function
err:
vbwProfiler.vbwExecuteLine 4207
    DestroySubHeaps = False
vbwProfiler.vbwProcOut 212
vbwProfiler.vbwExecuteLine 4208
End Function

Public Function ShowDropDown() As Long
vbwProfiler.vbwProcIn 213
    Dim i As Long

vbwProfiler.vbwExecuteLine 4209
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 4210
    m_oPic.Visible = True
vbwProfiler.vbwExecuteLine 4211
    m_oPic.Top = frmDesigner.cboComponents.Top + frmDesigner.cboComponents.Height
vbwProfiler.vbwExecuteLine 4212
    m_oPic.Left = frmDesigner.cboComponents.Left
vbwProfiler.vbwExecuteLine 4213
    m_oPic.Height = VB.Screen.TwipsPerPixelX * (m_lngCellHeight * (m_lngCellCount + 1))
vbwProfiler.vbwExecuteLine 4214
    m_oPic.Width = 2500  ' m_lngCellWidth
vbwProfiler.vbwExecuteLine 4215
    m_oPic.ZOrder
vbwProfiler.vbwExecuteLine 4216
    m_oSubPic.ZOrder

    ' temp: paint dropdown
    #If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 4217
        InfoPrint 1, timeGetTime & " clsCompList.ShowDropDown -- Stage 0: Menu Handle = " & m_oPic.hwnd
    #End If

    #If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 4218
        InfoPrint 1, timeGetTime & " clsCompList.ShowDropDown -- Stage 1: Focus Window Handle = " & GetFocus
    #End If

vbwProfiler.vbwExecuteLine 4219
    Call DrawItems(m_oPic, ByVal m_hStart)
vbwProfiler.vbwExecuteLine 4220
    SetForegroundWindow m_oPic.hwnd

    #If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 4221
        InfoPrint 1, timeGetTime & " clsCompList.ShowDropDown -- Stage 2: Focus Window Handle = " & GetFocus
    #End If

vbwProfiler.vbwExecuteLine 4222
    SetFocus m_oPic.hwnd

    #If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 4223
        InfoPrint 1, timeGetTime & " clsCompList.ShowDropDown -- Stage 3: Focus Window Handle = " & GetFocus
    #End If
vbwProfiler.vbwExecuteLine 4224
    ShowDropDown = True
vbwProfiler.vbwProcOut 213
vbwProfiler.vbwExecuteLine 4225
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 4226
    ShowDropDown = False
vbwProfiler.vbwProcOut 213
vbwProfiler.vbwExecuteLine 4227
End Function

Private Sub DrawItems(ByRef oPic As PictureBox, ByVal hAddress As Long)
vbwProfiler.vbwProcIn 214
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

vbwProfiler.vbwExecuteLine 4228
    oPic.DrawMode = 13  '  <-- damnit, I had this draw mode set to 1=BLACKNESS and couldnt figure out why the arrows wouldnt draw same color as pen/brush! argh

vbwProfiler.vbwExecuteLine 4229
    hMem = hAddress
vbwProfiler.vbwExecuteLine 4230
    lngNodeSize = LenB(oNode)
vbwProfiler.vbwExecuteLine 4231
    m_lngCellWidth = oPic.Width

'vbwLine 4232:    Do While hMem > 0
    Do While vbwProfiler.vbwExecuteLine(4232) Or hMem > 0
vbwProfiler.vbwExecuteLine 4233
        CopyMemory oNode, ByVal hMem, lngNodeSize
vbwProfiler.vbwExecuteLine 4234
        hMem = oNode.pNext

vbwProfiler.vbwExecuteLine 4235
        s = NodeBytesToString(oNode.bString)
vbwProfiler.vbwExecuteLine 4236
        l = Len(s)
vbwProfiler.vbwExecuteLine 4237
        yPos = i * m_lngCellHeight

vbwProfiler.vbwExecuteLine 4238
        With r
vbwProfiler.vbwExecuteLine 4239
            .Top = yPos - 1
vbwProfiler.vbwExecuteLine 4240
            .Bottom = yPos + m_lngCellHeight + 1
vbwProfiler.vbwExecuteLine 4241
            .Left = 0
vbwProfiler.vbwExecuteLine 4242
            .Right = m_lngCellWidth
vbwProfiler.vbwExecuteLine 4243
        End With

vbwProfiler.vbwExecuteLine 4244
        If oNode.bSelectedState = NODE_SELECTED Then
vbwProfiler.vbwExecuteLine 4245
            Call FillRect(ByVal oPic.hdc, r, m_hFocusRectBrush)
vbwProfiler.vbwExecuteLine 4246
            SetTextColor oPic.hdc, m_crColorHighlightText
vbwProfiler.vbwExecuteLine 4247
            SelectObject oPic.hdc, m_hHighlightArrowBrush
vbwProfiler.vbwExecuteLine 4248
            SelectObject oPic.hdc, m_hPen
        Else
vbwProfiler.vbwExecuteLine 4249 'B
vbwProfiler.vbwExecuteLine 4250
            Call FillRect(ByVal oPic.hdc, r, m_hDefaultBrush)
vbwProfiler.vbwExecuteLine 4251
            SetTextColor oPic.hdc, m_crColorWindowText
vbwProfiler.vbwExecuteLine 4252
            SelectObject oPic.hdc, m_hArrowBrush
vbwProfiler.vbwExecuteLine 4253
            SelectObject oPic.hdc, m_hOriginalPen
        End If
vbwProfiler.vbwExecuteLine 4254 'B

        ' only draw arrow if there is a branch off this node
vbwProfiler.vbwExecuteLine 4255
        If oNode.pBranch > 0 Then
vbwProfiler.vbwExecuteLine 4256
            Call NextArrow(i)
vbwProfiler.vbwExecuteLine 4257
            Call Polygon(ByVal oPic.hdc, m_ArrowPoints(0), 3)
        End If
vbwProfiler.vbwExecuteLine 4258 'B
vbwProfiler.vbwExecuteLine 4259
        Call TextOut(ByVal oPic.hdc, LEFT_CELL_PADDING, yPos, s, l)
vbwProfiler.vbwExecuteLine 4260
    i = i + 1
vbwProfiler.vbwExecuteLine 4261
    Loop
vbwProfiler.vbwProcOut 214
vbwProfiler.vbwExecuteLine 4262
End Sub

Private Sub NextArrow(ByVal iIndex As Long)
vbwProfiler.vbwProcIn 215
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngOffset As Long

vbwProfiler.vbwExecuteLine 4263
    lngLeft = 150
vbwProfiler.vbwExecuteLine 4264
    If iIndex = 0 Then
vbwProfiler.vbwExecuteLine 4265
        lngTop = 3
    Else
vbwProfiler.vbwExecuteLine 4266 'B
vbwProfiler.vbwExecuteLine 4267
        lngOffset = (m_lngCellHeight - ARROWHEIGHT) \ 2
vbwProfiler.vbwExecuteLine 4268
        lngTop = (iIndex * m_lngCellHeight) + lngOffset
    End If
vbwProfiler.vbwExecuteLine 4269 'B
vbwProfiler.vbwExecuteLine 4270
    m_ArrowPoints(0).x = lngLeft
vbwProfiler.vbwExecuteLine 4271
    m_ArrowPoints(0).y = lngTop
vbwProfiler.vbwExecuteLine 4272
    m_ArrowPoints(1).x = lngLeft + ARROWWIDTH
vbwProfiler.vbwExecuteLine 4273
    m_ArrowPoints(1).y = lngTop + ARROWMID
vbwProfiler.vbwExecuteLine 4274
    m_ArrowPoints(2).x = lngLeft
vbwProfiler.vbwExecuteLine 4275
    m_ArrowPoints(2).y = lngTop + ARROWHEIGHT
vbwProfiler.vbwProcOut 215
vbwProfiler.vbwExecuteLine 4276
End Sub

Private Function Load_Components() As Long
vbwProfiler.vbwProcIn 216
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

vbwProfiler.vbwExecuteLine 4277
    On Error GoTo errorhandler

    'make sure we are back in the program's install path
vbwProfiler.vbwExecuteLine 4278
    ChDir App.Path
vbwProfiler.vbwExecuteLine 4279
    Set oINI = New cINI
vbwProfiler.vbwExecuteLine 4280
    oINI.FileName = m_sFileName
vbwProfiler.vbwExecuteLine 4281
    sSections = oINI.RetreiveSectionNames

    ' check for zero array
vbwProfiler.vbwExecuteLine 4282
    m_lngCellCount = UBound(sSections)
vbwProfiler.vbwExecuteLine 4283
    If (m_lngCellCount = 0) And (sSections(0) = "") Then
vbwProfiler.vbwExecuteLine 4284
         GoTo errorhandler
    End If
vbwProfiler.vbwExecuteLine 4285 'B

    ' all is good, lets retrieve them and store them
vbwProfiler.vbwExecuteLine 4286
    lngNodeSize = LenB(oParent)
vbwProfiler.vbwExecuteLine 4287
    hAddress = m_hStart

    'todo:!!!! Check that our heap size is big enuf to contain LenB(uNode) * m_lngCellCount
vbwProfiler.vbwExecuteLine 4288
    For i = 0 To m_lngCellCount
vbwProfiler.vbwExecuteLine 4289
        b = StrConv(sSections(i), vbFromUnicode)
vbwProfiler.vbwExecuteLine 4290
        Erase oParent.bString
vbwProfiler.vbwExecuteLine 4291
        CopyMemory oParent.bString(0), b(0), Minimum(30, UBound(b) + 1)

        ' also retreive the subcategories for each main category
vbwProfiler.vbwExecuteLine 4292
        sKeyNames = oINI.ReadSection(sSections(i))
vbwProfiler.vbwExecuteLine 4293
        lKeyCount = UBound(sKeyNames)
vbwProfiler.vbwExecuteLine 4294
        If (lKeyCount = 0) And (sKeyNames(0) = "") Then
vbwProfiler.vbwExecuteLine 4295
             GoTo errorhandler
        End If
vbwProfiler.vbwExecuteLine 4296 'B

        ' create a heap that will hold this branch
vbwProfiler.vbwExecuteLine 4297
        lngBranchSize = (lKeyCount + 1) * lngNodeSize
vbwProfiler.vbwExecuteLine 4298
        hBranchAddress = HeapAlloc(m_hHeapSub, HEAP_ZERO_MEMORY, lngBranchSize)
vbwProfiler.vbwExecuteLine 4299
        oParent.pBranch = hBranchAddress
vbwProfiler.vbwExecuteLine 4300
        If i = m_lngCellCount Then
vbwProfiler.vbwExecuteLine 4301
            oParent.pNext = 0
        Else
vbwProfiler.vbwExecuteLine 4302 'B
vbwProfiler.vbwExecuteLine 4303
            oParent.pNext = hAddress + lngNodeSize
        End If
vbwProfiler.vbwExecuteLine 4304 'B
        'todo: check hBranch value for error in heap creation

        ' copy the uParent to our primary list before proceeding to branch copy
vbwProfiler.vbwExecuteLine 4305
        CopyMemory ByVal hAddress, oParent, lngNodeSize
vbwProfiler.vbwExecuteLine 4306
        CopyMemory oTemp, ByVal hAddress, lngNodeSize
vbwProfiler.vbwExecuteLine 4307
        hAddress = hAddress + lngNodeSize

        ' copy branch nodes subHeap
vbwProfiler.vbwExecuteLine 4308
        For j = 0 To lKeyCount
vbwProfiler.vbwExecuteLine 4309
            b = StrConv(oINI.ReadString(sSections(i), sKeyNames(j)), vbFromUnicode)
vbwProfiler.vbwExecuteLine 4310
            Erase oChild.bString
vbwProfiler.vbwExecuteLine 4311
            CopyMemory oChild.bString(0), b(0), Minimum(30, UBound(b) + 1)
vbwProfiler.vbwExecuteLine 4312
            oChild.pBranch = 0
vbwProfiler.vbwExecuteLine 4313
            If j = lKeyCount Then
vbwProfiler.vbwExecuteLine 4314
                oChild.pNext = 0
            Else
vbwProfiler.vbwExecuteLine 4315 'B
vbwProfiler.vbwExecuteLine 4316
                oChild.pNext = hBranchAddress + lngNodeSize
            End If
vbwProfiler.vbwExecuteLine 4317 'B
vbwProfiler.vbwExecuteLine 4318
            CopyMemory ByVal hBranchAddress, oChild, lngNodeSize
vbwProfiler.vbwExecuteLine 4319
            CopyMemory oTemp, ByVal hBranchAddress, lngNodeSize
vbwProfiler.vbwExecuteLine 4320
            hBranchAddress = hBranchAddress + lngNodeSize
vbwProfiler.vbwExecuteLine 4321
        Next
vbwProfiler.vbwExecuteLine 4322
    Next
vbwProfiler.vbwExecuteLine 4323
    Load_Components = True
vbwProfiler.vbwProcOut 216
vbwProfiler.vbwExecuteLine 4324
    Exit Function

errorhandler:
vbwProfiler.vbwExecuteLine 4325
    Load_Components = False
vbwProfiler.vbwProcOut 216
vbwProfiler.vbwExecuteLine 4326
End Function

Private Sub m_oPic_LostFocus()
vbwProfiler.vbwProcIn 217
vbwProfiler.vbwExecuteLine 4327
    m_oPic.Visible = False
vbwProfiler.vbwExecuteLine 4328
    m_oSubPic.Visible = False
vbwProfiler.vbwProcOut 217
vbwProfiler.vbwExecuteLine 4329
End Sub

Private Sub m_oPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 218
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
vbwProfiler.vbwExecuteLine 4330
    lngLastTime = timeGetTime

    ' if the time between moves is too fast, exit sub. We dont want this to be
    ' too sensitive
    #If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 4331
        InfoPrint 1, timeGetTime & " clsCompList:MouseMove -- mouse movement in main menu OK"
    #End If

vbwProfiler.vbwExecuteLine 4332
    m_lngTwipsX = VB.Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 4333
    lngNodeSize = LenB(m_oCurrentNode)
    ' determine which cell we are currently over
vbwProfiler.vbwExecuteLine 4334
    lngMouseOver = GetMouseOverCell(x, y)

vbwProfiler.vbwExecuteLine 4335
    If lngLastMouseOver <> lngMouseOver Then
vbwProfiler.vbwExecuteLine 4336
        With m_oSubPic
        ' reposition the sub list
vbwProfiler.vbwExecuteLine 4337
            .Visible = False
vbwProfiler.vbwExecuteLine 4338
            .Top = m_oPic.Top + m_lngTwipsX * (m_lngCellHeight * lngMouseOver)
            '.Height = VB.Screen.TwipsPerPixelY * ((m_Items(m_lngMouseOver).iSubCount + 1) * m_lngCellHeight)
vbwProfiler.vbwExecuteLine 4339
            .Left = m_oPic.Left + m_oPic.Width
vbwProfiler.vbwExecuteLine 4340
            .Width = m_oPic.Width
vbwProfiler.vbwExecuteLine 4341
            .Visible = True
vbwProfiler.vbwExecuteLine 4342
        End With

        ' update the selected state of the previous node in the heap
vbwProfiler.vbwExecuteLine 4343
        hAddress = m_hStart + (lngLastMouseOver * lngNodeSize)
vbwProfiler.vbwExecuteLine 4344
        CopyMemory m_oCurrentNode, ByVal hAddress, lngNodeSize
vbwProfiler.vbwExecuteLine 4345
        m_oCurrentNode.bSelectedState = NODE_DESELECTED
vbwProfiler.vbwExecuteLine 4346
        CopyMemory ByVal hAddress, m_oCurrentNode, lngNodeSize
    End If
vbwProfiler.vbwExecuteLine 4347 'B
vbwProfiler.vbwExecuteLine 4348
    lngLastMouseOver = lngMouseOver

    ' get the address for the subcategory to draw
vbwProfiler.vbwExecuteLine 4349
    hAddress = m_hStart + (lngMouseOver * lngNodeSize)
vbwProfiler.vbwExecuteLine 4350
    CopyMemory m_oCurrentNode, ByVal hAddress, lngNodeSize
    ' update the selected state of the node in the heap
vbwProfiler.vbwExecuteLine 4351
    m_oCurrentNode.bSelectedState = NODE_SELECTED
vbwProfiler.vbwExecuteLine 4352
    m_oPic.Tag = NodeBytesToString(m_oCurrentNode.bString)
vbwProfiler.vbwExecuteLine 4353
    CopyMemory ByVal hAddress, m_oCurrentNode, lngNodeSize

vbwProfiler.vbwExecuteLine 4354
    hBranchStart = m_oCurrentNode.pBranch
vbwProfiler.vbwExecuteLine 4355
    hAddress = hBranchStart

    ' determine how many nodes are in this branch so we can determine picbox height
'vbwLine 4356:    Do While hAddress > 0
    Do While vbwProfiler.vbwExecuteLine(4356) Or hAddress > 0
vbwProfiler.vbwExecuteLine 4357
        CopyMemory oTempNode, ByVal hAddress, lngNodeSize
vbwProfiler.vbwExecuteLine 4358
        hAddress = oTempNode.pNext
vbwProfiler.vbwExecuteLine 4359
        i = i + 1
vbwProfiler.vbwExecuteLine 4360
    Loop
vbwProfiler.vbwExecuteLine 4361
    With m_oSubPic
vbwProfiler.vbwExecuteLine 4362
        .Height = m_lngTwipsX * ((i) * m_lngCellHeight)
vbwProfiler.vbwExecuteLine 4363
        .Tag = hBranchStart
vbwProfiler.vbwExecuteLine 4364
    End With
    ' set our mem pointer back to start of the selected branch and draw it
vbwProfiler.vbwExecuteLine 4365
    Call DrawItems(m_oSubPic, hBranchStart)
    ' draw both dropdowns
vbwProfiler.vbwExecuteLine 4366
    Call DrawItems(m_oPic, m_hStart)
vbwProfiler.vbwProcOut 218
vbwProfiler.vbwExecuteLine 4367
End Sub

Private Function GetMouseOverCell(x, y) As Long
vbwProfiler.vbwProcIn 219
    Dim l As Long
    Dim i As Long

vbwProfiler.vbwExecuteLine 4368
    l = CLng(y)
vbwProfiler.vbwExecuteLine 4369
    l = l \ m_lngCellHeight
vbwProfiler.vbwExecuteLine 4370
    GetMouseOverCell = l
#If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 4371
    InfoPrint 1, timeGetTime & "clsCompList:GetMouseOverCell -- detected cell = " & l
#End If
vbwProfiler.vbwProcOut 219
vbwProfiler.vbwExecuteLine 4372
End Function

Private Sub m_oPic_Paint()
vbwProfiler.vbwProcIn 220
vbwProfiler.vbwExecuteLine 4373
    Call DrawItems(m_oPic, m_hStart)
vbwProfiler.vbwProcOut 220
vbwProfiler.vbwExecuteLine 4374
End Sub

Private Function NodeBytesToString(b() As Byte) As String
vbwProfiler.vbwProcIn 221
    Dim s As String
    Dim l As Long

vbwProfiler.vbwExecuteLine 4375
    s = Trim(StrConv(b, vbUnicode))
vbwProfiler.vbwExecuteLine 4376
    l = InStr(1, s, Chr(0))
vbwProfiler.vbwExecuteLine 4377
    s = Left(s, l - 1)
vbwProfiler.vbwExecuteLine 4378
    NodeBytesToString = s
vbwProfiler.vbwProcOut 221
vbwProfiler.vbwExecuteLine 4379
End Function

Private Sub m_oSubPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 222
    Dim l As Long
    Dim oNode As uNode
    Dim lngNodeSize As Long
    Dim s As String
    Dim hAddress As Long

vbwProfiler.vbwExecuteLine 4380
    lngNodeSize = LenB(oNode)
vbwProfiler.vbwExecuteLine 4381
    l = GetMouseOverCell(x, y)
vbwProfiler.vbwExecuteLine 4382
    hAddress = CLng(m_oSubPic.Tag) + (l * lngNodeSize)

vbwProfiler.vbwExecuteLine 4383
    CopyMemory oNode, ByVal hAddress, lngNodeSize
vbwProfiler.vbwExecuteLine 4384
    s = NodeBytesToString(oNode.bString)

vbwProfiler.vbwExecuteLine 4385
    m_oPic.Visible = False
vbwProfiler.vbwExecuteLine 4386
    m_oSubPic.Visible = False
vbwProfiler.vbwExecuteLine 4387
    RaiseEvent Click(m_oPic.Tag & "\" & s)
vbwProfiler.vbwProcOut 222
vbwProfiler.vbwExecuteLine 4388
End Sub

Private Sub m_oSubPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 223
    Dim lngMouseOver As Long
    Static lngLastMouseOver As Long
    Dim hAddress As Long
    Dim lngNodeSize As Long
    Dim oTempNode As uNode
    Static lngLastTime As Long

    'todo: this will not work on machines that have been running long enuf for timegettime to wrap around
    ' we will get negative numbers and so,
    'If timeGetTime - lngLastTime < MAXIMUM_POLL_RATE Then Exit Sub
vbwProfiler.vbwExecuteLine 4389
    lngLastTime = timeGetTime

vbwProfiler.vbwExecuteLine 4390
    lngNodeSize = LenB(m_oCurrentNode)
    ' determine which cell we are currently over
vbwProfiler.vbwExecuteLine 4391
    lngMouseOver = GetMouseOverCell(x, y)

vbwProfiler.vbwExecuteLine 4392
    If m_oCurrentNode.pBranch <= 0 Then
vbwProfiler.vbwExecuteLine 4393
        Debug.Print "clsComplist:m_oSubPic_MouseMove() -- Branch node address not set.  Exiting Sub"
vbwProfiler.vbwProcOut 223
vbwProfiler.vbwExecuteLine 4394
        Exit Sub
    End If
vbwProfiler.vbwExecuteLine 4395 'B

    #If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 4396
        InfoPrint 1, timeGetTime & " clsCompList:SubMenuMouseMove -- mouse movement detected OK"
    #End If

    ' update the selected state of the previous node in the heap
vbwProfiler.vbwExecuteLine 4397
    hAddress = m_oCurrentNode.pBranch + (lngLastMouseOver * lngNodeSize)
vbwProfiler.vbwExecuteLine 4398
    If hAddress >= 0 Then
vbwProfiler.vbwExecuteLine 4399
        CopyMemory oTempNode, ByVal hAddress, lngNodeSize
vbwProfiler.vbwExecuteLine 4400
        oTempNode.bSelectedState = NODE_DESELECTED
vbwProfiler.vbwExecuteLine 4401
        CopyMemory ByVal hAddress, oTempNode, lngNodeSize
    End If
vbwProfiler.vbwExecuteLine 4402 'B

vbwProfiler.vbwExecuteLine 4403
    lngLastMouseOver = lngMouseOver

    ' get the address for the subcategory to draw
vbwProfiler.vbwExecuteLine 4404
    hAddress = m_oCurrentNode.pBranch + (lngMouseOver * lngNodeSize)
vbwProfiler.vbwExecuteLine 4405
    CopyMemory oTempNode, ByVal hAddress, lngNodeSize
    ' update the selected state of the node in the heap
vbwProfiler.vbwExecuteLine 4406
    oTempNode.bSelectedState = NODE_SELECTED
vbwProfiler.vbwExecuteLine 4407
    CopyMemory ByVal hAddress, oTempNode, lngNodeSize

    ' set our mem pointer back to start of the selected branch and draw it
vbwProfiler.vbwExecuteLine 4408
    Call DrawItems(m_oSubPic, m_oCurrentNode.pBranch)

    ' deselect the item for next time
vbwProfiler.vbwExecuteLine 4409
    CopyMemory oTempNode, ByVal hAddress, lngNodeSize
vbwProfiler.vbwExecuteLine 4410
    oTempNode.bSelectedState = NODE_DESELECTED
vbwProfiler.vbwExecuteLine 4411
    CopyMemory ByVal hAddress, oTempNode, lngNodeSize
vbwProfiler.vbwProcOut 223
vbwProfiler.vbwExecuteLine 4412
End Sub


