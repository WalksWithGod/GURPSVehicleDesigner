Option Strict Off
Option Explicit On
Friend Class clsCompList
	
	Private Structure Rect
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	
	Private Structure POINTAPI
		Dim x As Integer
		Dim y As Integer
	End Structure
	
	'Private Type m_Cell
	'    r As RECT
	'    text As String
	'    iMore As Long   ' array index for extended drop down list
	'    iSubCount As Long ' number of subitems
	'End Type
	
	'-------------------
	' CONVERT TO HEAP
	Private Structure uNode
		<VBFixedArray(30)> Dim bString() As Byte
		Dim pBranch As Integer 'pointer to start of sublist heap
		Dim pNext As Integer ' pointer to next node
		Dim bSelectedState As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim bString(30)
		End Sub
	End Structure
	
	Const NODE_SELECTED As Short = 1
	Const NODE_DESELECTED As Short = 0
	
	
	Private m_hHeap As Integer
	Private m_hStart As Integer ' pointer to start of primary list
	Private m_hHeapSub As Integer
	'UPGRADE_WARNING: Arrays in structure m_oCurrentNode may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Private m_oCurrentNode As uNode
	
	Const PAGE_SIZE As Short = 4096 'only on Alphas, the page size is 8196
	Const HEAP_SIZE As Short = 8192
	Const MAX_HEAP_SIZE As Short = 16384
	
	Const MAXIMUM_POLL_RATE As Short = 25
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef hpvDest As Any, ByRef hpvSource As Any, ByVal cbCopy As Integer)
	Private Declare Function timeGetTime Lib "winmm.dll" () As Integer
	'-------------------
	'Private m_Items() As m_Cell
	'Private m_SubItems() As String ' two dimensional array.  First subscript is index to m_Items, 2nd subscript contains actual data
	Private m_ArrowPoints(2) As POINTAPI
	
	
	
	Private Const PT_MOVETO As Short = &H6s
	Private Const PT_LINETO As Short = &H2s
	Private Const PT_CLOSEFIGURE As Short = &H1s
	Private Const PT_BEZIERTO As Short = &H4s
	
	Private m_lngMaxWidth As Integer
	Private m_lngMaxHeight As Integer
	Private m_lngCellHeight As Integer
	Private m_lngCellCount As Integer
	Private m_lngCellWidth As Integer
	Private m_sFileName As String
	
	Private m_lngTwipsX As Integer
	Private m_lngTwipsY As Integer
	
	Private m_crColorWindowText As Integer
	Private m_crColorHighlightText As Integer
	Private m_hDefaultBrush As Integer
	Private m_hFocusRectBrush As Integer ' brush used for solid rect around selected item
	Private m_hOriginalBrush As Integer ' brush for non selected item rectangle
	Private m_hHighlightArrowBrush As Integer ' brush for the arrow so that it matches the color of the highlight text
	Private m_hArrowBrush As Integer
	Private m_hPen As Integer
	Private m_hOriginalPen As Integer
	
	Private WithEvents m_oPic As System.Windows.Forms.PictureBox
	Private WithEvents m_oSubPic As System.Windows.Forms.PictureBox
	
	Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Integer) As Integer
	Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Integer) As Integer
	Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Integer) As Integer
	Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Integer) As Integer
	Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
	Private Declare Function TextOut Lib "gdi32"  Alias "TextOutA"(ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
	'UPGRADE_WARNING: Structure POINTAPI may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Integer, ByRef lpPoint As POINTAPI, ByVal nCount As Integer) As Integer
	'UPGRADE_WARNING: Structure Rect may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FillRect Lib "user32" (ByVal hdc As Integer, ByRef lpRect As Rect, ByVal hBrush As Integer) As Integer
	
#If DEBUG_MODE Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private Declare Function GetActiveWindow Lib "user32" () As Long
	Private Declare Function GetFocus Lib "user32" () As Long
#End If
	Public Event Click(ByRef s As String)
	
	Const LEFT_CELL_PADDING As Short = 2
	Const TOP_CELL_PADDING As Short = 3
	Const MAX_SUBITEMS As Short = 32
	Const UNIQUE_KEY As String = "DROPDOWN"
	Const UNIQUE_KEY_2 As String = "SUBDROPDOWN"
	
	Const ARROWHEIGHT As Short = 6
	Const ARROWWIDTH As Short = 5
	Const ARROWMID As Short = 3
	
	
	' returns TRUE if successful
	Public Function SetFileName(ByRef s As String) As Integer
		Dim lRet As Integer
		If s <> "" Then ' todo: should maybe call routine to check if IsFile or something
			m_sFileName = s
			SetFileName = True
			
			' load lists
			If Not Load_Components Then
				MsgBox("Error loading vehicle components list.")
				Call Class_Terminate_Renamed()
			End If
		Else
			SetFileName = False
		End If
	End Function
	
	' This class handles the reading and listing of components
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Dim PS_SOLID As Object
		Dim CreatePen As Object
		Dim COLOR_HIGHLIGHT As Object
		Dim COLOR_WINDOW As Object
		Dim COLOR_HIGHLIGHTTEXT As Object
		Dim COLOR_WINDOWTEXT As Object
		Dim GetSysColor As Object
		Dim vbFlat As Object
		Dim frmDesigner As Object
		Dim lRet As Integer
		Dim lngTempHeapSize As Integer
		
		On Error GoTo errorhandler
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.Controls. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_oPic = frmDesigner.Controls.Add("VB.PictureBox", UNIQUE_KEY)
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.Controls. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_oSubPic = frmDesigner.Controls.Add("VB.PictureBox", UNIQUE_KEY_2)
		With m_oPic
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Font = VB6.FontChangeName(.Font, frmDesigner.treeVehicle.Font)
			.BringToFront()
			'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			'UPGRADE_ISSUE: PictureBox property m_oPic.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.ScaleMode = vbPixels
			.BorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
			'UPGRADE_ISSUE: PictureBox property m_oPic.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vbFlat. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Appearance = vbFlat
			.Visible = False
		End With
		
		With m_oSubPic
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Font = VB6.FontChangeName(.Font, frmDesigner.treeVehicle.Font)
			.BringToFront()
			'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			'UPGRADE_ISSUE: PictureBox property m_oSubPic.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.ScaleMode = vbPixels
			.BorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
			'UPGRADE_ISSUE: PictureBox property m_oSubPic.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vbFlat. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Appearance = vbFlat
			.Visible = False
		End With
		
		' create two brushes which we use to draw our selected rect
		'UPGRADE_WARNING: Couldn't resolve default property of object GetSysColor(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_crColorWindowText = GetSysColor(COLOR_WINDOWTEXT)
		'UPGRADE_WARNING: Couldn't resolve default property of object GetSysColor(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_crColorHighlightText = GetSysColor(COLOR_HIGHLIGHTTEXT)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object GetSysColor(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_hDefaultBrush = CreateSolidBrush(GetSysColor(COLOR_WINDOW))
		'UPGRADE_ISSUE: PictureBox property m_oPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		m_hOriginalBrush = SelectObject(m_oPic.hdc, m_hDefaultBrush) 'todo: im still not loading a brush for the sub pic, note that the arrow we draw using the polygon function, uses the currently loaded brush for its color and NOT textcolor
		'UPGRADE_WARNING: Couldn't resolve default property of object GetSysColor(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_hFocusRectBrush = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
		m_hHighlightArrowBrush = CreateSolidBrush(m_crColorHighlightText)
		m_hArrowBrush = CreateSolidBrush(m_crColorWindowText)
		
		' try these pens to see if they fix problems we are having drawing our fucking arrows
		'UPGRADE_WARNING: Couldn't resolve default property of object CreatePen(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_hPen = CreatePen(PS_SOLID, 1, m_crColorHighlightText)
		'UPGRADE_ISSUE: PictureBox property m_oPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		m_hOriginalPen = SelectObject(m_oPic.hdc, m_hPen)
		' get the font height for determining the cell height
		'UPGRADE_ISSUE: PictureBox method m_oPic.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		m_lngMaxHeight = m_oPic.TextHeight("W")
		m_lngCellHeight = m_lngMaxHeight + 4
		
		' create our primary heap
		m_hHeap = modHeaps.HeapCreate(HEAP_NO_SERIALIZE Or HEAP_GENERATE_EXCEPTIONS, HEAP_SIZE, 0)
		Debug.Print("Heap Created @ " & m_hHeap)
		m_hStart = modHeaps.HeapAlloc(m_hHeap, HEAP_ZERO_MEMORY, HEAP_SIZE)
		Debug.Print("Heap Allocated @ " & m_hStart)
		lngTempHeapSize = modHeaps.HeapSize(m_hHeap, 0, m_hStart)
		Debug.Print("Heap Size = " & lngTempHeapSize)
		
		m_hHeapSub = modHeaps.HeapCreate(HEAP_NO_SERIALIZE Or HEAP_GENERATE_EXCEPTIONS, MAX_HEAP_SIZE, 0)
		If m_hStart <= 0 Then GoTo errorhandler
		Exit Sub
		
errorhandler: 
		MsgBox("Could not ceate custom drop down menu")
		Call Class_Terminate_Renamed()
		Exit Sub
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		' erase arrays and destroy object
		Dim l As Integer
		Dim b As Boolean
		
		'todo: the way im handling original pens and brushes really only works for the main pic and not the sub pic.
		' ideally id like to make it so i can have an arbitrary number of sublists.  This would also mean each
		' subllist would use the same code and cleanup would be automatic
		'UPGRADE_ISSUE: PictureBox property m_oPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		l = SelectObject(m_oPic.hdc, m_hOriginalBrush)
		l = DeleteObject(m_hDefaultBrush)
		l = DeleteObject(m_hFocusRectBrush)
		
		'UPGRADE_NOTE: Object m_oPic may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_oPic.Image = Nothing
		'UPGRADE_NOTE: Object m_oSubPic may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_oSubPic.Image = Nothing
		
		b = DestroySubHeaps
		Debug.Print("DestroySubHeaps Returns " & b)
		
		modHeaps.HeapFree(m_hHeap, 0, m_hStart)
		modHeaps.HeapDestroy(m_hHeap)
		modHeaps.HeapDestroy(m_hHeapSub)
		
#If DEBUG_MODE Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		frmDesigner.txtInfo.Text = "clsCompList:Terminate() -- Successfully terminated custom dropdown menu control"
#End If
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private Function DestroySubHeaps() As Boolean
		On Error GoTo err_Renamed
		
		Dim i As Integer
		'UPGRADE_WARNING: Arrays in structure oNode may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim oNode As uNode
		Dim hAddress As Integer
		Dim lngNodeSize As Integer
		Dim l As Integer
		
		' goes through each parent node in the primary heap, finds the offset (pBranch) in the secondary heap
		' which is the pointer to the specific subheap and frees it.
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		lngNodeSize = LenB(oNode)
		hAddress = m_hStart
		
		For i = 0 To m_lngCellCount
			'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oNode, hAddress, lngNodeSize)
			l = modHeaps.HeapFree(m_hHeapSub, 0, oNode.pBranch)
			Debug.Print("clsCompList:DestroySubHeaps -- Itteration = " & i & " HeapFree Returning " & l)
			hAddress = oNode.pNext
		Next 
		DestroySubHeaps = True
		Exit Function
err_Renamed: 
		DestroySubHeaps = False
	End Function
	
	Public Function ShowDropDown() As Integer
		Dim frmDesigner As Object
		Dim i As Integer
		
		On Error GoTo errorhandler
		
		m_oPic.Visible = True
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.cboComponents. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_oPic.Top = VB6.TwipsToPixelsY(frmDesigner.cboComponents.Top + frmDesigner.cboComponents.Height)
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.cboComponents. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_oPic.Left = VB6.TwipsToPixelsX(frmDesigner.cboComponents.Left)
		m_oPic.Height = VB6.TwipsToPixelsY(VB6.TwipsPerPixelX * (m_lngCellHeight * (m_lngCellCount + 1)))
		m_oPic.Width = VB6.TwipsToPixelsX(2500) ' m_lngCellWidth
		m_oPic.BringToFront()
		m_oSubPic.BringToFront()
		
		' temp: paint dropdown
#If DEBUG_MODE Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		InfoPrint 1, timeGetTime & " clsCompList.ShowDropDown -- Stage 0: Menu Handle = " & m_oPic.hwnd
#End If
		
#If DEBUG_MODE Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		InfoPrint 1, timeGetTime & " clsCompList.ShowDropDown -- Stage 1: Focus Window Handle = " & GetFocus
#End If
		
		Call DrawItems(m_oPic, m_hStart)
		SetForegroundWindow(m_oPic.Handle.ToInt32)
		
#If DEBUG_MODE Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		InfoPrint 1, timeGetTime & " clsCompList.ShowDropDown -- Stage 2: Focus Window Handle = " & GetFocus
#End If
		
		SetFocus(m_oPic.Handle.ToInt32)
		
#If DEBUG_MODE Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		InfoPrint 1, timeGetTime & " clsCompList.ShowDropDown -- Stage 3: Focus Window Handle = " & GetFocus
#End If
		ShowDropDown = True
		Exit Function
errorhandler: 
		ShowDropDown = False
	End Function
	
	Private Sub DrawItems(ByRef oPic As System.Windows.Forms.PictureBox, ByVal hAddress As Integer)
		Dim SetTextColor As Object
		Dim i As Integer
		Dim l As Integer
		Dim s As String
		Dim yPos As Integer
		Dim r As Rect
		'UPGRADE_WARNING: Arrays in structure oNode may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim oNode As uNode
		Dim lngNodeSize As Integer
		Dim hMem As Integer
		Dim j As Integer
		Dim lRet As Integer
		Dim hOriginalPen As Integer
		
		'UPGRADE_ISSUE: PictureBox property oPic.DrawMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		oPic.DrawMode = 13 '  <-- damnit, I had this draw mode set to 1=BLACKNESS and couldnt figure out why the arrows wouldnt draw same color as pen/brush! argh
		
		hMem = hAddress
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		lngNodeSize = LenB(oNode)
		m_lngCellWidth = VB6.PixelsToTwipsX(oPic.Width)
		
		Do While hMem > 0
			'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oNode, hMem, lngNodeSize)
			hMem = oNode.pNext
			
			s = NodeBytesToString(oNode.bString)
			l = Len(s)
			yPos = i * m_lngCellHeight
			
			With r
				.Top = yPos - 1
				.Bottom = yPos + m_lngCellHeight + 1
				.Left_Renamed = 0
				.Right_Renamed = m_lngCellWidth
			End With
			
			If oNode.bSelectedState = NODE_SELECTED Then
				'UPGRADE_ISSUE: PictureBox property oPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Call FillRect(oPic.hdc, r, m_hFocusRectBrush)
				'UPGRADE_ISSUE: PictureBox property oPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetTextColor(oPic.hdc, m_crColorHighlightText)
				'UPGRADE_ISSUE: PictureBox property oPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SelectObject(oPic.hdc, m_hHighlightArrowBrush)
				'UPGRADE_ISSUE: PictureBox property oPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SelectObject(oPic.hdc, m_hPen)
			Else
				'UPGRADE_ISSUE: PictureBox property oPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Call FillRect(oPic.hdc, r, m_hDefaultBrush)
				'UPGRADE_ISSUE: PictureBox property oPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetTextColor(oPic.hdc, m_crColorWindowText)
				'UPGRADE_ISSUE: PictureBox property oPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SelectObject(oPic.hdc, m_hArrowBrush)
				'UPGRADE_ISSUE: PictureBox property oPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SelectObject(oPic.hdc, m_hOriginalPen)
			End If
			
			' only draw arrow if there is a branch off this node
			If oNode.pBranch > 0 Then
				Call NextArrow(i)
				'UPGRADE_ISSUE: PictureBox property oPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Call Polygon(oPic.hdc, m_ArrowPoints(0), 3)
			End If
			'UPGRADE_ISSUE: PictureBox property oPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call TextOut(oPic.hdc, LEFT_CELL_PADDING, yPos, s, l)
			i = i + 1
		Loop 
	End Sub
	
	Private Sub NextArrow(ByVal iIndex As Integer)
		Dim lngLeft As Integer
		Dim lngTop As Integer
		Dim lngOffset As Integer
		
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
	
	Private Function Load_Components() As Integer
		Dim sSections() As String 'main sections
		Dim sKeys() As String 'sub categories for each section
		Dim sKeyNames() As String
		Dim lKeyCount As Integer
		Dim i As Short
		Dim j As Integer
		Dim oINI As cINI
		'UPGRADE_WARNING: Arrays in structure oParent may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim oParent As uNode
		'UPGRADE_WARNING: Arrays in structure oChild may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim oChild As uNode
		Dim hAddress As Integer
		Dim lngNodeSize As Integer
		Dim lngBranchSize As Integer
		Dim hBranchAddress As Integer
		'UPGRADE_WARNING: Arrays in structure oTemp may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim oTemp As uNode
		Dim lngTempHeapSize As Integer
		Dim b() As Byte
		
		On Error GoTo errorhandler
		
		'make sure we are back in the program's install path
		ChDir(My.Application.Info.DirectoryPath)
		oINI = New cINI
		oINI.FileName = m_sFileName
		sSections = VB6.CopyArray(oINI.RetreiveSectionNames)
		
		' check for zero array
		m_lngCellCount = UBound(sSections)
		If (m_lngCellCount = 0) And (sSections(0) = "") Then GoTo errorhandler
		
		' all is good, lets retrieve them and store them
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		lngNodeSize = LenB(oParent)
		hAddress = m_hStart
		
		'todo:!!!! Check that our heap size is big enuf to contain LenB(uNode) * m_lngCellCount
		For i = 0 To m_lngCellCount
			'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
			b = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(sSections(i), vbFromUnicode))
			'UPGRADE_NOTE: Erase was upgraded to System.Array.Clear. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			System.Array.Clear(oParent.bString, 0, oParent.bString.Length)
			'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(30, UBound(b) + 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oParent.bString(0), b(0), modPerformance.Minimum(30, UBound(b) + 1))
			
			' also retreive the subcategories for each main category
			sKeyNames = oINI.ReadSection(sSections(i))
			lKeyCount = UBound(sKeyNames)
			If (lKeyCount = 0) And (sKeyNames(0) = "") Then GoTo errorhandler
			
			' create a heap that will hold this branch
			lngBranchSize = (lKeyCount + 1) * lngNodeSize
			hBranchAddress = modHeaps.HeapAlloc(m_hHeapSub, HEAP_ZERO_MEMORY, lngBranchSize)
			oParent.pBranch = hBranchAddress
			If i = m_lngCellCount Then
				oParent.pNext = 0
			Else
				oParent.pNext = hAddress + lngNodeSize
			End If
			'todo: check hBranch value for error in heap creation
			
			' copy the uParent to our primary list before proceeding to branch copy
			'UPGRADE_WARNING: Couldn't resolve default property of object oParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(hAddress, oParent, lngNodeSize)
			'UPGRADE_WARNING: Couldn't resolve default property of object oTemp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oTemp, hAddress, lngNodeSize)
			hAddress = hAddress + lngNodeSize
			
			' copy branch nodes subHeap
			For j = 0 To lKeyCount
				'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
				'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
				b = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(oINI.ReadString(sSections(i), sKeyNames(j)), vbFromUnicode))
				'UPGRADE_NOTE: Erase was upgraded to System.Array.Clear. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
				System.Array.Clear(oChild.bString, 0, oChild.bString.Length)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(30, UBound(b) + 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CopyMemory(oChild.bString(0), b(0), modPerformance.Minimum(30, UBound(b) + 1))
				oChild.pBranch = 0
				If j = lKeyCount Then
					oChild.pNext = 0
				Else
					oChild.pNext = hBranchAddress + lngNodeSize
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object oChild. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CopyMemory(hBranchAddress, oChild, lngNodeSize)
				'UPGRADE_WARNING: Couldn't resolve default property of object oTemp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CopyMemory(oTemp, hBranchAddress, lngNodeSize)
				hBranchAddress = hBranchAddress + lngNodeSize
			Next 
		Next 
		Load_Components = True
		Exit Function
		
errorhandler: 
		Load_Components = False
	End Function
	
	'UPGRADE_ISSUE: PictureBox event m_oPic.LostFocus was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub m_oPic_LostFocus()
		m_oPic.Visible = False
		m_oSubPic.Visible = False
	End Sub
	
	Private Sub m_oPic_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles m_oPic.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim lngNodeSize As Integer
		Dim i As Integer
		Dim hAddress As Integer
		Dim hBranchStart As Integer
		Dim lngMouseOver As Integer
		Static lngLastMouseOver As Integer
		'UPGRADE_WARNING: Arrays in structure oTempNode may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim oTempNode As uNode
		Static lngLastTime As Integer
		
		'todo: this returns a negative number if computer is running more than 23 days.
		'      queryperformancecounter doesnt work on all cpu's or old versions of windows95 either i think.
		'      need a better timer.  i could code around it by checking for neg values and in that case, subtract last from current
		'If timeGetTime - lngLastTime < MAXIMUM_POLL_RATE Then Exit Sub
		lngLastTime = timeGetTime
		
		' if the time between moves is too fast, exit sub. We dont want this to be
		' too sensitive
#If DEBUG_MODE Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		InfoPrint 1, timeGetTime & " clsCompList:MouseMove -- mouse movement in main menu OK"
#End If
		
		m_lngTwipsX = VB6.TwipsPerPixelX
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		lngNodeSize = LenB(m_oCurrentNode)
		' determine which cell we are currently over
		lngMouseOver = GetMouseOverCell(x, y)
		
		If lngLastMouseOver <> lngMouseOver Then
			With m_oSubPic
				' reposition the sub list
				.Visible = False
				.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(m_oPic.Top) + m_lngTwipsX * (m_lngCellHeight * lngMouseOver))
				'.Height = VB.Screen.TwipsPerPixelY * ((m_Items(m_lngMouseOver).iSubCount + 1) * m_lngCellHeight)
				.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(m_oPic.Left) + VB6.PixelsToTwipsX(m_oPic.Width))
				.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(m_oPic.Width))
				.Visible = True
			End With
			
			' update the selected state of the previous node in the heap
			hAddress = m_hStart + (lngLastMouseOver * lngNodeSize)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(m_oCurrentNode, hAddress, lngNodeSize)
			m_oCurrentNode.bSelectedState = NODE_DESELECTED
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(hAddress, m_oCurrentNode, lngNodeSize)
		End If
		lngLastMouseOver = lngMouseOver
		
		' get the address for the subcategory to draw
		hAddress = m_hStart + (lngMouseOver * lngNodeSize)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(m_oCurrentNode, hAddress, lngNodeSize)
		' update the selected state of the node in the heap
		m_oCurrentNode.bSelectedState = NODE_SELECTED
		m_oPic.Tag = NodeBytesToString(m_oCurrentNode.bString)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(hAddress, m_oCurrentNode, lngNodeSize)
		
		hBranchStart = m_oCurrentNode.pBranch
		hAddress = hBranchStart
		
		' determine how many nodes are in this branch so we can determine picbox height
		Do While hAddress > 0
			'UPGRADE_WARNING: Couldn't resolve default property of object oTempNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oTempNode, hAddress, lngNodeSize)
			hAddress = oTempNode.pNext
			i = i + 1
		Loop 
		With m_oSubPic
			.Height = VB6.TwipsToPixelsY(m_lngTwipsX * ((i) * m_lngCellHeight))
			.Tag = hBranchStart
		End With
		' set our mem pointer back to start of the selected branch and draw it
		Call DrawItems(m_oSubPic, hBranchStart)
		' draw both dropdowns
		Call DrawItems(m_oPic, m_hStart)
	End Sub
	
	Private Function GetMouseOverCell(ByRef x As Object, ByRef y As Object) As Integer
		Dim l As Integer
		Dim i As Integer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		l = CInt(y)
		l = l \ m_lngCellHeight
		GetMouseOverCell = l
#If DEBUG_MODE Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		InfoPrint 1, timeGetTime & "clsCompList:GetMouseOverCell -- detected cell = " & l
#End If
	End Function
	
	Private Sub m_oPic_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles m_oPic.Paint
		Call DrawItems(m_oPic, m_hStart)
	End Sub
	
	Private Function NodeBytesToString(ByRef b() As Byte) As String
		Dim s As String
		Dim l As Integer
		
		'UPGRADE_ISSUE: Constant vbUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		s = Trim(StrConv(System.Text.UnicodeEncoding.Unicode.GetString(b), vbUnicode))
		l = InStr(1, s, Chr(0))
		s = Left(s, l - 1)
		NodeBytesToString = s
	End Function
	
	Private Sub m_oSubPic_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles m_oSubPic.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim l As Integer
		'UPGRADE_WARNING: Arrays in structure oNode may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim oNode As uNode
		Dim lngNodeSize As Integer
		Dim s As String
		Dim hAddress As Integer
		
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		lngNodeSize = LenB(oNode)
		l = GetMouseOverCell(x, y)
		hAddress = CInt(m_oSubPic.Tag) + (l * lngNodeSize)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, hAddress, lngNodeSize)
		s = NodeBytesToString(oNode.bString)
		
		m_oPic.Visible = False
		m_oSubPic.Visible = False
		RaiseEvent Click(m_oPic.Tag & "\" & s)
	End Sub
	
	Private Sub m_oSubPic_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles m_oSubPic.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim lngMouseOver As Integer
		Static lngLastMouseOver As Integer
		Dim hAddress As Integer
		Dim lngNodeSize As Integer
		'UPGRADE_WARNING: Arrays in structure oTempNode may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim oTempNode As uNode
		Static lngLastTime As Integer
		
		'todo: this will not work on machines that have been running long enuf for timegettime to wrap around
		' we will get negative numbers and so,
		'If timeGetTime - lngLastTime < MAXIMUM_POLL_RATE Then Exit Sub
		lngLastTime = timeGetTime
		
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		lngNodeSize = LenB(m_oCurrentNode)
		' determine which cell we are currently over
		lngMouseOver = GetMouseOverCell(x, y)
		
		If m_oCurrentNode.pBranch <= 0 Then
			Debug.Print("clsComplist:m_oSubPic_MouseMove() -- Branch node address not set.  Exiting Sub")
			Exit Sub
		End If
		
#If DEBUG_MODE Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		InfoPrint 1, timeGetTime & " clsCompList:SubMenuMouseMove -- mouse movement detected OK"
#End If
		
		' update the selected state of the previous node in the heap
		hAddress = m_oCurrentNode.pBranch + (lngLastMouseOver * lngNodeSize)
		If hAddress >= 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oTempNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oTempNode, hAddress, lngNodeSize)
			oTempNode.bSelectedState = NODE_DESELECTED
			'UPGRADE_WARNING: Couldn't resolve default property of object oTempNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(hAddress, oTempNode, lngNodeSize)
		End If
		
		lngLastMouseOver = lngMouseOver
		
		' get the address for the subcategory to draw
		hAddress = m_oCurrentNode.pBranch + (lngMouseOver * lngNodeSize)
		'UPGRADE_WARNING: Couldn't resolve default property of object oTempNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oTempNode, hAddress, lngNodeSize)
		' update the selected state of the node in the heap
		oTempNode.bSelectedState = NODE_SELECTED
		'UPGRADE_WARNING: Couldn't resolve default property of object oTempNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(hAddress, oTempNode, lngNodeSize)
		
		' set our mem pointer back to start of the selected branch and draw it
		Call DrawItems(m_oSubPic, m_oCurrentNode.pBranch)
		
		' deselect the item for next time
		'UPGRADE_WARNING: Couldn't resolve default property of object oTempNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oTempNode, hAddress, lngNodeSize)
		oTempNode.bSelectedState = NODE_DESELECTED
		'UPGRADE_WARNING: Couldn't resolve default property of object oTempNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(hAddress, oTempNode, lngNodeSize)
	End Sub
End Class