Option Strict Off
Option Explicit On
Module modMatrix
	
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Integer, ByVal dwFlags As Integer, ByRef lpMem As Any) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Integer, ByVal dwFlags As Integer, ByRef lpMem As Any, ByVal dwBytes As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Integer, ByVal dwFlags As Integer, ByRef lpMem As Any) As Integer
	Public Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Integer) As Integer
	Public Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Integer, ByVal dwInitialSize As Integer, ByVal dwMaximumSize As Integer) As Integer
	Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Integer, ByVal dwFlags As Integer, ByVal dwBytes As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef Destination As Any, ByRef Source As Any, ByVal Length As Integer)
	
	Public Const HEAP_NO_SERIALIZE As Short = &H1s
	Public Const HEAP_GROWABLE As Short = &H2s
	Public Const HEAP_GENERATE_EXCEPTIONS As Short = &H4s
	Public Const HEAP_ZERO_MEMORY As Short = &H8s
	Public Const HEAP_REALLOC_IN_PLACE_ONLY As Short = &H10s '  0x00000010
	Public Const HEAP_TAIL_CHECKING_ENABLED As Short = &H20s '  0x00000020
	Public Const HEAP_FREE_CHECKING_ENABLED As Short = &H40s '  0x00000040
	Public Const HEAP_DISABLE_COALESCE_ON_FREE As Short = &H80s '  0x00000080
	Public Const HEAP_CREATE_ALIGN_16 As Short = &H100s '  0x00010000
	Public Const HEAP_CREATE_ENABLE_TRACING As Short = &H200s '  0x00020000
	
	Public Const PAGE_SIZE As Short = 4096 'only on Alphas, the page size is 8196
	Public Const HEAP_SIZE As Short = 8192
	Public Const MAX_HEAP_SIZE As Short = 16384
	
	Public Const ELEMENT_SIZE As Short = 4
	
	Public Structure TABLE '<-- this is really our array descriptor
		Dim hHeap As Integer
		Dim hAddress As Integer
		Dim lngNumRows As Integer
		Dim lngNumColumns As Integer
		'lngElementSize As Long ' note: i removed this.  I am forcing element size to be SINGLE.
	End Structure
	
	' Keep it simple for now.  One heap for each matrix rather than trying to share heaps.
	' Don't want to be bothered with tracking offsets
	
	Public Function createTable(ByVal lngNumRows As Integer, ByVal lngNumColumn As Integer) As Integer
		' creates a heap based array that is lngNumRows by lngNumColumns with each array item of width elementSize
		' if tbl.hHeap <> 0, it tries to create the new table on the same heap at the end of the
		Dim lngMemoryNeeded As Integer
		Dim lngRet As Integer
		Dim tbl As TABLE
		
		With tbl
			If (lngNumRows > 0) And (lngNumColumn > 0) And (ELEMENT_SIZE > 0) Then
				'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				lngMemoryNeeded = LenB(tbl) + (lngNumRows * lngNumColumn * ELEMENT_SIZE)
				.hHeap = HeapCreate(HEAP_NO_SERIALIZE, lngMemoryNeeded, MAX_HEAP_SIZE)
				.hAddress = HeapAlloc(.hHeap, HEAP_NO_SERIALIZE, lngMemoryNeeded)
				.lngNumColumns = lngNumColumn
				.lngNumRows = lngNumRows
				' copy the descriptor to the start of the address space
				'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object tbl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CopyMemory(.hAddress, tbl, LenB(tbl))
				
				' todo: need to pass in a byReference lPtr so that i can return success true/false
				'       rather than passing out the resulting pointer and just relying on checking
				'       that lngRet is <> 0. (Note i cant even check for >0 cuz some times the address
				'       in a long variable will be negative (wouldnt happen with a dword)
				lngRet = .hAddress
			End If
		End With
		createTable = lngRet
	End Function
	
	Public Function destroyTable(ByVal ptrTable As Integer) As Boolean
		Dim tbl As TABLE
		Dim lngRet As Integer
		If ptrTable <> 0 Then
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object tbl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(tbl, ptrTable, LenB(tbl))
			With tbl
				lngRet = HeapFree(.hHeap, HEAP_NO_SERIALIZE, .hAddress)
				lngRet = HeapDestroy(.hHeap)
			End With
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object tbl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(tbl, 0, LenB(tbl))
		End If
#If DEBUG_MODE Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		If lngRet Then
		Debug.Print "modMatrix:destroyTable -- SUCCESS.  Heap based table destroyed"
		Else
		Debug.Print "modMatrix:destroyTable() -- Failed to destroy heap based table."
		End If
#End If
		
		destroyTable = lngRet
	End Function
	
	Public Function getTableItemValue(ByVal ptrTable As Integer, ByVal lngRow As Integer, ByVal lngColumn As Integer, ByRef sngResult As Single) As Boolean
		Dim tbl As TABLE
		Dim hValueOffset As Integer
		Dim hStart As Integer
		'       We use Row Major order just like we're used to in VB
		'e.g.   getMatrixValue (0,0) = first row, first column
		'       getMatrixValue (0,3) = first row, 3rd column.
		
		If ptrTable <> 0 Then
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object tbl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(tbl, ptrTable, LenB(tbl))
			'todo: verify row/column doesnt exceed bounds
			With tbl
				'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				hStart = .hAddress + LenB(tbl)
				hValueOffset = hStart + (ELEMENT_SIZE * .lngNumRows * lngColumn) + (ELEMENT_SIZE * lngRow)
			End With
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			System.Diagnostics.Debug.Assert(LenB(sngResult) = ELEMENT_SIZE, "")
			CopyMemory(sngResult, hValueOffset, ELEMENT_SIZE)
			getTableItemValue = True
		End If
		
	End Function
	
	Public Function setTableItemValue(ByVal ptrTable As Integer, ByVal lngRow As Integer, ByVal lngColumn As Integer, ByVal sngValue As Single) As Boolean
		setTableItemValue = True
	End Function
End Module