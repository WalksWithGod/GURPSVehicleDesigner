Option Strict Off
Option Explicit On
Module modHeaps
	' MPJ - This module is converted from some PowerBASIC code i was working on...hence some of the odd variable
	' names which seem to indicate WORD and DOUBLE WORD variable type.
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Integer, ByVal dwFlags As Integer, ByRef lpMem As Any) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Integer, ByVal dwFlags As Integer, ByRef lpMem As Any, ByVal dwBytes As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Integer, ByVal dwFlags As Integer, ByRef lpMem As Any) As Integer
	Public Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Integer) As Integer
	Public Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Integer, ByVal dwInitialSize As Integer, ByVal dwMaximumSize As Integer) As Integer
	Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Integer, ByVal dwFlags As Integer, ByVal dwBytes As Integer) As Integer
	
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
	
	
	Public Function TerminateHeap(ByVal hHeap As Integer, ByVal Address As Integer) As Integer
		Dim lngRet As Integer
		' free the memory
		lngRet = HeapFree(hHeap, HEAP_NO_SERIALIZE, Address)
		If lngRet <> True Then
			' there are no guarantees the heap will be freed on demand
			' an error here should be expected
			Debug.Print("Code = " & Str(lngRet) & "TODO: Cant free heap!")
		End If
		
		' destroy the object
		lngRet = HeapDestroy(hHeap)
		If lngRet <> True Then
			Debug.Print("Code = " & Str(lngRet) & " TODO!  Cant destroy heap!")
			' in this case, i guess we return TRUE anyways and
			' allow windows to destroy the heap when the process ends.
			' Nonetheless, this Function must always return %TRUE
		End If
		TerminateHeap = True
	End Function
	
	Public Function ExpandHeap(ByRef hAddress As Integer, ByRef hCurrentHeap As Integer, ByVal lngNeeded As Integer) As Integer
		Dim hMem As Integer
		Dim lngRet As Integer
		
		hMem = hAddress
		
		If hAddress = 0 Then
			' no block allocated in the heap, we can jsut allocate a new one
			hAddress = HeapAlloc(hCurrentHeap, HEAP_NO_SERIALIZE, lngNeeded)
			' if this fails here, hAddress = 0 and that is our retval
		Else
			' try expanding the memory block of the existing heap
			hAddress = HeapReAlloc(hCurrentHeap, HEAP_NO_SERIALIZE Or HEAP_REALLOC_IN_PLACE_ONLY, hMem, lngNeeded)
			' if we are unable to expand heap, we should not attempt to create a new one, we just
			' let the caller call functions to terminate and then create new
		End If
		
		' returns new address, will be 0 if fails
		ExpandHeap = hAddress
	End Function
	
	Function CheckHeapSize(ByRef hCurrentHeap As Integer, ByRef hCurrentAddress As Integer, ByRef lngHeapSize As Integer, ByRef lngExistingLength As Integer, ByRef lngNewLength As Integer, ByRef hOldHeap As Integer, ByRef hOldAddress As Integer) As Integer
		'verifies that the amount of data we want to add
		' will fit into our existing heap.  Returns FALSE if
		' the heap is too small and needs to be expanded
		' via ExpandHeap()
		Dim lngNeeded As Integer
		
		' Compute total size needed
		lngNeeded = lngExistingLength + lngNewLength
		' make sure our Heap is ok and that we have enuf space
		If lngNeeded > lngHeapSize Then
			hOldAddress = hCurrentAddress
			hOldHeap = hCurrentHeap
			
			ExpandHeap(hCurrentAddress, hCurrentHeap, lngNeeded)
			If hOldAddress = hCurrentAddress Then
				hOldAddress = 0
			End If
		End If
		
		' update the size
		lngHeapSize = HeapSize(hCurrentHeap, HEAP_NO_SERIALIZE, hCurrentAddress)
		CheckHeapSize = True
	End Function
End Module