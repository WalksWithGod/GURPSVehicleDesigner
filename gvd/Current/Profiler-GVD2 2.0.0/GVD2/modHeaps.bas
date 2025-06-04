Attribute VB_Name = "modHeaps"
Option Explicit
' MPJ - This module is converted from some PowerBASIC code i was working on...hence some of the odd variable
' names which seem to indicate WORD and DOUBLE WORD variable type.

Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Public Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Public Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long

Public Const HEAP_NO_SERIALIZE = &H1
Public Const HEAP_GROWABLE = &H2
Public Const HEAP_GENERATE_EXCEPTIONS = &H4
Public Const HEAP_ZERO_MEMORY = &H8
Public Const HEAP_REALLOC_IN_PLACE_ONLY = &H10                  '  0x00000010
Public Const HEAP_TAIL_CHECKING_ENABLED = &H20                  '  0x00000020
Public Const HEAP_FREE_CHECKING_ENABLED = &H40                   '  0x00000040
Public Const HEAP_DISABLE_COALESCE_ON_FREE = &H80               '  0x00000080
Public Const HEAP_CREATE_ALIGN_16 = &H100                       '  0x00010000
Public Const HEAP_CREATE_ENABLE_TRACING = &H200                  '  0x00020000

Public Const PAGE_SIZE = 4096   'only on Alphas, the page size is 8196
Public Const HEAP_SIZE = 8192
Public Const MAX_HEAP_SIZE = 16384


Public Function TerminateHeap(ByVal hHeap As Long, ByVal Address As Long) As Long
vbwProfiler.vbwProcIn 233
   Dim lngRet As Long
   ' free the memory
vbwProfiler.vbwExecuteLine 4552
   lngRet = HeapFree(hHeap, HEAP_NO_SERIALIZE, Address)
vbwProfiler.vbwExecuteLine 4553
   If lngRet <> True Then
      ' there are no guarantees the heap will be freed on demand
      ' an error here should be expected
vbwProfiler.vbwExecuteLine 4554
      Debug.Print "Code = " & Str$(lngRet) & "TODO: Cant free heap!"
   End If
vbwProfiler.vbwExecuteLine 4555 'B

   ' destroy the object
vbwProfiler.vbwExecuteLine 4556
   lngRet = HeapDestroy(hHeap)
vbwProfiler.vbwExecuteLine 4557
   If lngRet <> True Then
vbwProfiler.vbwExecuteLine 4558
      Debug.Print "Code = " & Str$(lngRet) & " TODO!  Cant destroy heap!"
      ' in this case, i guess we return TRUE anyways and
      ' allow windows to destroy the heap when the process ends.
      ' Nonetheless, this Function must always return %TRUE
   End If
vbwProfiler.vbwExecuteLine 4559 'B
vbwProfiler.vbwExecuteLine 4560
   TerminateHeap = True
vbwProfiler.vbwProcOut 233
vbwProfiler.vbwExecuteLine 4561
End Function

Public Function ExpandHeap(ByRef hAddress As Long, ByRef hCurrentHeap As Long, ByVal lngNeeded As Long) As Long
vbwProfiler.vbwProcIn 234
        Dim hMem As Long
        Dim lngRet As Long

vbwProfiler.vbwExecuteLine 4562
        hMem = hAddress

vbwProfiler.vbwExecuteLine 4563
        If hAddress = 0 Then
            ' no block allocated in the heap, we can jsut allocate a new one
vbwProfiler.vbwExecuteLine 4564
            hAddress = HeapAlloc(hCurrentHeap, HEAP_NO_SERIALIZE, lngNeeded)
            ' if this fails here, hAddress = 0 and that is our retval
        Else
vbwProfiler.vbwExecuteLine 4565 'B
            ' try expanding the memory block of the existing heap
vbwProfiler.vbwExecuteLine 4566
            hAddress = HeapReAlloc(hCurrentHeap, HEAP_NO_SERIALIZE Or HEAP_REALLOC_IN_PLACE_ONLY, ByVal (hMem), lngNeeded)
            ' if we are unable to expand heap, we should not attempt to create a new one, we just
            ' let the caller call functions to terminate and then create new
        End If
vbwProfiler.vbwExecuteLine 4567 'B

        ' returns new address, will be 0 if fails
vbwProfiler.vbwExecuteLine 4568
        ExpandHeap = hAddress
vbwProfiler.vbwProcOut 234
vbwProfiler.vbwExecuteLine 4569
End Function

Function CheckHeapSize(hCurrentHeap As Long, hCurrentAddress As Long, lngHeapSize As Long, lngExistingLength As Long, lngNewLength As Long, hOldHeap As Long, hOldAddress As Long) As Long
    'verifies that the amount of data we want to add
    ' will fit into our existing heap.  Returns FALSE if
    ' the heap is too small and needs to be expanded
    ' via ExpandHeap()
vbwProfiler.vbwProcIn 235
    Dim lngNeeded As Long

    ' Compute total size needed
vbwProfiler.vbwExecuteLine 4570
    lngNeeded = lngExistingLength + lngNewLength
    ' make sure our Heap is ok and that we have enuf space
vbwProfiler.vbwExecuteLine 4571
    If lngNeeded > lngHeapSize Then
vbwProfiler.vbwExecuteLine 4572
        hOldAddress = hCurrentAddress
vbwProfiler.vbwExecuteLine 4573
        hOldHeap = hCurrentHeap

vbwProfiler.vbwExecuteLine 4574
        ExpandHeap hCurrentAddress, hCurrentHeap, lngNeeded
vbwProfiler.vbwExecuteLine 4575
        If hOldAddress = hCurrentAddress Then
vbwProfiler.vbwExecuteLine 4576
            hOldAddress = 0
        End If
vbwProfiler.vbwExecuteLine 4577 'B
    End If
vbwProfiler.vbwExecuteLine 4578 'B

    ' update the size
vbwProfiler.vbwExecuteLine 4579
    lngHeapSize = HeapSize(hCurrentHeap, HEAP_NO_SERIALIZE, hCurrentAddress)
vbwProfiler.vbwExecuteLine 4580
    CheckHeapSize = True
vbwProfiler.vbwProcOut 235
vbwProfiler.vbwExecuteLine 4581
End Function

