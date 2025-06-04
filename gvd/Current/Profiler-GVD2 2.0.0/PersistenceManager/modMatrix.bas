Attribute VB_Name = "modMatrix"
Option Explicit


Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Public Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Public Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

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

Public Const ELEMENT_SIZE = 4

Public Type TABLE  '<-- this is really our array descriptor
    hHeap As Long
    hAddress As Long
    lngNumRows As Long
    lngNumColumns As Long
     'lngElementSize As Long ' note: i removed this.  I am forcing element size to be SINGLE.
End Type

' Keep it simple for now.  One heap for each matrix rather than trying to share heaps.
' Don't want to be bothered with tracking offsets

Public Function createTable(ByVal lngNumRows As Long, ByVal lngNumColumn As Long) As Long
    ' creates a heap based array that is lngNumRows by lngNumColumns with each array item of width elementSize
    ' if tbl.hHeap <> 0, it tries to create the new table on the same heap at the end of the
vbwProfiler.vbwProcIn 24
    Dim lngMemoryNeeded As Long
    Dim lngRet As Long
    Dim tbl As TABLE

vbwProfiler.vbwExecuteLine 343
    With tbl
vbwProfiler.vbwExecuteLine 344
        If (lngNumRows > 0) And (lngNumColumn > 0) And (ELEMENT_SIZE > 0) Then
vbwProfiler.vbwExecuteLine 345
            lngMemoryNeeded = LenB(tbl) + (lngNumRows * lngNumColumn * ELEMENT_SIZE)
vbwProfiler.vbwExecuteLine 346
            .hHeap = HeapCreate(HEAP_NO_SERIALIZE, lngMemoryNeeded, MAX_HEAP_SIZE)
vbwProfiler.vbwExecuteLine 347
            .hAddress = HeapAlloc(.hHeap, HEAP_NO_SERIALIZE, lngMemoryNeeded)
vbwProfiler.vbwExecuteLine 348
            .lngNumColumns = lngNumColumn
vbwProfiler.vbwExecuteLine 349
            .lngNumRows = lngNumRows
            ' copy the descriptor to the start of the address space
vbwProfiler.vbwExecuteLine 350
            CopyMemory ByVal (.hAddress), tbl, LenB(tbl)

            ' todo: need to pass in a byReference lPtr so that i can return success true/false
            '       rather than passing out the resulting pointer and just relying on checking
            '       that lngRet is <> 0. (Note i cant even check for >0 cuz some times the address
            '       in a long variable will be negative (wouldnt happen with a dword)
vbwProfiler.vbwExecuteLine 351
            lngRet = .hAddress
        End If
vbwProfiler.vbwExecuteLine 352 'B
vbwProfiler.vbwExecuteLine 353
    End With
vbwProfiler.vbwExecuteLine 354
    createTable = lngRet
vbwProfiler.vbwProcOut 24
vbwProfiler.vbwExecuteLine 355
End Function

Public Function destroyTable(ByVal ptrTable As Long) As Boolean
vbwProfiler.vbwProcIn 25
    Dim tbl As TABLE
    Dim lngRet As Long
vbwProfiler.vbwExecuteLine 356
    If ptrTable <> 0 Then
vbwProfiler.vbwExecuteLine 357
        CopyMemory tbl, ByVal (ptrTable), LenB(tbl)
vbwProfiler.vbwExecuteLine 358
        With tbl
vbwProfiler.vbwExecuteLine 359
            lngRet = HeapFree(.hHeap, HEAP_NO_SERIALIZE, .hAddress)
vbwProfiler.vbwExecuteLine 360
            lngRet = HeapDestroy(.hHeap)
vbwProfiler.vbwExecuteLine 361
        End With
vbwProfiler.vbwExecuteLine 362
        CopyMemory tbl, 0&, LenB(tbl)
    End If
vbwProfiler.vbwExecuteLine 363 'B
    #If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 364
        If lngRet Then
vbwProfiler.vbwExecuteLine 365
            Debug.Print "modMatrix:destroyTable -- SUCCESS.  Heap based table destroyed"
        Else
vbwProfiler.vbwExecuteLine 366 'B
vbwProfiler.vbwExecuteLine 367
            Debug.Print "modMatrix:destroyTable() -- Failed to destroy heap based table."
        End If
vbwProfiler.vbwExecuteLine 368 'B
    #End If

vbwProfiler.vbwExecuteLine 369
    destroyTable = lngRet
vbwProfiler.vbwProcOut 25
vbwProfiler.vbwExecuteLine 370
End Function

Public Function getTableItemValue(ByVal ptrTable As Long, ByVal lngRow As Long, ByVal lngColumn As Long, ByRef sngResult As Single) As Boolean
vbwProfiler.vbwProcIn 26
    Dim tbl As TABLE
    Dim hValueOffset As Long
    Dim hStart As Long
'       We use Row Major order just like we're used to in VB
'e.g.   getMatrixValue (0,0) = first row, first column
'       getMatrixValue (0,3) = first row, 3rd column.

vbwProfiler.vbwExecuteLine 371
    If ptrTable <> 0 Then
vbwProfiler.vbwExecuteLine 372
        CopyMemory tbl, ByVal (ptrTable), LenB(tbl)
        'todo: verify row/column doesnt exceed bounds
vbwProfiler.vbwExecuteLine 373
        With tbl
vbwProfiler.vbwExecuteLine 374
            hStart = .hAddress + LenB(tbl)
vbwProfiler.vbwExecuteLine 375
            hValueOffset = hStart + (ELEMENT_SIZE * .lngNumRows * lngColumn) + (ELEMENT_SIZE * lngRow)
vbwProfiler.vbwExecuteLine 376
        End With
vbwProfiler.vbwExecuteLine 377
        Debug.Assert LenB(sngResult) = ELEMENT_SIZE
vbwProfiler.vbwExecuteLine 378
        CopyMemory sngResult, ByVal (hValueOffset), ELEMENT_SIZE
vbwProfiler.vbwExecuteLine 379
        getTableItemValue = True
    End If
vbwProfiler.vbwExecuteLine 380 'B

vbwProfiler.vbwProcOut 26
vbwProfiler.vbwExecuteLine 381
End Function

Public Function setTableItemValue(ByVal ptrTable As Long, ByVal lngRow As Long, ByVal lngColumn As Long, ByVal sngValue As Single) As Boolean
vbwProfiler.vbwProcIn 27
vbwProfiler.vbwExecuteLine 382
    setTableItemValue = True
vbwProfiler.vbwProcOut 27
vbwProfiler.vbwExecuteLine 383
End Function


