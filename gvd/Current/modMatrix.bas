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
    Dim lngMemoryNeeded As Long
    Dim lngRet As Long
    Dim tbl As TABLE
        
    With tbl
        If (lngNumRows > 0) And (lngNumColumn > 0) And (ELEMENT_SIZE > 0) Then
            lngMemoryNeeded = LenB(tbl) + (lngNumRows * lngNumColumn * ELEMENT_SIZE)
            .hHeap = HeapCreate(HEAP_NO_SERIALIZE, lngMemoryNeeded, MAX_HEAP_SIZE)
            .hAddress = HeapAlloc(.hHeap, HEAP_NO_SERIALIZE, lngMemoryNeeded)
            .lngNumColumns = lngNumColumn
            .lngNumRows = lngNumRows
            ' copy the descriptor to the start of the address space
            CopyMemory ByVal (.hAddress), tbl, LenB(tbl)
            
            ' todo: need to pass in a byReference lPtr so that i can return success true/false
            '       rather than passing out the resulting pointer and just relying on checking
            '       that lngRet is <> 0. (Note i cant even check for >0 cuz some times the address
            '       in a long variable will be negative (wouldnt happen with a dword)
            lngRet = .hAddress
        End If
    End With
    createTable = lngRet
End Function

Public Function destroyTable(ByVal ptrTable As Long) As Boolean
    Dim tbl As TABLE
    Dim lngRet As Long
    If ptrTable <> 0 Then
        CopyMemory tbl, ByVal (ptrTable), LenB(tbl)
        With tbl
            lngRet = HeapFree(.hHeap, HEAP_NO_SERIALIZE, .hAddress)
            lngRet = HeapDestroy(.hHeap)
        End With
        CopyMemory tbl, 0&, LenB(tbl)
    End If
    #If DEBUG_MODE Then
        If lngRet Then
            Debug.Print "modMatrix:destroyTable -- SUCCESS.  Heap based table destroyed"
        Else
            Debug.Print "modMatrix:destroyTable() -- Failed to destroy heap based table."
        End If
    #End If
    
    destroyTable = lngRet
End Function

Public Function getTableItemValue(ByVal ptrTable As Long, ByVal lngRow As Long, ByVal lngColumn As Long, ByRef sngResult As Single) As Boolean
    Dim tbl As TABLE
    Dim hValueOffset As Long
    Dim hStart As Long
'       We use Row Major order just like we're used to in VB
'e.g.   getMatrixValue (0,0) = first row, first column
'       getMatrixValue (0,3) = first row, 3rd column.
    
    If ptrTable <> 0 Then
        CopyMemory tbl, ByVal (ptrTable), LenB(tbl)
        'todo: verify row/column doesnt exceed bounds
        With tbl
            hStart = .hAddress + LenB(tbl)
            hValueOffset = hStart + (ELEMENT_SIZE * .lngNumRows * lngColumn) + (ELEMENT_SIZE * lngRow)
        End With
        Debug.Assert LenB(sngResult) = ELEMENT_SIZE
        CopyMemory sngResult, ByVal (hValueOffset), ELEMENT_SIZE
        getTableItemValue = True
    End If
   
End Function

Public Function setTableItemValue(ByVal ptrTable As Long, ByVal lngRow As Long, ByVal lngColumn As Long, ByVal sngValue As Single) As Boolean
    setTableItemValue = True
End Function
