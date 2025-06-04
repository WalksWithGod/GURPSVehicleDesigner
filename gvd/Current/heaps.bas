%PAGE_SIZE = 4096   'except on Alphas where the page size is 8196
%MAX_HEAP_SIZE = 16384
TYPE NODE
    dwNext AS NODE PTR
    pData AS DWORD 'pointer to our entry
    length AS LONG ' length of the entry
END TYPE

TYPE LINKED_LIST
    hHeap AS DWORD
    Address AS DWORD     ' This is needed since its actually the TOP of the heap!!
    ListSize AS DWORD
    ListMaxSize AS DWORD
    pNode AS NODE PTR    ' pointer to Head node Address
END TYPE

' This is part of the ZIP_ARCHIVE structure
TYPE FILES_TO_PROCESS
    hHeap AS DWORD       ' object handle for Heap
    Address AS DWORD     ' pointer to start of File list string
    HeapSize AS DWORD    ' total allocated size of heap in bytes
    MaxHeapSize AS DWORD   ' max reserved heap size
    Length AS LONG       ' length of the string data stored in the heap
    Count AS LONG         ' number of elements
    List AS LINKED_LIST  ' linked_list heap (and first element in that heap)
END TYPE


DECLARE FUNCTION TerminateHeap(BYVAL hHeap AS DWORD, BYVAL Address AS DWORD)AS LONG
DECLARE FUNCTION ExpandHeap(BYREF hAddress AS DWORD, BYREF hCurrentHeap AS DWORD, BYVAL lngNeeded AS LONG)AS LONG
DECLARE FUNCTION AppendToHeap(pFP AS FILES_TO_PROCESS PTR, sFiles AS STRING) AS LONG
DECLARE FUNCTION CheckHeapSize(hCurrentHeap AS DWORD, hCurrentAddress AS DWORD, dwHeapSize AS DWORD, lngExistingLength AS LONG, lngNewLength AS LONG, hOldHeap AS DWORD, hOldAddress AS DWORD) AS LONG
DECLARE FUNCTION Node_Insert_Front ( pHeadNode AS NODE PTR, pNew AS NODE PTR, BYVAL pData AS DWORD, BYVAL lngBufferLen AS LONG) AS LONG
DECLARE FUNCTION UpdateLinkedList(pLL AS LINKED_LIST PTR,BYVAL hBuffer AS DWORD, BYVAL lngLength AS LONG) AS DWORD


FUNCTION TerminateHeap(BYVAL hHeap AS DWORD, BYVAL Address AS DWORD) AS LONG
   DIM lngRet AS LONG

   ' free the memory
   lngRet = HeapFree(hHeap, %HEAP_NO_SERIALIZE, Address)
   IF lngRet <> %TRUE THEN
      ' there are no guarantees the heap will be freed on demand
      ' an error here should be expected
      CPrint "Code = " & STR$(lngRet) & "TODO: Cant free heap!"
   END IF

   ' destroy the object
   lngRet = HeapDestroy(hHeap)

   IF lngRet <> %TRUE THEN
      CPrint "Code = " & STR$(lngRet) & " TODO!  Cant destroy heap!"
      ' in this case, i guess we return TRUE anyways and
      ' allow windows to destroy the heap when the process ends.
      ' Nonetheless, this Function must always return %TRUE
   END IF

   FUNCTION = %TRUE
END FUNCTION




FUNCTION CheckHeapSize(hCurrentHeap AS DWORD, hCurrentAddress AS DWORD, dwHeapSize AS DWORD, lngExistingLength AS LONG, lngNewLength AS LONG, hOldHeap AS DWORD, hOldAddress AS DWORD) AS LONG
    'verifies that the amount of data we want to add
    ' will fit into our existing heap.  Returns FALSE if
    ' the heap is too small and needs to be expanded
    ' via ExpandHeap()
    LOCAL lngNeeded AS LONG

    ' Compute total size needed
    lngNeeded = lngExistingLength + lngNewLength

    ' make sure our Heap is ok and that we have enuf space
    IF lngNeeded > dwHeapSize THEN
        hOldAddress = hCurrentAddress
        hOldHeap = hCurrentHeap

        ExpandHeap hCurrentAddress, hCurrentHeap,lngNeeded
        IF hOldAddress = hCurrentAddress THEN
            hOldAddress = %NULL
        END IF
    END IF

    ' update struct
    dwHeapSize = HeapSize(hCurrentHeap, %HEAP_NO_SERIALIZE, hCurrentAddress)

    FUNCTION = %TRUE

END FUNCTION

FUNCTION AppendToHeap(pFP AS FILES_TO_PROCESS PTR, sFiles AS STRING) AS LONG
    DIM pByte AS BYTE PTR
    DIM lngRet AS LONG
    DIM lngNewElements AS LONG
    DIM pZ AS ASCIIZ PTR
    DIM pNode AS NODE PTR
    DIM pNodeOffset AS NODE PTR
    DIM i AS LONG
    DIM hOldAddress AS DWORD
    DIM hOldHeap AS DWORD
    DIM lngNewLength AS LONG
    DIM lngExistingLength AS LONG

    lngNewElements = PARSECOUNT(sFiles, CHR$(0))

    ' make sure we have enuf room IN our linked list
    lngNewLength = lngNewElements * SIZEOF(@pNODE)
    lngExistingLength = @pFP.Count * SIZEOF(@pNode)
    lngRet = CheckHeapSize(@pFP.List.hHeap, @pFP.List.Address, @pFP.List.ListSize, lngExistingLength, lngNewLength, hOldHeap, hOldAddress)

    ' find offset to begin copying new data to our filebuffer heap
    pz = @pFP.Address + @pFP.Length

    ' find offset to begin copying new data to our Linked List heap
    pByte = @pFP.List.Address
    pNode = pByte + ( @pFP.Count * SIZEOF(@pNode))

    ' append the buffer and add the new linked list node
    FOR i = 1 TO lngNewElements
        @pZ = PARSE$(sFiles, CHR$(0), i)
        ' update our count
        @pFP.Count = @pFP.Count + 1

        ' create the linked list item
        lngRet = Node_Insert_Front (BYVAL @pFP.List.pNode, BYVAL pNode, BYVAL pz,  LEN(@pZ)+1 )
        IF NOT ISTRUE lngRet THEN
           CPrint "Error adding linked list node."
        END IF

        ' move our pointers to next append position
        pZ = pZ + LEN(@pZ)+1 ' add one for the $NULL terminator
        INCR pNode
    NEXT

    ' update the new length of our data in this heap
    @pFP.Length = @pFP.Length + LEN(sFiles)

    FUNCTION = %TRUE
END FUNCTION


FUNCTION Node_Insert_Front ( pHeadNode AS NODE PTR, pNew AS NODE PTR, BYVAL pData AS DWORD, BYVAL lngBufferLen AS LONG) AS LONG

    @pNew.dwNext = pHeadNode
    @pNew.pData = pData
    pHeadNode = pNew

    FUNCTION = %TRUE

END FUNCTION



FUNCTION ExpandHeap(BYREF hAddress AS DWORD, BYREF hCurrentHeap AS DWORD, BYVAL lngNeeded AS LONG) AS LONG
        DIM hMem AS DWORD
        DIM lngRet AS LONG
        DIM hOldHeap AS DWORD

        hOldHeap = hCurrentHeap
        hMem = hAddress ' preserve another copy since this might get modified

        ' try expanding the memory block of the existing heap
        IF hAddress = %NULL THEN
            hAddress = HeapAlloc(hCurrentHeap, %HEAP_NO_SERIALIZE, lngNeeded)
        ELSE
            hAddress = HeapReAlloc(hCurrentHeap, %HEAP_NO_SERIALIZE OR %HEAP_REALLOC_IN_PLACE_ONLY, BYVAL(hMem),lngNeeded)
            IF hAddress <> %NULL THEN FUNCTION = %TRUE
        END IF

        ' if that fails, try creating a new heap and allocating a new memory block
        IF hAddress = %NULL THEN
            ' attempt to create a new heap that is at least as big as lngNeeded
            hCurrentHeap = HeapCreate(%HEAP_NO_SERIALIZE, lngNeeded, %MAX_HEAP_SIZE)
            IF hCurrentHeap <> %NULL THEN
                hAddress = HeapAlloc(hCurrentHeap, %HEAP_NO_SERIALIZE, lngNeeded)
                IF hAddress = %NULL THEN
                    lngRet = TerminateHeap(hOldHeap, hMem)
                    CPrint "Failed to allocated to new heap"
                    EXIT FUNCTION
                END IF
                FUNCTION = %TRUE
            ELSE
                CPrint "Failed to create heap"
                EXIT FUNCTION
            END IF
        END IF

END FUNCTION





                                            