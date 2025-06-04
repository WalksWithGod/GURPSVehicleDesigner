VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Property Variables
Private m_lCompressedSize As Long
Private m_lOriginalSize As Long

'Declares
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

' these properties can be used to retreive the original and compressed size information
' immediately aftre a compression or decompression method has been called
Public Property Get CompressedSize() As Long
    CompressedSize = m_lCompressedSize
End Property

Public Property Get OriginalSize() As Long
    OriginalSize = m_lOriginalSize
End Property


Public Function CompressByteArray(ByRef byteArray() As Byte, ByVal bUseSizePrefix As Boolean) As Long
    'this compresses the chunk and adds the  Original size to the beginning of the compressed buffer
    'this way, you can uncompress the data by calling UncompressB without having to pass in the
    'original length because its already included
    Dim lResult As Long
    Dim lSize As Long
    Dim lOffset As Long
        
    lSize = UBound(byteArray) + 1
    
    'Allocate memory for byte array
    Dim lBufferSize As Long
    Dim TempBuffer() As Byte
    
    lBufferSize = lSize + (lSize * 0.01) + 12
    ReDim TempBuffer(lBufferSize)
    
    If bUseSizePrefix Then
        lOffset = 4
        'pre-pend the original uncompressed length so that the resultant compressed data is self contained
        'and can be decompressed without any other information required
        Call CopyMemory(TempBuffer(0), lSize, 4)
    End If
    
    'do the actual compression using the zlib.dll compress() function
    lResult = compress(TempBuffer(lOffset), lBufferSize, byteArray(0), lSize)
    
    'resize our return buffer to store the compressed results and then copy it in
    ReDim Preserve byteArray(lOffset + lBufferSize - 1)
    CopyMemory byteArray(0), TempBuffer(0), lBufferSize
    Erase TempBuffer
    
    'Set properties if no error occurred
    If lResult = 0 Then m_lCompressedSize = UBound(byteArray) + 1
    
    'Return error code (if any)
    CompressByteArray = lResult
End Function

Public Function DecompressByteArray(ByRef byteArray() As Byte, ByVal lOriginalSize As Long, ByVal bUseSizePrefix As Boolean) As Long
    'This function decompresses a byte array that was compressed with CompressByteArray
    'NOTE: If bUseSizePrefix was used to create the compressed data, then it MUST also be used here!
    'if bUseSizePrefix is set, lOriginalSize argument will be ignored
    Dim lResult As Long
    Dim lSize As Long
    Dim lOffset As Long
    Dim lBufferSize As Long
    Dim TempBuffer() As Byte
              
    If bUseSizePrefix Then
        ' ignore any value passed in lOriginalSize and use the pre-pended length info in the array
        'extract the first long value from TheData which is our OrigSize
        Call CopyMemory(lSize, byteArray(0), 4)
        lOffset = 4
    Else
        lSize = lOriginalSize
    End If
    
     'Allocate memory for buffers
    lBufferSize = lSize + (lSize * 0.01) + 12
    ReDim TempBuffer(lBufferSize)
    
    'do the actual decompression using the zlib.dll uncompress function
    lResult = uncompress(TempBuffer(0), lBufferSize, byteArray(lOffset), UBound(byteArray) - lOffset - 1)
    
    'resize our byteArray to contain the decompressed results
    ReDim Preserve byteArray(lSize - 1)
    CopyMemory byteArray(0), TempBuffer(0), lSize
    
    'Reset properties
    If lResult = 0 Then
        m_lCompressedSize = 0
        lSize = 0
    End If
    
    'Return error code (if any)
    DecompressByteArray = lResult
End Function


Public Function CompressString(sData As String) As Long
    Dim lSize As Long
    Dim sTemp As String
    Dim lResult As Long
    
    'Allocate string space for the buffers
    m_lOriginalSize = Len(sData)
    sTemp = String(m_lOriginalSize + (m_lOriginalSize * 0.01) + 12, 0)
    lSize = Len(sTemp)
    
    'Compress string (temporary string buffer) data
    lResult = compress(ByVal sTemp, lSize, ByVal sData, Len(sData))
    
    'Set original value
    m_lOriginalSize = Len(sData)
    
    'Crop the string and set it to the actual string.
    sData = Left$(sTemp, lSize)
    
    'Set compressed size of string.
    m_lCompressedSize = lSize
    sTemp = ""
    
    'Return error code (if any)
    CompressString = lResult

End Function

Public Function DecompressString(sData As String, lOriginalSize As Long) As Long
    Dim lSize As Long
    Dim sTemp As String
    Dim lResult As Long
    
    'Allocate string space
    sTemp = String(lOriginalSize + (lOriginalSize * 0.01) + 12, 0)
    lSize = Len(sTemp)
    
    'Decompress
    lResult = uncompress(ByVal sTemp, lSize, ByVal sData, Len(sData))
    
    'Make string the size of the uncompressed string
    sData = Left$(sTemp, lSize)
    
    If lResult = 0 Then
        m_lCompressedSize = 0
        m_lOriginalSize = 0
    End If
    
    'Return error code (if any)
    DecompressString = lResult
End Function



