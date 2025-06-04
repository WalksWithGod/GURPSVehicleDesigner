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

 'Default Property Values:
Const m_def_CompressedSize = 0
Const m_def_OriginalSize = 0
 'Property Variables:
Dim m_CompressedSize As Long
Dim m_OriginalSize As Long

'Declares
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Enum CZErrors
[Insufficient Buffer] = -5
End Enum


Public Function CompressData(ByRef TheData() As Byte) As Long
vbwProfiler.vbwProcIn 196
    Dim result As Long

vbwProfiler.vbwExecuteLine 4009
    OriginalSize = UBound(TheData) + 1

    'Allocate memory for byte array
    Dim BufferSize As Long
    Dim TempBuffer() As Byte

vbwProfiler.vbwExecuteLine 4010
    BufferSize = UBound(TheData) + 1
vbwProfiler.vbwExecuteLine 4011
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
vbwProfiler.vbwExecuteLine 4012
    ReDim TempBuffer(BufferSize)

    'Compress byte array (data)
vbwProfiler.vbwExecuteLine 4013
    result = compress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1)

    'Truncate to compressed size
vbwProfiler.vbwExecuteLine 4014
    ReDim Preserve TheData(BufferSize - 1)
vbwProfiler.vbwExecuteLine 4015
    CopyMemory TheData(0), TempBuffer(0), BufferSize

    'Cleanup
vbwProfiler.vbwExecuteLine 4016
    Erase TempBuffer

    'Set properties if no error occurred
vbwProfiler.vbwExecuteLine 4017
    If result = 0 Then
vbwProfiler.vbwExecuteLine 4018
         CompressedSize = UBound(TheData) + 1
    End If
vbwProfiler.vbwExecuteLine 4019 'B

    'Return error code (if any)
vbwProfiler.vbwExecuteLine 4020
    CompressData = result

vbwProfiler.vbwProcOut 196
vbwProfiler.vbwExecuteLine 4021
End Function


Public Function DecompressData(ByRef TheData() As Byte, OrigSize As Long) As Long
vbwProfiler.vbwProcIn 197
    Dim result As Long

    'Allocate memory for buffers
    Dim BufferSize As Long
    Dim TempBuffer() As Byte

vbwProfiler.vbwExecuteLine 4022
    BufferSize = OrigSize
vbwProfiler.vbwExecuteLine 4023
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
vbwProfiler.vbwExecuteLine 4024
    ReDim TempBuffer(BufferSize)

    'Decompress data
vbwProfiler.vbwExecuteLine 4025
    result = uncompress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1)

    'Truncate buffer to compressed size
vbwProfiler.vbwExecuteLine 4026
    ReDim Preserve TheData(BufferSize - 1)
vbwProfiler.vbwExecuteLine 4027
    CopyMemory TheData(0), TempBuffer(0), BufferSize

    'Reset properties
vbwProfiler.vbwExecuteLine 4028
    If result = 0 Then
vbwProfiler.vbwExecuteLine 4029
    CompressedSize = 0
vbwProfiler.vbwExecuteLine 4030
    OriginalSize = 0
    End If
vbwProfiler.vbwExecuteLine 4031 'B

    'Return error code (if any)
vbwProfiler.vbwExecuteLine 4032
    DecompressData = result

vbwProfiler.vbwProcOut 197
vbwProfiler.vbwExecuteLine 4033
End Function

Public Function CompressString(TheString As String) As Long
vbwProfiler.vbwProcIn 198

vbwProfiler.vbwExecuteLine 4034
    OriginalSize = Len(TheString)

    'Allocate string space for the buffers
    Dim CmpSize As Long
    Dim TBuff As String
    Dim ret As Long

vbwProfiler.vbwExecuteLine 4035
    TBuff = String(OriginalSize + (OriginalSize * 0.01) + 12, 0)
vbwProfiler.vbwExecuteLine 4036
    CmpSize = Len(TBuff)

    'Compress string (temporary string buffer) data
vbwProfiler.vbwExecuteLine 4037
    ret = compress(ByVal TBuff, CmpSize, ByVal TheString, Len(TheString))

    'Set original value
vbwProfiler.vbwExecuteLine 4038
    OriginalSize = Len(TheString)

    'Crop the string and set it to the actual string.
vbwProfiler.vbwExecuteLine 4039
    TheString = Left$(TBuff, CmpSize)

    'Set compressed size of string.
vbwProfiler.vbwExecuteLine 4040
    CompressedSize = CmpSize

    'Cleanup
vbwProfiler.vbwExecuteLine 4041
    TBuff = ""

    'Return error code (if any)
vbwProfiler.vbwExecuteLine 4042
    CompressString = ret

vbwProfiler.vbwProcOut 198
vbwProfiler.vbwExecuteLine 4043
End Function

Public Function DecompressString(TheString As String, OrigSize As Long) As Long
vbwProfiler.vbwProcIn 199

    'Allocate string space
    Dim CmpSize As Long
    Dim TBuff As String
    Dim result As Long

vbwProfiler.vbwExecuteLine 4044
    TBuff = String(OriginalSize + (OriginalSize * 0.01) + 12, 0)
vbwProfiler.vbwExecuteLine 4045
    CmpSize = Len(TBuff)

    'Decompress
vbwProfiler.vbwExecuteLine 4046
    result = uncompress(ByVal TBuff, CmpSize, ByVal TheString, Len(TheString))

    'Make string the size of the uncompressed string
vbwProfiler.vbwExecuteLine 4047
    TheString = Left$(TBuff, CmpSize)

    'Reset properties
vbwProfiler.vbwExecuteLine 4048
    If result = 0 Then
vbwProfiler.vbwExecuteLine 4049
    CompressedSize = 0
vbwProfiler.vbwExecuteLine 4050
    OriginalSize = 0
    End If
vbwProfiler.vbwExecuteLine 4051 'B

    'Return error code (if any)
vbwProfiler.vbwExecuteLine 4052
    DecompressString = result

vbwProfiler.vbwProcOut 199
vbwProfiler.vbwExecuteLine 4053
End Function


Public Function CompressB(ByRef TheData() As Byte) As Long
    '//this compresses the chunk and adds the  Original size to the beginning of the compressed buffer
vbwProfiler.vbwProcIn 200
    Dim result As Long
    Dim lngSize As Long

vbwProfiler.vbwExecuteLine 4054
    Const LONGLENGTH = 4

vbwProfiler.vbwExecuteLine 4055
    lngSize = UBound(TheData) + 1

    'Allocate memory for byte array
    Dim BufferSize As Long
    Dim TempBuffer() As Byte

vbwProfiler.vbwExecuteLine 4056
    BufferSize = lngSize
vbwProfiler.vbwExecuteLine 4057
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
vbwProfiler.vbwExecuteLine 4058
    ReDim TempBuffer(BufferSize)

    'add the length
vbwProfiler.vbwExecuteLine 4059
    Call CopyMemory(TempBuffer(0), lngSize, LONGLENGTH)

    'Compress byte array (data)
vbwProfiler.vbwExecuteLine 4060
    result = compress(TempBuffer(4), BufferSize, TheData(0), lngSize)

    'Truncate to compressed size
vbwProfiler.vbwExecuteLine 4061
    ReDim Preserve TheData(BufferSize - 1)
vbwProfiler.vbwExecuteLine 4062
    CopyMemory TheData(0), TempBuffer(0), BufferSize

    'Cleanup
vbwProfiler.vbwExecuteLine 4063
    Erase TempBuffer

    'Set properties if no error occurred
vbwProfiler.vbwExecuteLine 4064
    If result = 0 Then
vbwProfiler.vbwExecuteLine 4065
         CompressedSize = UBound(TheData) + 1
    End If
vbwProfiler.vbwExecuteLine 4066 'B

    'Return error code (if any)
vbwProfiler.vbwExecuteLine 4067
    CompressB = result

vbwProfiler.vbwProcOut 200
vbwProfiler.vbwExecuteLine 4068
End Function

Public Function UncompressB(ByRef TheData() As Byte) As Long
    '//this function decompresses a chunk of data that was compressed using the
    '//CompressB function.
vbwProfiler.vbwProcIn 201
    Dim result As Long
    Dim lngSize As Long

vbwProfiler.vbwExecuteLine 4069
    Const LONGLENGTH = 4
    'Allocate memory for buffers
    Dim BufferSize As Long
    Dim TempBuffer() As Byte

    '//extract the first long value from TheData which is our OrigSize
vbwProfiler.vbwExecuteLine 4070
    Call CopyMemory(lngSize, TheData(0), LONGLENGTH)

vbwProfiler.vbwExecuteLine 4071
    BufferSize = lngSize
vbwProfiler.vbwExecuteLine 4072
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
vbwProfiler.vbwExecuteLine 4073
    ReDim TempBuffer(BufferSize)

    'Decompress data
vbwProfiler.vbwExecuteLine 4074
    result = uncompress(TempBuffer(0), BufferSize, TheData(4), UBound(TheData) - 3)

    'Truncate buffer to compressed size
vbwProfiler.vbwExecuteLine 4075
    ReDim Preserve TheData(lngSize - 1)
vbwProfiler.vbwExecuteLine 4076
    CopyMemory TheData(0), TempBuffer(0), lngSize

    'Reset properties
vbwProfiler.vbwExecuteLine 4077
    If result = 0 Then
vbwProfiler.vbwExecuteLine 4078
    CompressedSize = 0
vbwProfiler.vbwExecuteLine 4079
    lngSize = 0
    End If
vbwProfiler.vbwExecuteLine 4080 'B

    'Return error code (if any)
vbwProfiler.vbwExecuteLine 4081
    UncompressB = result

vbwProfiler.vbwProcOut 201
vbwProfiler.vbwExecuteLine 4082
End Function


Public Property Get CompressedSize() As Long
vbwProfiler.vbwProcIn 202
vbwProfiler.vbwExecuteLine 4083
    CompressedSize = m_CompressedSize
vbwProfiler.vbwProcOut 202
vbwProfiler.vbwExecuteLine 4084
End Property

Public Property Let CompressedSize(ByVal New_CompressedSize As Long)
    'If Ambient.UserMode = False Then Exit Property
vbwProfiler.vbwProcIn 203
vbwProfiler.vbwExecuteLine 4085
    m_CompressedSize = New_CompressedSize

vbwProfiler.vbwProcOut 203
vbwProfiler.vbwExecuteLine 4086
End Property

Public Property Get OriginalSize() As Long
vbwProfiler.vbwProcIn 204
vbwProfiler.vbwExecuteLine 4087
    OriginalSize = m_OriginalSize
vbwProfiler.vbwProcOut 204
vbwProfiler.vbwExecuteLine 4088
End Property

Public Property Let OriginalSize(ByVal New_OriginalSize As Long)
   ' If Ambient.UserMode = False Then Exit Property
vbwProfiler.vbwProcIn 205
vbwProfiler.vbwExecuteLine 4089
    m_OriginalSize = New_OriginalSize

vbwProfiler.vbwProcOut 205
vbwProfiler.vbwExecuteLine 4090
End Property


