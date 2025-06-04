Attribute VB_Name = "modGuid"
Option Explicit

Private Declare Function CoCreateGuid Lib "OLE32.DLL" _
    (pGUID As GUID) As Long

Private Declare Function StringFromGUID2 Lib "OLE32.DLL" _
    (pGUID As GUID, _
     ByVal PointerToString As Long, _
     ByVal MaxLength As Long) As Long

Private Type GUID
    Guid1 As Long
    Guid2 As Long
    Guid3 As Long
    Guid4(0 To 7) As Byte
End Type

Public Function CreateGUID() As String
vbwProfiler.vbwProcIn 638

    Dim udtGUID As GUID
    Dim sGUID As String
    Dim lResult As Long

vbwProfiler.vbwExecuteLine 11511
    lResult = CoCreateGuid(udtGUID)

vbwProfiler.vbwExecuteLine 11512
    If lResult Then
vbwProfiler.vbwExecuteLine 11513
        sGUID = ""
    Else
vbwProfiler.vbwExecuteLine 11514 'B
vbwProfiler.vbwExecuteLine 11515
        sGUID = String$(38, 0)
vbwProfiler.vbwExecuteLine 11516
        StringFromGUID2 udtGUID, StrPtr(sGUID), 39
    End If
vbwProfiler.vbwExecuteLine 11517 'B

vbwProfiler.vbwExecuteLine 11518
    CreateGUID = sGUID

vbwProfiler.vbwProcOut 638
vbwProfiler.vbwExecuteLine 11519
End Function


