Attribute VB_Name = "modCTemplate"
Option Explicit

Public Type CT_Version
    Major As Long
    Minor As Long
    Revision As Long
    Reserved As Long
    Misc As Long
End Type

Public Type THeader
    Magic As Long
    CRC32 As Long
    GUID As String
    ComponentName As String
    Author As String * 128
    Email As String * 128
    URL As String * 256
    Modified As Date
    Version As CT_Version
    Reserved As Long
    Misc As Long
End Type
