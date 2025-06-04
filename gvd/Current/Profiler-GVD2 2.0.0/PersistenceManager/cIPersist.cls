VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Long)
'
vbwProfiler.vbwProcIn 7
vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 134
End Sub
Public Sub StoreProperties(ByVal op As clsObjProperties)
'
vbwProfiler.vbwProcIn 8
vbwProfiler.vbwProcOut 8
vbwProfiler.vbwExecuteLine 135
End Sub
Public Property Get Classname() As String
'
vbwProfiler.vbwProcIn 9
vbwProfiler.vbwProcOut 9
vbwProfiler.vbwExecuteLine 136
End Property
Public Property Get GUID() As String
' this is the GUID for a definition file
' this is required in our schema for all objects that are serializeable
vbwProfiler.vbwProcIn 10
vbwProfiler.vbwProcOut 10
vbwProfiler.vbwExecuteLine 137
End Property

