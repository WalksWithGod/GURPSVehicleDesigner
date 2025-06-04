VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'todo: change this to just getPropertyCount, getProperty(byval index as long)
' just like i did for cINode.  The reason is simple, if i call getNextPropertyItem
' and then somewhere another function calls getFirstPropertyItem, the "m_lngCurrentProperty"
' variable gets MOVED and effects all the other functions that were calling it!  This is
' critical to change.
'//Public Functions
Public Function getFirstPropertyItem() As cPropertyItem
End Function
Public Function getNextPropertyItem() As cPropertyItem
End Function
Public Function getPropertyItemByIndex(ByVal l As Long) As cPropertyItem
End Function

