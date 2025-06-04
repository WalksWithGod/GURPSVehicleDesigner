Option Strict Off
Option Explicit On
Public Interface _cIDisplay
	Function getFirstPropertyItem() As cPropertyItem
	Function getNextPropertyItem() As cPropertyItem
	Function getPropertyItemByIndex(ByVal l As Integer) As cPropertyItem
End Interface
<System.Runtime.InteropServices.ProgId("cIDisplay_NET.cIDisplay")> Public Class cIDisplay
	Implements _cIDisplay
	
	'todo: change this to just getPropertyCount, getProperty(byval index as long)
	' just like i did for cINode.  The reason is simple, if i call getNextPropertyItem
	' and then somewhere another function calls getFirstPropertyItem, the "m_lngCurrentProperty"
	' variable gets MOVED and effects all the other functions that were calling it!  This is
	' critical to change.
	'//Public Functions
	Public Function getFirstPropertyItem() As cPropertyItem Implements _cIDisplay.getFirstPropertyItem
	End Function
	Public Function getNextPropertyItem() As cPropertyItem Implements _cIDisplay.getNextPropertyItem
	End Function
	Public Function getPropertyItemByIndex(ByVal l As Integer) As cPropertyItem Implements _cIDisplay.getPropertyItemByIndex
	End Function
End Class