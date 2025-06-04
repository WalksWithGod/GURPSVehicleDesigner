Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cISystem_NET.cISystem")> Public Class cISystem
	
	Public Function isFunctional() As Boolean
		'todo: what about components like "beds" which aren't systems per se, but can be used by crew to "gain rest poitns" back?
		' otherwise this just mimics damage amount.  Like a bed, you would
		' only wonder if it was "functional" if it was heavily damaged.
		' as for a system that uses power, isfunctional would look at both
		' damage and power availability.  Need to consider this more
	End Function
End Class