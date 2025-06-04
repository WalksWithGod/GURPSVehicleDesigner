Option Strict Off
Option Explicit On
Module modOutputConstants
	
	'This contains constants used for print output text.  This module is shared across both projects in the group.
	
	' category ID's
	Public Const CID_SUBASSEMBLY As Short = 0
	
	
	' these fall into the formatting stuff
	Public Const START_B As Short = 1
	Public Const END_B As Short = 2
	
	
	'
	Public Const USER As Short = 3
	
	' these fall into the attribute names
	Public Const DESCR As String = "CustomDescription"
	Public Const ORIENT As String = "Orientation"
End Module