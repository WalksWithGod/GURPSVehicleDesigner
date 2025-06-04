Option Strict Off
Option Explicit On
Module modTimerProc
	
	' This timer is used for one purpose only in GVD... and thats to control the Clearing/Reloading of the
	' TreeX control.  There is a crashing bug in the TreeX control that prevents the programmer from
	' calling TreeX.RemoveAllItems from within the TreeX_ItemDragged() event.  Unfortunately, do to the way
	' im handling Profiles (e.g. Power and Fuel) one TreeX must accommodate many different profiles (Rather than
	' creating a seperate instance of the TreeX component for each!) and this necessitates that TreeX only be
	' used to display the internal profile and not be used to store the internal layouts of any profile.  As a
	' result, when items are dragged to new positions in the treeX, the TreeX cant be relied upon to have
	' updated index values or even that visible nodes exist within the internal represenation of that profile.
	' So the TreeX must be refreshed after a drag event.  So since we cant call .
	Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Integer, ByVal nIDEvent As Integer) As Integer
	Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Integer, ByVal nIDEvent As Integer, ByVal uElapse As Integer, ByVal lpTimerFunc As Integer) As Integer
	' SetTimer is public cuz its called from frmDesigner, however i wanted all these declares in the same routine
	
	Public m_lngTimerID As Integer
	Public Const TIMER_DELAY As Short = 1
	
	Sub TimerProc(ByVal hwnd As Integer, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTime As Integer)
		Dim m_oCurrentVeh As Object
		'TODO: This is dangerous in the IDE because if this function is called in a seperate thread
		' while the main thread also access the .Show method, we hang the IDE.  Not really likely to happen
		' in the compiled EXE since the user cant be fast enuf to drag an item in the TreeX and then
		' produce another call to .Show
		KillTimer(0, m_lngTimerID)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.ActiveProfile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Profiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).Show()
	End Sub
End Module