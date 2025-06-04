Attribute VB_Name = "modTimerProc"
Option Explicit

' This timer is used for one purpose only in GVD... and thats to control the Clearing/Reloading of the
' TreeX control.  There is a crashing bug in the TreeX control that prevents the programmer from
' calling TreeX.RemoveAllItems from within the TreeX_ItemDragged() event.  Unfortunately, do to the way
' im handling Profiles (e.g. Power and Fuel) one TreeX must accommodate many different profiles (Rather than
' creating a seperate instance of the TreeX component for each!) and this necessitates that TreeX only be
' used to display the internal profile and not be used to store the internal layouts of any profile.  As a
' result, when items are dragged to new positions in the treeX, the TreeX cant be relied upon to have
' updated index values or even that visible nodes exist within the internal represenation of that profile.
' So the TreeX must be refreshed after a drag event.  So since we cant call .
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
' SetTimer is public cuz its called from frmDesigner, however i wanted all these declares in the same routine

Public m_lngTimerID As Long
Public Const TIMER_DELAY = 1

Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    'TODO: This is dangerous in the IDE because if this function is called in a seperate thread
    ' while the main thread also access the .Show method, we hang the IDE.  Not really likely to happen
    ' in the compiled EXE since the user cant be fast enuf to drag an item in the TreeX and then
    ' produce another call to .Show
    KillTimer 0, m_lngTimerID
    m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).Show
End Sub
