Option Strict Off
Option Explicit On
Module modSubclass
	
	Private Const GWL_WNDPROC As Short = (-4)
	Private Declare Function SetWindowLong Lib "user32"  Alias "SetWindowLongA"(ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	Private Declare Function CallWindowProc Lib "user32"  Alias "CallWindowProcA"(ByVal lpPrevWndFunc As Integer, ByVal hWnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
	
	Private Const WM_SIZE As Short = &H5s
	Private Const WM_CONTEXTMENU As Short = &H7Bs
	Private Const WM_COMMAND As Short = &H111s
	Private Const WM_LBUTTONDOWN As Short = &H201s
	Private Const WM_MOVE As Short = &H3s
	Private Const WM_KEYDOWN As Short = &H100s
	
	Public m_hOrigWndProc_ComboBox As Integer
	Public m_hOrigWndProc_TabStrip As Integer
	
	' todo: should rename these. This isnt really a hook, its subclassing (setwindowhookex would be using hooks to intercept messages to the main application )
	Public Sub SetHook(ByRef hWnd As Object, ByRef sType As String, ByRef bSet As Boolean)
		Dim hOrigWndProc As Object
		Dim lRet As Integer
		If bSet Then
			If sType = "test1" Then
				
				'UPGRADE_WARNING: Add a delegate for AddressOf ComboWndProc Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
				'UPGRADE_WARNING: Couldn't resolve default property of object hWnd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_hOrigWndProc_ComboBox = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf ComboWndProc)
			ElseIf sType = "test2" Then 
				'UPGRADE_WARNING: Add a delegate for AddressOf TabWndProc Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
				'UPGRADE_WARNING: Couldn't resolve default property of object hWnd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_hOrigWndProc_TabStrip = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf TabWndProc)
			Else
				MsgBox("GVD_CUSTOM_ERROR: Unsupported Subclass Handle")
			End If
		Else
			If sType = "test1" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object hOrigWndProc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				hOrigWndProc = m_hOrigWndProc_ComboBox
			ElseIf sType = "test2" Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object hOrigWndProc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				hOrigWndProc = m_hOrigWndProc_TabStrip
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object hOrigWndProc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object hWnd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lRet = SetWindowLong(hWnd, GWL_WNDPROC, hOrigWndProc)
		End If
	End Sub
	
	Public Function ComboWndProc(ByVal hWnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
		Dim frmDesigner As Object
		On Error Resume Next
		Select Case Msg
			
			Case WM_LBUTTONDOWN
				
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.ShowCustomDropDown. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call frmDesigner.ShowCustomDropDown()
				ComboWndProc = 0
				Exit Function
				
			Case WM_KEYDOWN
				ComboWndProc = 0
#If DEBUG_MODE Then
				'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
				InfoPrint 1, "modSubClass:ComboWndProc -- Trapped WM_KEYDOWN code = " & wParam
#End If
				Exit Function
				
				' Case WM_CONTEXTMENU
				'     Form1.PopupMenu Form1.mnuBP
				'     AppWndProc = 0
				'     Exit Function
		End Select
		ComboWndProc = CallWindowProc(m_hOrigWndProc_ComboBox, hWnd, Msg, wParam, lParam)
	End Function
	
	
	Public Function TabWndProc(ByVal hWnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
		Dim frmDesigner As Object
		On Error Resume Next
		
		Select Case Msg
			'Case WM_MOVE
			'    Debug.Print "TabWndProc:WM_MOVE"
			'    Call frmDesigner.TabStrip_Resize
			Case WM_SIZE
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.TabStrip_Resize. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call frmDesigner.TabStrip_Resize()
		End Select
		TabWndProc = CallWindowProc(m_hOrigWndProc_TabStrip, hWnd, Msg, wParam, lParam)
	End Function
End Module