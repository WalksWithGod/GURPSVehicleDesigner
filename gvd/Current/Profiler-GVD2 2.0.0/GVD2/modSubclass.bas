Attribute VB_Name = "modSubclass"
Option Explicit

Private Const GWL_WNDPROC = (-4)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_SIZE = &H5
Private Const WM_CONTEXTMENU = &H7B
Private Const WM_COMMAND = &H111
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_MOVE = &H3
Private Const WM_KEYDOWN = &H100

Public m_hOrigWndProc_ComboBox As Long
Public m_hOrigWndProc_TabStrip As Long

 ' todo: should rename these. This isnt really a hook, its subclassing (setwindowhookex would be using hooks to intercept messages to the main application )
Public Sub SetHook(hWnd, sType As String, bSet As Boolean)
vbwProfiler.vbwProcIn 206
    Dim hOrigWndProc
vbwProfiler.vbwExecuteLine 4091
    If bSet Then
vbwProfiler.vbwExecuteLine 4092
        If sType = "test1" Then

vbwProfiler.vbwExecuteLine 4093
            m_hOrigWndProc_ComboBox = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf ComboWndProc)
'vbwLine 4094:        ElseIf sType = "test2" Then
        ElseIf vbwProfiler.vbwExecuteLine(4094) Or sType = "test2" Then
vbwProfiler.vbwExecuteLine 4095
            m_hOrigWndProc_TabStrip = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf TabWndProc)
        Else
vbwProfiler.vbwExecuteLine 4096 'B
vbwProfiler.vbwExecuteLine 4097
            MsgBox "GVD_CUSTOM_ERROR: Unsupported Subclass Handle"
        End If
vbwProfiler.vbwExecuteLine 4098 'B
    Else
vbwProfiler.vbwExecuteLine 4099 'B
        Dim lRet As Long
vbwProfiler.vbwExecuteLine 4100
        If sType = "test1" Then
vbwProfiler.vbwExecuteLine 4101
            hOrigWndProc = m_hOrigWndProc_ComboBox
'vbwLine 4102:        ElseIf sType = "test2" Then
        ElseIf vbwProfiler.vbwExecuteLine(4102) Or sType = "test2" Then
vbwProfiler.vbwExecuteLine 4103
            hOrigWndProc = m_hOrigWndProc_TabStrip
        End If
vbwProfiler.vbwExecuteLine 4104 'B
vbwProfiler.vbwExecuteLine 4105
        lRet = SetWindowLong(hWnd, GWL_WNDPROC, hOrigWndProc)
    End If
vbwProfiler.vbwExecuteLine 4106 'B
vbwProfiler.vbwProcOut 206
vbwProfiler.vbwExecuteLine 4107
End Sub

Public Function ComboWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
vbwProfiler.vbwProcIn 207
vbwProfiler.vbwExecuteLine 4108
    On Error Resume Next
vbwProfiler.vbwExecuteLine 4109
    Select Case Msg

'vbwLine 4110:        Case WM_LBUTTONDOWN
        Case IIf(vbwProfiler.vbwExecuteLine(4110), VBWPROFILER_EMPTY, _
        WM_LBUTTONDOWN)

vbwProfiler.vbwExecuteLine 4111
            Call frmDesigner.ShowCustomDropDown
vbwProfiler.vbwExecuteLine 4112
            ComboWndProc = 0
vbwProfiler.vbwProcOut 207
vbwProfiler.vbwExecuteLine 4113
            Exit Function

'vbwLine 4114:        Case WM_KEYDOWN
        Case IIf(vbwProfiler.vbwExecuteLine(4114), VBWPROFILER_EMPTY, _
        WM_KEYDOWN)
vbwProfiler.vbwExecuteLine 4115
            ComboWndProc = 0
            #If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 4116
                InfoPrint 1, "modSubClass:ComboWndProc -- Trapped WM_KEYDOWN code = " & wParam
            #End If
vbwProfiler.vbwProcOut 207
vbwProfiler.vbwExecuteLine 4117
            Exit Function

       ' Case WM_CONTEXTMENU
       '     Form1.PopupMenu Form1.mnuBP
       '     AppWndProc = 0
       '     Exit Function
    End Select
vbwProfiler.vbwExecuteLine 4118 'B
vbwProfiler.vbwExecuteLine 4119
    ComboWndProc = CallWindowProc(m_hOrigWndProc_ComboBox, hWnd, Msg, wParam, lParam)
vbwProfiler.vbwProcOut 207
vbwProfiler.vbwExecuteLine 4120
End Function


Public Function TabWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
vbwProfiler.vbwProcIn 208
vbwProfiler.vbwExecuteLine 4121
    On Error Resume Next

vbwProfiler.vbwExecuteLine 4122
    Select Case Msg
        'Case WM_MOVE
        '    Debug.Print "TabWndProc:WM_MOVE"
        '    Call frmDesigner.TabStrip_Resize
'vbwLine 4123:        Case WM_SIZE
        Case IIf(vbwProfiler.vbwExecuteLine(4123), VBWPROFILER_EMPTY, _
        WM_SIZE)
vbwProfiler.vbwExecuteLine 4124
            Call frmDesigner.TabStrip_Resize
    End Select
vbwProfiler.vbwExecuteLine 4125 'B
vbwProfiler.vbwExecuteLine 4126
    TabWndProc = CallWindowProc(m_hOrigWndProc_TabStrip, hWnd, Msg, wParam, lParam)
vbwProfiler.vbwProcOut 208
vbwProfiler.vbwExecuteLine 4127
End Function


