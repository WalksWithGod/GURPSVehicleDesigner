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
    Dim hOrigWndProc
    If bSet Then
        If sType = "test1" Then
        
            m_hOrigWndProc_ComboBox = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf ComboWndProc)
        ElseIf sType = "test2" Then
            m_hOrigWndProc_TabStrip = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf TabWndProc)
        Else
            MsgBox "GVD_CUSTOM_ERROR: Unsupported Subclass Handle"
        End If
    Else
        Dim lRet As Long
        If sType = "test1" Then
            hOrigWndProc = m_hOrigWndProc_ComboBox
        ElseIf sType = "test2" Then
            hOrigWndProc = m_hOrigWndProc_TabStrip
        End If
        lRet = SetWindowLong(hWnd, GWL_WNDPROC, hOrigWndProc)
    End If
End Sub

Public Function ComboWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Select Case Msg
    
        Case WM_LBUTTONDOWN
            
            Call frmDesigner.ShowCustomDropDown
            ComboWndProc = 0
            Exit Function
       
        Case WM_KEYDOWN
            ComboWndProc = 0
            #If DEBUG_MODE Then
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


Public Function TabWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    
    Select Case Msg
        'Case WM_MOVE
        '    Debug.Print "TabWndProc:WM_MOVE"
        '    Call frmDesigner.TabStrip_Resize
        Case WM_SIZE
            Call frmDesigner.TabStrip_Resize
    End Select
    TabWndProc = CallWindowProc(m_hOrigWndProc_TabStrip, hWnd, Msg, wParam, lParam)
End Function
