VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type size
        cx As Long
        cy As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long

Private m_frm As Object   ' the form or picturebox this splitter is attached to
Private WithEvents m_imgSplitter As PictureBox
Attribute m_imgSplitter.VB_VarHelpID = -1



Private m_bOrientation As Boolean ' if TRUE, the control will be horizontal

Private m_hWnd As Long
Private m_lngSplitterColor As Long
Private m_lngLeft As Long
Private m_lngTop As Long
Private m_lngHeight As Long
Private m_lngWidth As Long
Private m_ctlTopLeft() As Object  ' all controls that are on the top or left side of this splitter
Private m_ctlBottomRight() As Object ' all the controls that are on the bottom or right side of this

Private m_lngMaxLeft As Long
Private m_lngMinLeft As Long
Private m_lngMaxTop As Long
Private m_lngMinTop As Long


Private m_rectClient As Rect


Public Enum ControlOrientation
    ctlTopLeft = False
    ctlbottomright = True
End Enum

Public Enum SplitterOrientation
    splitvertical = False
    splithorizontal = True
End Enum


Private m_lngTopPadding As Long
Private m_lngBottomPadding As Long
Private m_lngLeftPadding As Long
Private m_lngRightPadding As Long



'TODO: Thoughts on improving this code:
'1)  When parent is resized, first change HEIGHT of splitter itself if Verticle, or WIDTH if Horizontal...then
'2)  Only excersise code for Parent Resize _IF_ last Height (or previous Width) is different.  Else, just call the normal Move code
'3)  Determine what should set the Min/Max/ vars.  Im thinking the parent should set these... and the client shouldn't calc these themselves right?
'    this would probably be better for also later adding Hidden panes (just set Min to extreme left or right, then auto move it there and then disable
'   the Image so it cant be moved.
'4)

' STILL TODO:
'1 ) fix the borders/padding code... should be more flexible too right?
'2) Should code be here or in parents to adjust locations of splitters when resizing such that they are offscreen or covered by another splitter?
'    actually,  i should call the MOVE code of the Splitter to auto move that to a new location.  The parent can determine where it should
' move to, but otherwise the code is already there to handle the movement and repositioning
'3) Still add code to inform Vsplitter2 that its parent has resized.

Event Resize()


Public Property Let SplitterColor(ByVal value As Long)
    m_lngSplitterColor = value
End Property

Public Property Let Orientation(ByVal value As SplitterOrientation)
    m_bOrientation = value
    m_imgSplitter.MousePointer = IIf(m_bOrientation, vbSizeNS, vbSizeWE)
    
End Property



Public Property Let hWnd(ByVal lngValue As Long)
    m_hWnd = lngValue
End Property

Public Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Public Property Let Width(ByVal lngValue As Long)
    m_lngWidth = lngValue
    Call ResizeSplitter
End Property

Public Property Get Width() As Long
    Width = m_lngWidth
End Property

Public Property Let Height(ByVal lngValue As Long)
    m_lngHeight = lngValue
    Call ResizeSplitter
End Property

Public Property Get Height() As Long
    Height = m_lngHeight
End Property

Public Property Let Top(ByVal lngValue As Long)
    m_lngTop = lngValue
    Call ResizeSplitter
End Property

Public Property Get Top() As Long
    Top = m_lngTop
End Property

Public Property Let Left(ByVal lngValue As Long)
    m_lngLeft = lngValue
    Call ResizeSplitter
End Property

Public Property Get Left() As Long
    Left = m_lngLeft
End Property

Public Property Let MaxLeft(ByVal lngValue As Long)
    m_lngMaxLeft = lngValue
End Property

Public Property Let MinLeft(ByVal lngValue As Long)
    m_lngMinLeft = lngValue
End Property

Public Property Let MaxTop(ByVal lngValue As Long)
    m_lngMaxTop = lngValue
End Property

Public Property Let MinTop(ByVal lngValue As Long)
    m_lngMinTop = lngValue
End Property

Public Property Let LeftPadding(ByVal lngValue As Long)
    m_lngLeftPadding = lngValue
End Property

Public Property Get LeftPadding() As Long
    LeftPadding = m_lngLeftPadding
End Property

Public Property Let RightPadding(ByVal lngValue As Long)
    m_lngRightPadding = lngValue
End Property

Public Property Get RightPadding() As Long
    RightPadding = m_lngRightPadding
End Property

Public Property Let TopPadding(ByVal lngValue As Long)
    m_lngTopPadding = lngValue
End Property

Public Property Get TopPadding() As Long
    TopPadding = m_lngTopPadding
End Property

Public Property Let BottomPadding(ByVal lngValue As Long)
    m_lngBottomPadding = lngValue
End Property

Public Property Get BottomPadding() As Long
    BottomPadding = m_lngBottomPadding
End Property


Public Sub AddControl(ByRef ctl As Object, ByVal loc As ControlOrientation)
    
    Dim lngUpper As Long
    
    'ctl.ScaleMode = vbPixels
    
    If loc = ctlTopLeft Then
        If m_ctlTopLeft(0) Is Nothing Then
            Set m_ctlTopLeft(0) = ctl
            
        Else
            lngUpper = UBound(m_ctlTopLeft) + 1
            ReDim Preserve m_ctlTopLeft(lngUpper)
            Set m_ctlTopLeft(lngUpper) = ctl
        End If
    Else
        If m_ctlBottomRight(0) Is Nothing Then
            Set m_ctlBottomRight(0) = ctl
           
        Else
            lngUpper = UBound(m_ctlBottomRight) + 1
            ReDim Preserve m_ctlBottomRight(lngUpper)
            Set m_ctlBottomRight(lngUpper) = ctl
        End If
    End If
    
    m_imgSplitter.ZOrder
End Sub

Public Sub RemoveAllControls()
    Dim i As Long
    On Error GoTo err
    
    For i = 0 To UBound(m_ctlTopLeft)
        Set m_ctlTopLeft(i) = Nothing
    Next
    
    For i = 0 To UBound(m_ctlBottomRight)
        Set m_ctlBottomRight(i) = Nothing
    Next
    Exit Sub
err:
    Debug.Print "Error in cSplitter.RemoveAllControls() - " & err.Description
    Resume Next
    
End Sub
Private Sub Class_Initialize()

    ReDim m_ctlTopLeft(0)
    ReDim m_ctlBottomRight(0)
    m_lngSplitterColor = vbWhite
End Sub

Private Sub Class_Terminate()
    Dim i As Long
    On Error GoTo err
    
    Call RemoveAllControls
    
    frmDesigner.Controls.Remove (m_imgSplitter.Tag)
    Set m_imgSplitter = Nothing
    Set m_frm = Nothing
    Exit Sub
    
err:
    Debug.Print "Error in cSplitter.Terminate() - " & err.Description
    Resume Next
End Sub

Private Sub ResizeSplitter()
    Call MoveWindow(m_hWnd, m_lngLeft, m_lngTop, m_lngWidth, m_lngHeight, True)
End Sub

Private Sub GetClientDimensions(ByRef oCtl As Object, ByRef lngLeft As Long, ByRef lngTop As Long, ByRef lngWidth As Long, ByRef lngHeight As Long)
   
   Dim r As Rect
   Dim pt As POINTAPI
   Dim P(0 To 1) As POINTAPI
   
    On Error GoTo err
    
   ' store the original dimensions of the treeview
    If TypeOf oCtl Is PropListCtl Then
        GetWindowRect oCtl.Window, r  ' lame ass proplistctl uses .Window instead of .hwnd so test for it
    Else
        GetWindowRect oCtl.hWnd, r
    End If
    
    P(0).x = r.Left
    P(0).y = r.Top
    P(1).x = r.Right
    P(1).y = r.Bottom
    
    ScreenToClient m_frm.hWnd, P(0)
    ScreenToClient m_frm.hWnd, P(1)
    lngLeft = P(0).x
    lngTop = P(0).y
    lngHeight = P(1).y - P(0).y
    lngWidth = P(1).x - P(0).x
    Exit Sub
    
err:
    Debug.Print "cSplitter:GetClientDimensions() -- " & err.Description
    Exit Sub
End Sub


Public Sub SetPosition(ByRef frm As Object, ByVal UniqueKey As String, ByVal Top As Long, ByVal Left As Long, ByVal Width As Long, ByVal Height As Long)
    
    ' create our Image control
    If Not m_frm Is Nothing Then Set m_frm = Nothing
    Set m_frm = frm
    
    If Not m_imgSplitter Is Nothing Then Set m_imgSplitter = Nothing
    
    Set m_imgSplitter = frmDesigner.Controls.Add("VB.PictureBox", UniqueKey, frm)
    ' set the hwnd property to the same as the image splitter's
    m_hWnd = m_imgSplitter.hWnd
    
    m_imgSplitter.Tag = UniqueKey
    m_lngLeft = Left
    m_lngTop = Top
    m_lngWidth = Width
    m_lngHeight = Height
    
    m_imgSplitter.BorderStyle = 0 'none
    
    Call ResizeSplitter
    m_imgSplitter.Visible = True

End Sub

Public Sub SetPadding(ByVal Top As Long, ByVal Left As Long, ByVal Right As Long, ByVal Bottom As Long)

    m_lngTopPadding = Top
    m_lngLeftPadding = Left
    m_lngRightPadding = Right
    m_lngBottomPadding = Bottom

    Call ParentResized
End Sub

Public Sub Move(ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngWidth As Long, ByVal lngHeight As Long)

  
    m_lngTop = lngTop
    m_lngLeft = lngLeft
    m_lngWidth = lngWidth
    m_lngHeight = lngHeight
    Call ResizeSplitter
  
End Sub

Private Sub m_imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    With m_imgSplitter
        .BackColor = m_lngSplitterColor     ' Make the splitter visible
        .ZOrder
    End With
End Sub

Private Sub m_imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    m_imgSplitter.BackColor = frmDesigner.BackColor   ' change the color back to normal

End Sub



Private Sub m_imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim lppt As POINTAPI
    Dim lngCurrentLeft As Long
    Dim lngCurrentTop As Long
    Dim lngCurrentWidth As Long
    Dim lngCurrentHeight As Long
    Dim lngXDiff As Long
    Dim lngYDiff As Long
    Dim rControl As Rect
    
    Dim i As Long
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngHeight As Long
    Dim lngWidth As Long
    
    On Error Resume Next
    lngCurrentLeft = m_lngLeft
    lngCurrentTop = m_lngTop
    lngCurrentHeight = m_lngHeight
    lngCurrentWidth = m_lngWidth
    
    If Button = vbLeftButton Then
        Call GetCursorPos(lppt)
        Call ScreenToClient(m_frm.hWnd, lppt)
        
        '01/07/02 - Using Pixels instead of Twips to overcome resize limitation of Tabbed dialog. Switching ALL
        ' resize instances to MoveWindow() for consistancy therefore converting to Twips is no longer necessary
        ' convert the x, y values to twips
        'lppt.y = lppt.y * Screen.TwipsPerPixelY
        'lppt.x = lppt.x * Screen.TwipsPerPixelX
        
        ' move the splitter
        Select Case m_bOrientation
            Case splithorizontal
                ' check for Max and Min Y coordinate restrictions
                If lppt.y > m_lngMaxTop Then
                    m_lngTop = m_lngMaxTop
                ElseIf lppt.y < m_lngMinTop Then
                    m_lngTop = m_lngMinTop
                Else
                    m_lngTop = lppt.y
                End If
                
                ' calc the difference in movement from last position
                lngYDiff = m_lngTop - lngCurrentTop
                If lngYDiff = 0 Then Exit Sub ' short circuit rest of sub if possible
                
                ' adjust the HEIGHT for controls above the splitter
                For i = 0 To UBound(m_ctlTopLeft)
                    If Not m_ctlTopLeft(i) Is Nothing Then
                        Call GetClientDimensions(m_ctlTopLeft(i), lngLeft, lngTop, lngWidth, lngHeight)
                        lngTop = m_lngMinTop
                        lngHeight = m_lngTop - lngTop
                        Call MoveWindow(m_ctlTopLeft(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, True)
                    End If
                Next
                
                ' adjust the HEIGHT and TOPS for controls below the splitter
                For i = 0 To UBound(m_ctlBottomRight)
                    If Not m_ctlBottomRight(i) Is Nothing Then
                        Call GetClientDimensions(m_ctlBottomRight(i), lngLeft, lngTop, lngWidth, lngHeight)
                        lngTop = m_lngTop + m_lngHeight
                        lngHeight = m_lngMaxTop - lngTop
                        Call MoveWindow(m_ctlBottomRight(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, True)
                    End If
                Next
        
            ' split verticle
            Case Else
            
                ' check for Max and Min X coordinate restrictions
                If lppt.x < m_lngMinLeft Then
                    m_lngLeft = m_lngMinLeft
                    Exit Sub
                ElseIf lppt.x > m_lngMaxLeft Then
                    m_lngLeft = m_lngMaxLeft
                    Exit Sub
                Else
                    m_lngLeft = lppt.x
                End If
               
                ' calc the difference in movement from last position
                lngXDiff = m_lngLeft - lngCurrentLeft
                If lngXDiff = 0 Then Exit Sub ' short circuit rest of sub if possible
                
                ' adjust the WIDTHS for controls LEFT of the veritcle splitter
                For i = 0 To UBound(m_ctlTopLeft)
                     If Not m_ctlTopLeft(i) Is Nothing Then
                        Call GetClientDimensions(m_ctlTopLeft(i), lngLeft, lngTop, lngWidth, lngHeight)
                        lngLeft = m_lngMinLeft
                        lngWidth = m_lngLeft - lngLeft
                        If TypeOf m_ctlTopLeft(i) Is TreeX Then
                            Debug.Print m_ctlTopLeft(i).Width & " TREEX BEFORE"
                        End If
                        
                        If TypeOf m_ctlTopLeft(i) Is TreeX Then
                            With m_ctlTopLeft(i)
                                .Left = lngLeft * Screen.TwipsPerPixelX
                               ' .Top = lngTop * Screen.TwipsPerPixelY
                                .Width = lngWidth * Screen.TwipsPerPixelX
                                '.Height = lngHeight * Screen.TwipsPerPixelY
                            End With
                        Else
                            Call MoveWindow(m_ctlTopLeft(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, True)
                        End If
                     End If
                Next
                ' adjust the LEFTS and WIDTHS for controls RIGHT of the verticle splitter
                For i = 0 To UBound(m_ctlBottomRight)
                    If Not m_ctlBottomRight(i) Is Nothing Then
                        Call GetClientDimensions(m_ctlBottomRight(i), lngLeft, lngTop, lngWidth, lngHeight)
                        lngLeft = m_lngLeft + m_lngWidth
                        lngWidth = m_lngMaxLeft - m_lngLeft
                                               
                        ' if this is the ONLY control on the right side, we can adjust the height value as well
                        If UBound(m_ctlBottomRight) = 0 Then
                            lngHeight = m_lngHeight
                        End If
                        
                        If (TypeOf m_ctlBottomRight(i) Is PropListCtl) Or (TypeOf m_ctlBottomRight(i) Is ListView) Then
                            'NOTE: The lame ass propertylist control will not save these location settings when
                            ' MoveWindow API is used to resize/move the window. Therefore, we must set them explicitly
                            ' using TWIPS
                            'Debug.Print "PLC1 BEFORE LEFT  = " & frmDesigner.PLC1.Left & " WIDTH = " & frmDesigner.PLC1.Width
                            'Call MoveWindow(m_ctlBottomRight(i).Window, lngLeft, lngTop, lngWidth, lngHeight, True)
                            'Debug.Print "PLC1  AFTER LEFT  = " & frmDesigner.PLC1.Left & " WIDTH = " & frmDesigner.PLC1.Width
                            
                            '06/09/02 MPJ
                            'NOTE2: Damn, the ListView is also having the same problem.  Its internally not updating the LEFT/Top/Width/Height
                            ' properties correctly when using the MoveWindow API... the effect is that vbBeginDrag starts with the mouse
                            ' pointer over the wrong control!  So we cant use MoveWindow API here... dont know why this is so.
                            With m_ctlBottomRight(i)
                                .Left = lngLeft * Screen.TwipsPerPixelX
                                .Width = lngWidth * Screen.TwipsPerPixelX
                                .Top = lngTop * Screen.TwipsPerPixelY
                                .Height = lngHeight * Screen.TwipsPerPixelY
                            End With
                        Else
                            Call MoveWindow(m_ctlBottomRight(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, True)
                        End If
                        If TypeOf m_ctlBottomRight(i) Is cSplitter Then
                            ' when calling MoveWindow on this control, it only updates the PicBox and not the actual internally stored  dimensions of this cSplitter class.
                            ' so we (for now) are explicitly setting the proper internal dimensions of the control from here
                            With m_ctlBottomRight(i)
                                .Left = lngLeft
                                .Top = lngTop
                                .Height = lngHeight
                                .Width = lngWidth
                            End With
                        End If
                    End If
                Next
        End Select
        
        ' move the splitter itself
         Call Move(m_lngLeft, m_lngTop, m_lngWidth, m_lngHeight)
         RaiseEvent Resize
    End If

End Sub

Public Sub ParentResized()
' when the parent window resizes, the client area that the splitter must be bound by, should be updated.
' these values are stored in the RECT m_rectClient

' min max boundaries need to be updated


' Next, the size of the splitter needs to be updated

' All child controls need to be resized (Do i call ResizeControls or a seperate sub for this?)


Dim ret As Long
On Error Resume Next
ret = GetClientRect(m_frm.hWnd, m_rectClient)

' 01/07/02 - No longer using TWIPS since switching resize method to MoveWindow which uses pixels only
' Note: MoveWindow API is being used instead because it overcomes the Resize limits of the Tab Dialog control
'm_rectClient.Left = m_rectClient.Left * Screen.TwipsPerPixelX
'm_rectClient.Right = m_rectClient.Right * Screen.TwipsPerPixelX
'm_rectClient.Top = m_rectClient.Top * Screen.TwipsPerPixelY
'm_rectClient.Bottom = m_rectClient.Bottom * Screen.TwipsPerPixelY

m_lngMinTop = m_rectClient.Top + m_lngTopPadding
m_lngMaxTop = m_rectClient.Bottom - m_lngBottomPadding
m_lngMaxLeft = m_rectClient.Right - m_lngRightPadding
m_lngMinLeft = m_rectClient.Left + m_lngLeftPadding

If m_bOrientation = splithorizontal Then
    Width = m_rectClient.Right - m_lngLeft
    
Else
    Top = m_lngMinTop
    Height = m_lngMaxTop - m_lngMinTop 'm_rectClient.Bottom - m_lngTop
End If

Call ResizeControls
         
    
End Sub




Public Sub ResizeControls()

    Dim lngCurrentLeft As Long
    Dim lngCurrentTop As Long
    Dim i As Long
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngHeight As Long
    Dim lngWidth As Long
    Dim lngUpper As Long
    Dim rControl As Rect
    
    On Error Resume Next
    If m_bOrientation = splithorizontal Then
        lngUpper = UBound(m_ctlTopLeft)
        For i = 0 To lngUpper
            If Not m_ctlTopLeft(i) Is Nothing Then
                Call GetClientDimensions(m_ctlTopLeft(i), lngLeft, lngTop, lngWidth, lngHeight)
                
                lngLeft = m_lngLeft
                lngTop = m_lngMinTop
                lngHeight = m_lngTop - lngTop
                lngWidth = m_rectClient.Right - lngLeft - m_lngRightPadding
                
                Call MoveWindow(m_ctlTopLeft(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, 1)
            End If
        Next
        lngUpper = UBound(m_ctlBottomRight)
        For i = 0 To lngUpper
            If Not m_ctlBottomRight(i) Is Nothing Then
                Call GetClientDimensions(m_ctlBottomRight(i), lngLeft, lngTop, lngWidth, lngHeight)
                lngLeft = m_lngLeft
                lngTop = m_lngTop + m_lngHeight
                lngWidth = m_rectClient.Right - lngLeft - m_lngRightPadding
                
                If i = lngUpper Then
                    lngHeight = m_rectClient.Bottom - lngTop - m_lngBottomPadding
                End If
                
                Call MoveWindow(m_ctlBottomRight(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, 1)
            End If
        Next
    Else ' verticle
        lngUpper = UBound(m_ctlTopLeft)
        For i = 0 To lngUpper
            If Not m_ctlTopLeft(i) Is Nothing Then
                
                lngLeft = m_lngMinLeft
                lngTop = m_lngMinTop
                
                lngHeight = m_lngHeight
                lngWidth = m_lngLeft - lngLeft
                If TypeOf m_ctlTopLeft(i) Is TreeX Then
                    With m_ctlTopLeft(i)
                        .Left = lngLeft * Screen.TwipsPerPixelX
                        .Top = lngTop * Screen.TwipsPerPixelY
                        .Width = lngWidth * Screen.TwipsPerPixelX
                        .Height = lngHeight * Screen.TwipsPerPixelY
                    End With
                Else
                    Call MoveWindow(m_ctlTopLeft(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, 1)
                End If
            End If
        
        Next
        
        lngUpper = UBound(m_ctlBottomRight)
        If m_ctlBottomRight(lngUpper) Is Nothing Then
           lngUpper = lngUpper - 1
        End If
        For i = 0 To lngUpper
            If Not m_ctlBottomRight(i) Is Nothing Then
                Call GetClientDimensions(m_ctlBottomRight(i), lngLeft, lngTop, lngWidth, lngHeight)
                lngLeft = m_lngLeft + m_lngWidth
                If i = 0 Then lngTop = m_lngMinTop
                
               ' ' never change the height of a child Splitter that's horizontal
                If TypeOf m_ctlBottomRight(i) Is cSplitter Then
                    lngHeight = m_ctlBottomRight(i).Height
                Else
                    If i = lngUpper Then                   'if last control, its height can change
                        lngHeight = m_lngMaxTop - lngTop    ' child control's height is same as the splitters height
                    End If
                End If
                
                'lngWidth = m_frm.Width - m_lngLeft - m_lngRightPadding - m_lngWidth  '          'm_ctlTopLeft(i).Width
                lngWidth = m_lngMaxLeft - (m_lngLeft + m_lngWidth)
                
                If TypeOf m_ctlBottomRight(i) Is PropListCtl Then
                    'Call MoveWindow(m_ctlBottomRight(i).Window, lngLeft, lngTop, lngWidth, lngHeight, 1)
                    With m_ctlBottomRight(i)
                        .Left = lngLeft * Screen.TwipsPerPixelX
                        .Width = lngWidth * Screen.TwipsPerPixelX
                        .Top = lngTop * Screen.TwipsPerPixelY
                        .Height = lngHeight * Screen.TwipsPerPixelY
                    End With
                Else
                    Call MoveWindow(m_ctlBottomRight(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, 1)
                End If
            End If
        Next
    End If
    
    RaiseEvent Resize
        
End Sub

