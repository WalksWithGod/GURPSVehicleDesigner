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
vbwProfiler.vbwProcIn 158
vbwProfiler.vbwExecuteLine 3658
    m_lngSplitterColor = value
vbwProfiler.vbwProcOut 158
vbwProfiler.vbwExecuteLine 3659
End Property

Public Property Let Orientation(ByVal value As SplitterOrientation)
vbwProfiler.vbwProcIn 159
vbwProfiler.vbwExecuteLine 3660
    m_bOrientation = value
vbwProfiler.vbwExecuteLine 3661
    m_imgSplitter.MousePointer = IIf(m_bOrientation, vbSizeNS, vbSizeWE)

vbwProfiler.vbwProcOut 159
vbwProfiler.vbwExecuteLine 3662
End Property



Public Property Let hWnd(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 160
vbwProfiler.vbwExecuteLine 3663
    m_hWnd = lngValue
vbwProfiler.vbwProcOut 160
vbwProfiler.vbwExecuteLine 3664
End Property

Public Property Get hWnd() As Long
vbwProfiler.vbwProcIn 161
vbwProfiler.vbwExecuteLine 3665
    hWnd = m_hWnd
vbwProfiler.vbwProcOut 161
vbwProfiler.vbwExecuteLine 3666
End Property

Public Property Let Width(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 162
vbwProfiler.vbwExecuteLine 3667
    m_lngWidth = lngValue
vbwProfiler.vbwExecuteLine 3668
    Call ResizeSplitter
vbwProfiler.vbwProcOut 162
vbwProfiler.vbwExecuteLine 3669
End Property

Public Property Get Width() As Long
vbwProfiler.vbwProcIn 163
vbwProfiler.vbwExecuteLine 3670
    Width = m_lngWidth
vbwProfiler.vbwProcOut 163
vbwProfiler.vbwExecuteLine 3671
End Property

Public Property Let Height(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 164
vbwProfiler.vbwExecuteLine 3672
    m_lngHeight = lngValue
vbwProfiler.vbwExecuteLine 3673
    Call ResizeSplitter
vbwProfiler.vbwProcOut 164
vbwProfiler.vbwExecuteLine 3674
End Property

Public Property Get Height() As Long
vbwProfiler.vbwProcIn 165
vbwProfiler.vbwExecuteLine 3675
    Height = m_lngHeight
vbwProfiler.vbwProcOut 165
vbwProfiler.vbwExecuteLine 3676
End Property

Public Property Let Top(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 166
vbwProfiler.vbwExecuteLine 3677
    m_lngTop = lngValue
vbwProfiler.vbwExecuteLine 3678
    Call ResizeSplitter
vbwProfiler.vbwProcOut 166
vbwProfiler.vbwExecuteLine 3679
End Property

Public Property Get Top() As Long
vbwProfiler.vbwProcIn 167
vbwProfiler.vbwExecuteLine 3680
    Top = m_lngTop
vbwProfiler.vbwProcOut 167
vbwProfiler.vbwExecuteLine 3681
End Property

Public Property Let Left(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 168
vbwProfiler.vbwExecuteLine 3682
    m_lngLeft = lngValue
vbwProfiler.vbwExecuteLine 3683
    Call ResizeSplitter
vbwProfiler.vbwProcOut 168
vbwProfiler.vbwExecuteLine 3684
End Property

Public Property Get Left() As Long
vbwProfiler.vbwProcIn 169
vbwProfiler.vbwExecuteLine 3685
    Left = m_lngLeft
vbwProfiler.vbwProcOut 169
vbwProfiler.vbwExecuteLine 3686
End Property

Public Property Let MaxLeft(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 170
vbwProfiler.vbwExecuteLine 3687
    m_lngMaxLeft = lngValue
vbwProfiler.vbwProcOut 170
vbwProfiler.vbwExecuteLine 3688
End Property

Public Property Let MinLeft(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 171
vbwProfiler.vbwExecuteLine 3689
    m_lngMinLeft = lngValue
vbwProfiler.vbwProcOut 171
vbwProfiler.vbwExecuteLine 3690
End Property

Public Property Let MaxTop(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 172
vbwProfiler.vbwExecuteLine 3691
    m_lngMaxTop = lngValue
vbwProfiler.vbwProcOut 172
vbwProfiler.vbwExecuteLine 3692
End Property

Public Property Let MinTop(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 173
vbwProfiler.vbwExecuteLine 3693
    m_lngMinTop = lngValue
vbwProfiler.vbwProcOut 173
vbwProfiler.vbwExecuteLine 3694
End Property

Public Property Let LeftPadding(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 174
vbwProfiler.vbwExecuteLine 3695
    m_lngLeftPadding = lngValue
vbwProfiler.vbwProcOut 174
vbwProfiler.vbwExecuteLine 3696
End Property

Public Property Get LeftPadding() As Long
vbwProfiler.vbwProcIn 175
vbwProfiler.vbwExecuteLine 3697
    LeftPadding = m_lngLeftPadding
vbwProfiler.vbwProcOut 175
vbwProfiler.vbwExecuteLine 3698
End Property

Public Property Let RightPadding(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 176
vbwProfiler.vbwExecuteLine 3699
    m_lngRightPadding = lngValue
vbwProfiler.vbwProcOut 176
vbwProfiler.vbwExecuteLine 3700
End Property

Public Property Get RightPadding() As Long
vbwProfiler.vbwProcIn 177
vbwProfiler.vbwExecuteLine 3701
    RightPadding = m_lngRightPadding
vbwProfiler.vbwProcOut 177
vbwProfiler.vbwExecuteLine 3702
End Property

Public Property Let TopPadding(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 178
vbwProfiler.vbwExecuteLine 3703
    m_lngTopPadding = lngValue
vbwProfiler.vbwProcOut 178
vbwProfiler.vbwExecuteLine 3704
End Property

Public Property Get TopPadding() As Long
vbwProfiler.vbwProcIn 179
vbwProfiler.vbwExecuteLine 3705
    TopPadding = m_lngTopPadding
vbwProfiler.vbwProcOut 179
vbwProfiler.vbwExecuteLine 3706
End Property

Public Property Let BottomPadding(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 180
vbwProfiler.vbwExecuteLine 3707
    m_lngBottomPadding = lngValue
vbwProfiler.vbwProcOut 180
vbwProfiler.vbwExecuteLine 3708
End Property

Public Property Get BottomPadding() As Long
vbwProfiler.vbwProcIn 181
vbwProfiler.vbwExecuteLine 3709
    BottomPadding = m_lngBottomPadding
vbwProfiler.vbwProcOut 181
vbwProfiler.vbwExecuteLine 3710
End Property


Public Sub AddControl(ByRef ctl As Object, ByVal loc As ControlOrientation)
vbwProfiler.vbwProcIn 182

    Dim lngUpper As Long

    'ctl.ScaleMode = vbPixels

vbwProfiler.vbwExecuteLine 3711
    If loc = ctlTopLeft Then
vbwProfiler.vbwExecuteLine 3712
        If m_ctlTopLeft(0) Is Nothing Then
vbwProfiler.vbwExecuteLine 3713
            Set m_ctlTopLeft(0) = ctl

        Else
vbwProfiler.vbwExecuteLine 3714 'B
vbwProfiler.vbwExecuteLine 3715
            lngUpper = UBound(m_ctlTopLeft) + 1
vbwProfiler.vbwExecuteLine 3716
            ReDim Preserve m_ctlTopLeft(lngUpper)
vbwProfiler.vbwExecuteLine 3717
            Set m_ctlTopLeft(lngUpper) = ctl
        End If
vbwProfiler.vbwExecuteLine 3718 'B
    Else
vbwProfiler.vbwExecuteLine 3719 'B
vbwProfiler.vbwExecuteLine 3720
        If m_ctlBottomRight(0) Is Nothing Then
vbwProfiler.vbwExecuteLine 3721
            Set m_ctlBottomRight(0) = ctl

        Else
vbwProfiler.vbwExecuteLine 3722 'B
vbwProfiler.vbwExecuteLine 3723
            lngUpper = UBound(m_ctlBottomRight) + 1
vbwProfiler.vbwExecuteLine 3724
            ReDim Preserve m_ctlBottomRight(lngUpper)
vbwProfiler.vbwExecuteLine 3725
            Set m_ctlBottomRight(lngUpper) = ctl
        End If
vbwProfiler.vbwExecuteLine 3726 'B
    End If
vbwProfiler.vbwExecuteLine 3727 'B

vbwProfiler.vbwExecuteLine 3728
    m_imgSplitter.ZOrder
vbwProfiler.vbwProcOut 182
vbwProfiler.vbwExecuteLine 3729
End Sub

Public Sub RemoveAllControls()
vbwProfiler.vbwProcIn 183
    Dim i As Long
vbwProfiler.vbwExecuteLine 3730
    On Error GoTo err

vbwProfiler.vbwExecuteLine 3731
    For i = 0 To UBound(m_ctlTopLeft)
vbwProfiler.vbwExecuteLine 3732
        Set m_ctlTopLeft(i) = Nothing
vbwProfiler.vbwExecuteLine 3733
    Next

vbwProfiler.vbwExecuteLine 3734
    For i = 0 To UBound(m_ctlBottomRight)
vbwProfiler.vbwExecuteLine 3735
        Set m_ctlBottomRight(i) = Nothing
vbwProfiler.vbwExecuteLine 3736
    Next
vbwProfiler.vbwProcOut 183
vbwProfiler.vbwExecuteLine 3737
    Exit Sub
err:
vbwProfiler.vbwExecuteLine 3738
    Debug.Print "Error in cSplitter.RemoveAllControls() - " & err.Description
vbwProfiler.vbwExecuteLine 3739
    Resume Next

vbwProfiler.vbwProcOut 183
vbwProfiler.vbwExecuteLine 3740
End Sub
Private Sub Class_Initialize()
vbwProfiler.vbwProcIn 184

vbwProfiler.vbwExecuteLine 3741
    ReDim m_ctlTopLeft(0)
vbwProfiler.vbwExecuteLine 3742
    ReDim m_ctlBottomRight(0)
vbwProfiler.vbwExecuteLine 3743
    m_lngSplitterColor = vbWhite
vbwProfiler.vbwProcOut 184
vbwProfiler.vbwExecuteLine 3744
End Sub

Private Sub Class_Terminate()
vbwProfiler.vbwProcIn 185
    Dim i As Long
vbwProfiler.vbwExecuteLine 3745
    On Error GoTo err

vbwProfiler.vbwExecuteLine 3746
    Call RemoveAllControls

vbwProfiler.vbwExecuteLine 3747
    frmDesigner.Controls.Remove (m_imgSplitter.Tag)
vbwProfiler.vbwExecuteLine 3748
    Set m_imgSplitter = Nothing
vbwProfiler.vbwExecuteLine 3749
    Set m_frm = Nothing
vbwProfiler.vbwProcOut 185
vbwProfiler.vbwExecuteLine 3750
    Exit Sub

err:
vbwProfiler.vbwExecuteLine 3751
    Debug.Print "Error in cSplitter.Terminate() - " & err.Description
vbwProfiler.vbwExecuteLine 3752
    Resume Next
vbwProfiler.vbwProcOut 185
vbwProfiler.vbwExecuteLine 3753
End Sub

Private Sub ResizeSplitter()
vbwProfiler.vbwProcIn 186
vbwProfiler.vbwExecuteLine 3754
    Call MoveWindow(m_hWnd, m_lngLeft, m_lngTop, m_lngWidth, m_lngHeight, True)
vbwProfiler.vbwProcOut 186
vbwProfiler.vbwExecuteLine 3755
End Sub

Private Sub GetClientDimensions(ByRef oCtl As Object, ByRef lngLeft As Long, ByRef lngTop As Long, ByRef lngWidth As Long, ByRef lngHeight As Long)
vbwProfiler.vbwProcIn 187

   Dim r As Rect
   Dim pt As POINTAPI
   Dim P(0 To 1) As POINTAPI

vbwProfiler.vbwExecuteLine 3756
    On Error GoTo err

   ' store the original dimensions of the treeview
vbwProfiler.vbwExecuteLine 3757
    If TypeOf oCtl Is PropListCtl Then
vbwProfiler.vbwExecuteLine 3758
        GetWindowRect oCtl.Window, r  ' lame ass proplistctl uses .Window instead of .hwnd so test for it
    Else
vbwProfiler.vbwExecuteLine 3759 'B
vbwProfiler.vbwExecuteLine 3760
        GetWindowRect oCtl.hWnd, r
    End If
vbwProfiler.vbwExecuteLine 3761 'B

vbwProfiler.vbwExecuteLine 3762
    P(0).x = r.Left
vbwProfiler.vbwExecuteLine 3763
    P(0).y = r.Top
vbwProfiler.vbwExecuteLine 3764
    P(1).x = r.Right
vbwProfiler.vbwExecuteLine 3765
    P(1).y = r.Bottom

vbwProfiler.vbwExecuteLine 3766
    ScreenToClient m_frm.hWnd, P(0)
vbwProfiler.vbwExecuteLine 3767
    ScreenToClient m_frm.hWnd, P(1)
vbwProfiler.vbwExecuteLine 3768
    lngLeft = P(0).x
vbwProfiler.vbwExecuteLine 3769
    lngTop = P(0).y
vbwProfiler.vbwExecuteLine 3770
    lngHeight = P(1).y - P(0).y
vbwProfiler.vbwExecuteLine 3771
    lngWidth = P(1).x - P(0).x
vbwProfiler.vbwProcOut 187
vbwProfiler.vbwExecuteLine 3772
    Exit Sub

err:
vbwProfiler.vbwExecuteLine 3773
    Debug.Print "cSplitter:GetClientDimensions() -- " & err.Description
vbwProfiler.vbwProcOut 187
vbwProfiler.vbwExecuteLine 3774
    Exit Sub
vbwProfiler.vbwProcOut 187
vbwProfiler.vbwExecuteLine 3775
End Sub


Public Sub SetPosition(ByRef frm As Object, ByVal UniqueKey As String, ByVal Top As Long, ByVal Left As Long, ByVal Width As Long, ByVal Height As Long)
vbwProfiler.vbwProcIn 188

    ' create our Image control
vbwProfiler.vbwExecuteLine 3776
    If Not m_frm Is Nothing Then
vbwProfiler.vbwExecuteLine 3777
         Set m_frm = Nothing
    End If
vbwProfiler.vbwExecuteLine 3778 'B
vbwProfiler.vbwExecuteLine 3779
    Set m_frm = frm

vbwProfiler.vbwExecuteLine 3780
    If Not m_imgSplitter Is Nothing Then
vbwProfiler.vbwExecuteLine 3781
         Set m_imgSplitter = Nothing
    End If
vbwProfiler.vbwExecuteLine 3782 'B

vbwProfiler.vbwExecuteLine 3783
    Set m_imgSplitter = frmDesigner.Controls.Add("VB.PictureBox", UniqueKey, frm)
    ' set the hwnd property to the same as the image splitter's
vbwProfiler.vbwExecuteLine 3784
    m_hWnd = m_imgSplitter.hWnd

vbwProfiler.vbwExecuteLine 3785
    m_imgSplitter.Tag = UniqueKey
vbwProfiler.vbwExecuteLine 3786
    m_lngLeft = Left
vbwProfiler.vbwExecuteLine 3787
    m_lngTop = Top
vbwProfiler.vbwExecuteLine 3788
    m_lngWidth = Width
vbwProfiler.vbwExecuteLine 3789
    m_lngHeight = Height

vbwProfiler.vbwExecuteLine 3790
    m_imgSplitter.BorderStyle = 0 'none

vbwProfiler.vbwExecuteLine 3791
    Call ResizeSplitter
vbwProfiler.vbwExecuteLine 3792
    m_imgSplitter.Visible = True

vbwProfiler.vbwProcOut 188
vbwProfiler.vbwExecuteLine 3793
End Sub

Public Sub SetPadding(ByVal Top As Long, ByVal Left As Long, ByVal Right As Long, ByVal Bottom As Long)
vbwProfiler.vbwProcIn 189

vbwProfiler.vbwExecuteLine 3794
    m_lngTopPadding = Top
vbwProfiler.vbwExecuteLine 3795
    m_lngLeftPadding = Left
vbwProfiler.vbwExecuteLine 3796
    m_lngRightPadding = Right
vbwProfiler.vbwExecuteLine 3797
    m_lngBottomPadding = Bottom

vbwProfiler.vbwExecuteLine 3798
    Call ParentResized
vbwProfiler.vbwProcOut 189
vbwProfiler.vbwExecuteLine 3799
End Sub

Public Sub Move(ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngWidth As Long, ByVal lngHeight As Long)
vbwProfiler.vbwProcIn 190


vbwProfiler.vbwExecuteLine 3800
    m_lngTop = lngTop
vbwProfiler.vbwExecuteLine 3801
    m_lngLeft = lngLeft
vbwProfiler.vbwExecuteLine 3802
    m_lngWidth = lngWidth
vbwProfiler.vbwExecuteLine 3803
    m_lngHeight = lngHeight
vbwProfiler.vbwExecuteLine 3804
    Call ResizeSplitter

vbwProfiler.vbwProcOut 190
vbwProfiler.vbwExecuteLine 3805
End Sub

Private Sub m_imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 191

vbwProfiler.vbwExecuteLine 3806
    With m_imgSplitter
vbwProfiler.vbwExecuteLine 3807
        .BackColor = m_lngSplitterColor     ' Make the splitter visible
vbwProfiler.vbwExecuteLine 3808
        .ZOrder
vbwProfiler.vbwExecuteLine 3809
    End With
vbwProfiler.vbwProcOut 191
vbwProfiler.vbwExecuteLine 3810
End Sub

Private Sub m_imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 192

vbwProfiler.vbwExecuteLine 3811
    m_imgSplitter.BackColor = frmDesigner.BackColor   ' change the color back to normal

vbwProfiler.vbwProcOut 192
vbwProfiler.vbwExecuteLine 3812
End Sub



Private Sub m_imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 193

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

vbwProfiler.vbwExecuteLine 3813
    On Error Resume Next
vbwProfiler.vbwExecuteLine 3814
    lngCurrentLeft = m_lngLeft
vbwProfiler.vbwExecuteLine 3815
    lngCurrentTop = m_lngTop
vbwProfiler.vbwExecuteLine 3816
    lngCurrentHeight = m_lngHeight
vbwProfiler.vbwExecuteLine 3817
    lngCurrentWidth = m_lngWidth

vbwProfiler.vbwExecuteLine 3818
    If Button = vbLeftButton Then
vbwProfiler.vbwExecuteLine 3819
        Call GetCursorPos(lppt)
vbwProfiler.vbwExecuteLine 3820
        Call ScreenToClient(m_frm.hWnd, lppt)

        '01/07/02 - Using Pixels instead of Twips to overcome resize limitation of Tabbed dialog. Switching ALL
        ' resize instances to MoveWindow() for consistancy therefore converting to Twips is no longer necessary
        ' convert the x, y values to twips
        'lppt.y = lppt.y * Screen.TwipsPerPixelY
        'lppt.x = lppt.x * Screen.TwipsPerPixelX

        ' move the splitter
vbwProfiler.vbwExecuteLine 3821
        Select Case m_bOrientation
'vbwLine 3822:            Case splithorizontal
            Case IIf(vbwProfiler.vbwExecuteLine(3822), VBWPROFILER_EMPTY, _
        splithorizontal)
                ' check for Max and Min Y coordinate restrictions
vbwProfiler.vbwExecuteLine 3823
                If lppt.y > m_lngMaxTop Then
vbwProfiler.vbwExecuteLine 3824
                    m_lngTop = m_lngMaxTop
'vbwLine 3825:                ElseIf lppt.y < m_lngMinTop Then
                ElseIf vbwProfiler.vbwExecuteLine(3825) Or lppt.y < m_lngMinTop Then
vbwProfiler.vbwExecuteLine 3826
                    m_lngTop = m_lngMinTop
                Else
vbwProfiler.vbwExecuteLine 3827 'B
vbwProfiler.vbwExecuteLine 3828
                    m_lngTop = lppt.y
                End If
vbwProfiler.vbwExecuteLine 3829 'B

                ' calc the difference in movement from last position
vbwProfiler.vbwExecuteLine 3830
                lngYDiff = m_lngTop - lngCurrentTop
vbwProfiler.vbwExecuteLine 3831
                If lngYDiff = 0 Then ' short circuit rest of sub if possible
vbwProfiler.vbwProcOut 193
vbwProfiler.vbwExecuteLine 3832
                     Exit Sub
                End If
vbwProfiler.vbwExecuteLine 3833 'B

                ' adjust the HEIGHT for controls above the splitter
vbwProfiler.vbwExecuteLine 3834
                For i = 0 To UBound(m_ctlTopLeft)
vbwProfiler.vbwExecuteLine 3835
                    If Not m_ctlTopLeft(i) Is Nothing Then
vbwProfiler.vbwExecuteLine 3836
                        Call GetClientDimensions(m_ctlTopLeft(i), lngLeft, lngTop, lngWidth, lngHeight)
vbwProfiler.vbwExecuteLine 3837
                        lngTop = m_lngMinTop
vbwProfiler.vbwExecuteLine 3838
                        lngHeight = m_lngTop - lngTop
vbwProfiler.vbwExecuteLine 3839
                        Call MoveWindow(m_ctlTopLeft(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, True)
                    End If
vbwProfiler.vbwExecuteLine 3840 'B
vbwProfiler.vbwExecuteLine 3841
                Next

                ' adjust the HEIGHT and TOPS for controls below the splitter
vbwProfiler.vbwExecuteLine 3842
                For i = 0 To UBound(m_ctlBottomRight)
vbwProfiler.vbwExecuteLine 3843
                    If Not m_ctlBottomRight(i) Is Nothing Then
vbwProfiler.vbwExecuteLine 3844
                        Call GetClientDimensions(m_ctlBottomRight(i), lngLeft, lngTop, lngWidth, lngHeight)
vbwProfiler.vbwExecuteLine 3845
                        lngTop = m_lngTop + m_lngHeight
vbwProfiler.vbwExecuteLine 3846
                        lngHeight = m_lngMaxTop - lngTop
vbwProfiler.vbwExecuteLine 3847
                        Call MoveWindow(m_ctlBottomRight(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, True)
                    End If
vbwProfiler.vbwExecuteLine 3848 'B
vbwProfiler.vbwExecuteLine 3849
                Next

            ' split verticle
            Case Else
vbwProfiler.vbwExecuteLine 3850 'B

                ' check for Max and Min X coordinate restrictions
vbwProfiler.vbwExecuteLine 3851
                If lppt.x < m_lngMinLeft Then
vbwProfiler.vbwExecuteLine 3852
                    m_lngLeft = m_lngMinLeft
vbwProfiler.vbwProcOut 193
vbwProfiler.vbwExecuteLine 3853
                    Exit Sub
'vbwLine 3854:                ElseIf lppt.x > m_lngMaxLeft Then
                ElseIf vbwProfiler.vbwExecuteLine(3854) Or lppt.x > m_lngMaxLeft Then
vbwProfiler.vbwExecuteLine 3855
                    m_lngLeft = m_lngMaxLeft
vbwProfiler.vbwProcOut 193
vbwProfiler.vbwExecuteLine 3856
                    Exit Sub
                Else
vbwProfiler.vbwExecuteLine 3857 'B
vbwProfiler.vbwExecuteLine 3858
                    m_lngLeft = lppt.x
                End If
vbwProfiler.vbwExecuteLine 3859 'B

                ' calc the difference in movement from last position
vbwProfiler.vbwExecuteLine 3860
                lngXDiff = m_lngLeft - lngCurrentLeft
vbwProfiler.vbwExecuteLine 3861
                If lngXDiff = 0 Then ' short circuit rest of sub if possible
vbwProfiler.vbwProcOut 193
vbwProfiler.vbwExecuteLine 3862
                     Exit Sub
                End If
vbwProfiler.vbwExecuteLine 3863 'B

                ' adjust the WIDTHS for controls LEFT of the veritcle splitter
vbwProfiler.vbwExecuteLine 3864
                For i = 0 To UBound(m_ctlTopLeft)
vbwProfiler.vbwExecuteLine 3865
                     If Not m_ctlTopLeft(i) Is Nothing Then
vbwProfiler.vbwExecuteLine 3866
                        Call GetClientDimensions(m_ctlTopLeft(i), lngLeft, lngTop, lngWidth, lngHeight)
vbwProfiler.vbwExecuteLine 3867
                        lngLeft = m_lngMinLeft
vbwProfiler.vbwExecuteLine 3868
                        lngWidth = m_lngLeft - lngLeft
vbwProfiler.vbwExecuteLine 3869
                        If TypeOf m_ctlTopLeft(i) Is TreeX Then
vbwProfiler.vbwExecuteLine 3870
                            Debug.Print m_ctlTopLeft(i).Width & " TREEX BEFORE"
                        End If
vbwProfiler.vbwExecuteLine 3871 'B

vbwProfiler.vbwExecuteLine 3872
                        If TypeOf m_ctlTopLeft(i) Is TreeX Then
vbwProfiler.vbwExecuteLine 3873
                            With m_ctlTopLeft(i)
vbwProfiler.vbwExecuteLine 3874
                                .Left = lngLeft * Screen.TwipsPerPixelX
                               ' .Top = lngTop * Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 3875
                                .Width = lngWidth * Screen.TwipsPerPixelX
                                '.Height = lngHeight * Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 3876
                            End With
                        Else
vbwProfiler.vbwExecuteLine 3877 'B
vbwProfiler.vbwExecuteLine 3878
                            Call MoveWindow(m_ctlTopLeft(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, True)
                        End If
vbwProfiler.vbwExecuteLine 3879 'B
                     End If
vbwProfiler.vbwExecuteLine 3880 'B
vbwProfiler.vbwExecuteLine 3881
                Next
                ' adjust the LEFTS and WIDTHS for controls RIGHT of the verticle splitter
vbwProfiler.vbwExecuteLine 3882
                For i = 0 To UBound(m_ctlBottomRight)
vbwProfiler.vbwExecuteLine 3883
                    If Not m_ctlBottomRight(i) Is Nothing Then
vbwProfiler.vbwExecuteLine 3884
                        Call GetClientDimensions(m_ctlBottomRight(i), lngLeft, lngTop, lngWidth, lngHeight)
vbwProfiler.vbwExecuteLine 3885
                        lngLeft = m_lngLeft + m_lngWidth
vbwProfiler.vbwExecuteLine 3886
                        lngWidth = m_lngMaxLeft - m_lngLeft

                        ' if this is the ONLY control on the right side, we can adjust the height value as well
vbwProfiler.vbwExecuteLine 3887
                        If UBound(m_ctlBottomRight) = 0 Then
vbwProfiler.vbwExecuteLine 3888
                            lngHeight = m_lngHeight
                        End If
vbwProfiler.vbwExecuteLine 3889 'B

vbwProfiler.vbwExecuteLine 3890
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
vbwProfiler.vbwExecuteLine 3891
                            With m_ctlBottomRight(i)
vbwProfiler.vbwExecuteLine 3892
                                .Left = lngLeft * Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 3893
                                .Width = lngWidth * Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 3894
                                .Top = lngTop * Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 3895
                                .Height = lngHeight * Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 3896
                            End With
                        Else
vbwProfiler.vbwExecuteLine 3897 'B
vbwProfiler.vbwExecuteLine 3898
                            Call MoveWindow(m_ctlBottomRight(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, True)
                        End If
vbwProfiler.vbwExecuteLine 3899 'B
vbwProfiler.vbwExecuteLine 3900
                        If TypeOf m_ctlBottomRight(i) Is cSplitter Then
                            ' when calling MoveWindow on this control, it only updates the PicBox and not the actual internally stored  dimensions of this cSplitter class.
                            ' so we (for now) are explicitly setting the proper internal dimensions of the control from here
vbwProfiler.vbwExecuteLine 3901
                            With m_ctlBottomRight(i)
vbwProfiler.vbwExecuteLine 3902
                                .Left = lngLeft
vbwProfiler.vbwExecuteLine 3903
                                .Top = lngTop
vbwProfiler.vbwExecuteLine 3904
                                .Height = lngHeight
vbwProfiler.vbwExecuteLine 3905
                                .Width = lngWidth
vbwProfiler.vbwExecuteLine 3906
                            End With
                        End If
vbwProfiler.vbwExecuteLine 3907 'B
                    End If
vbwProfiler.vbwExecuteLine 3908 'B
vbwProfiler.vbwExecuteLine 3909
                Next
        End Select
vbwProfiler.vbwExecuteLine 3910 'B

        ' move the splitter itself
vbwProfiler.vbwExecuteLine 3911
         Call Move(m_lngLeft, m_lngTop, m_lngWidth, m_lngHeight)
vbwProfiler.vbwExecuteLine 3912
         RaiseEvent Resize
    End If
vbwProfiler.vbwExecuteLine 3913 'B

vbwProfiler.vbwProcOut 193
vbwProfiler.vbwExecuteLine 3914
End Sub

Public Sub ParentResized()
' when the parent window resizes, the client area that the splitter must be bound by, should be updated.
' these values are stored in the RECT m_rectClient
vbwProfiler.vbwProcIn 194

' min max boundaries need to be updated


' Next, the size of the splitter needs to be updated

' All child controls need to be resized (Do i call ResizeControls or a seperate sub for this?)


Dim ret As Long
vbwProfiler.vbwExecuteLine 3915
On Error Resume Next
vbwProfiler.vbwExecuteLine 3916
ret = GetClientRect(m_frm.hWnd, m_rectClient)

' 01/07/02 - No longer using TWIPS since switching resize method to MoveWindow which uses pixels only
' Note: MoveWindow API is being used instead because it overcomes the Resize limits of the Tab Dialog control
'm_rectClient.Left = m_rectClient.Left * Screen.TwipsPerPixelX
'm_rectClient.Right = m_rectClient.Right * Screen.TwipsPerPixelX
'm_rectClient.Top = m_rectClient.Top * Screen.TwipsPerPixelY
'm_rectClient.Bottom = m_rectClient.Bottom * Screen.TwipsPerPixelY

vbwProfiler.vbwExecuteLine 3917
m_lngMinTop = m_rectClient.Top + m_lngTopPadding
vbwProfiler.vbwExecuteLine 3918
m_lngMaxTop = m_rectClient.Bottom - m_lngBottomPadding
vbwProfiler.vbwExecuteLine 3919
m_lngMaxLeft = m_rectClient.Right - m_lngRightPadding
vbwProfiler.vbwExecuteLine 3920
m_lngMinLeft = m_rectClient.Left + m_lngLeftPadding

vbwProfiler.vbwExecuteLine 3921
If m_bOrientation = splithorizontal Then
vbwProfiler.vbwExecuteLine 3922
    Width = m_rectClient.Right - m_lngLeft

Else
vbwProfiler.vbwExecuteLine 3923 'B
vbwProfiler.vbwExecuteLine 3924
    Top = m_lngMinTop
vbwProfiler.vbwExecuteLine 3925
    Height = m_lngMaxTop - m_lngMinTop 'm_rectClient.Bottom - m_lngTop
End If
vbwProfiler.vbwExecuteLine 3926 'B

vbwProfiler.vbwExecuteLine 3927
Call ResizeControls


vbwProfiler.vbwProcOut 194
vbwProfiler.vbwExecuteLine 3928
End Sub




Public Sub ResizeControls()
vbwProfiler.vbwProcIn 195

    Dim lngCurrentLeft As Long
    Dim lngCurrentTop As Long
    Dim i As Long
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngHeight As Long
    Dim lngWidth As Long
    Dim lngUpper As Long
    Dim rControl As Rect

vbwProfiler.vbwExecuteLine 3929
    On Error Resume Next
vbwProfiler.vbwExecuteLine 3930
    If m_bOrientation = splithorizontal Then
vbwProfiler.vbwExecuteLine 3931
        lngUpper = UBound(m_ctlTopLeft)
vbwProfiler.vbwExecuteLine 3932
        For i = 0 To lngUpper
vbwProfiler.vbwExecuteLine 3933
            If Not m_ctlTopLeft(i) Is Nothing Then
vbwProfiler.vbwExecuteLine 3934
                Call GetClientDimensions(m_ctlTopLeft(i), lngLeft, lngTop, lngWidth, lngHeight)

vbwProfiler.vbwExecuteLine 3935
                lngLeft = m_lngLeft
vbwProfiler.vbwExecuteLine 3936
                lngTop = m_lngMinTop
vbwProfiler.vbwExecuteLine 3937
                lngHeight = m_lngTop - lngTop
vbwProfiler.vbwExecuteLine 3938
                lngWidth = m_rectClient.Right - lngLeft - m_lngRightPadding

vbwProfiler.vbwExecuteLine 3939
                Call MoveWindow(m_ctlTopLeft(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, 1)
            End If
vbwProfiler.vbwExecuteLine 3940 'B
vbwProfiler.vbwExecuteLine 3941
        Next
vbwProfiler.vbwExecuteLine 3942
        lngUpper = UBound(m_ctlBottomRight)
vbwProfiler.vbwExecuteLine 3943
        For i = 0 To lngUpper
vbwProfiler.vbwExecuteLine 3944
            If Not m_ctlBottomRight(i) Is Nothing Then
vbwProfiler.vbwExecuteLine 3945
                Call GetClientDimensions(m_ctlBottomRight(i), lngLeft, lngTop, lngWidth, lngHeight)
vbwProfiler.vbwExecuteLine 3946
                lngLeft = m_lngLeft
vbwProfiler.vbwExecuteLine 3947
                lngTop = m_lngTop + m_lngHeight
vbwProfiler.vbwExecuteLine 3948
                lngWidth = m_rectClient.Right - lngLeft - m_lngRightPadding

vbwProfiler.vbwExecuteLine 3949
                If i = lngUpper Then
vbwProfiler.vbwExecuteLine 3950
                    lngHeight = m_rectClient.Bottom - lngTop - m_lngBottomPadding
                End If
vbwProfiler.vbwExecuteLine 3951 'B

vbwProfiler.vbwExecuteLine 3952
                Call MoveWindow(m_ctlBottomRight(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, 1)
            End If
vbwProfiler.vbwExecuteLine 3953 'B
vbwProfiler.vbwExecuteLine 3954
        Next
    Else ' verticle
vbwProfiler.vbwExecuteLine 3955 'B
vbwProfiler.vbwExecuteLine 3956
        lngUpper = UBound(m_ctlTopLeft)
vbwProfiler.vbwExecuteLine 3957
        For i = 0 To lngUpper
vbwProfiler.vbwExecuteLine 3958
            If Not m_ctlTopLeft(i) Is Nothing Then

vbwProfiler.vbwExecuteLine 3959
                lngLeft = m_lngMinLeft
vbwProfiler.vbwExecuteLine 3960
                lngTop = m_lngMinTop

vbwProfiler.vbwExecuteLine 3961
                lngHeight = m_lngHeight
vbwProfiler.vbwExecuteLine 3962
                lngWidth = m_lngLeft - lngLeft
vbwProfiler.vbwExecuteLine 3963
                If TypeOf m_ctlTopLeft(i) Is TreeX Then
vbwProfiler.vbwExecuteLine 3964
                    With m_ctlTopLeft(i)
vbwProfiler.vbwExecuteLine 3965
                        .Left = lngLeft * Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 3966
                        .Top = lngTop * Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 3967
                        .Width = lngWidth * Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 3968
                        .Height = lngHeight * Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 3969
                    End With
                Else
vbwProfiler.vbwExecuteLine 3970 'B
vbwProfiler.vbwExecuteLine 3971
                    Call MoveWindow(m_ctlTopLeft(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, 1)
                End If
vbwProfiler.vbwExecuteLine 3972 'B
            End If
vbwProfiler.vbwExecuteLine 3973 'B

vbwProfiler.vbwExecuteLine 3974
        Next

vbwProfiler.vbwExecuteLine 3975
        lngUpper = UBound(m_ctlBottomRight)
vbwProfiler.vbwExecuteLine 3976
        If m_ctlBottomRight(lngUpper) Is Nothing Then
vbwProfiler.vbwExecuteLine 3977
           lngUpper = lngUpper - 1
        End If
vbwProfiler.vbwExecuteLine 3978 'B
vbwProfiler.vbwExecuteLine 3979
        For i = 0 To lngUpper
vbwProfiler.vbwExecuteLine 3980
            If Not m_ctlBottomRight(i) Is Nothing Then
vbwProfiler.vbwExecuteLine 3981
                Call GetClientDimensions(m_ctlBottomRight(i), lngLeft, lngTop, lngWidth, lngHeight)
vbwProfiler.vbwExecuteLine 3982
                lngLeft = m_lngLeft + m_lngWidth
vbwProfiler.vbwExecuteLine 3983
                If i = 0 Then
vbwProfiler.vbwExecuteLine 3984
                     lngTop = m_lngMinTop
                End If
vbwProfiler.vbwExecuteLine 3985 'B

               ' ' never change the height of a child Splitter that's horizontal
vbwProfiler.vbwExecuteLine 3986
                If TypeOf m_ctlBottomRight(i) Is cSplitter Then
vbwProfiler.vbwExecuteLine 3987
                    lngHeight = m_ctlBottomRight(i).Height
                Else
vbwProfiler.vbwExecuteLine 3988 'B
vbwProfiler.vbwExecuteLine 3989
                    If i = lngUpper Then                   'if last control, its height can change
vbwProfiler.vbwExecuteLine 3990
                        lngHeight = m_lngMaxTop - lngTop    ' child control's height is same as the splitters height
                    End If
vbwProfiler.vbwExecuteLine 3991 'B
                End If
vbwProfiler.vbwExecuteLine 3992 'B

                'lngWidth = m_frm.Width - m_lngLeft - m_lngRightPadding - m_lngWidth  '          'm_ctlTopLeft(i).Width
vbwProfiler.vbwExecuteLine 3993
                lngWidth = m_lngMaxLeft - (m_lngLeft + m_lngWidth)

vbwProfiler.vbwExecuteLine 3994
                If TypeOf m_ctlBottomRight(i) Is PropListCtl Then
                    'Call MoveWindow(m_ctlBottomRight(i).Window, lngLeft, lngTop, lngWidth, lngHeight, 1)
vbwProfiler.vbwExecuteLine 3995
                    With m_ctlBottomRight(i)
vbwProfiler.vbwExecuteLine 3996
                        .Left = lngLeft * Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 3997
                        .Width = lngWidth * Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 3998
                        .Top = lngTop * Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 3999
                        .Height = lngHeight * Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 4000
                    End With
                Else
vbwProfiler.vbwExecuteLine 4001 'B
vbwProfiler.vbwExecuteLine 4002
                    Call MoveWindow(m_ctlBottomRight(i).hWnd, lngLeft, lngTop, lngWidth, lngHeight, 1)
                End If
vbwProfiler.vbwExecuteLine 4003 'B
            End If
vbwProfiler.vbwExecuteLine 4004 'B
vbwProfiler.vbwExecuteLine 4005
        Next
    End If
vbwProfiler.vbwExecuteLine 4006 'B

vbwProfiler.vbwExecuteLine 4007
    RaiseEvent Resize

vbwProfiler.vbwProcOut 195
vbwProfiler.vbwExecuteLine 4008
End Sub


