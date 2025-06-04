VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCustomListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' basic custom listbox which tries to emulate the Outlook 2000 rules wizard rule edit listbox.
' lots of limitations for general purpose use which is why this is NOT a general purpose class.
' its strictly for our unit convertor
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const SRCCOPY = &HCC0020
Private Const TRANSPARENT = 1


Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Private Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Private Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
' Brush Styles
Private Const BS_SOLID = 0
Private Const BS_NULL = 1
Private Const BS_HOLLOW = BS_NULL
Private Const BS_HATCHED = 2
Private Const BS_PATTERN = 3
Private Const BS_INDEXED = 4
Private Const BS_DIBPATTERN = 5
Private Const BS_DIBPATTERNPT = 6
Private Const BS_PATTERN8X8 = 7
Private Const BS_DIBPATTERN8X8 = 8

'  Hatch Styles
Private Const HS_HORIZONTAL = 0              '  -----
Private Const HS_VERTICAL = 1                '  |||||
Private Const HS_FDIAGONAL = 2               '  \\\\\
Private Const HS_BDIAGONAL = 3               '  /////
Private Const HS_CROSS = 4                   '  +++++
Private Const HS_DIAGCROSS = 5               '  xxxxx
Private Const HS_FDIAGONAL1 = 6
Private Const HS_BDIAGONAL1 = 7
Private Const HS_SOLID = 8
Private Const HS_DENSE1 = 9
Private Const HS_DENSE2 = 10
Private Const HS_DENSE3 = 11
Private Const HS_DENSE4 = 12
Private Const HS_DENSE5 = 13
Private Const HS_DENSE6 = 14
Private Const HS_DENSE7 = 15
Private Const HS_DENSE8 = 16
Private Const HS_NOSHADE = 17
Private Const HS_HALFTONE = 18
Private Const HS_SOLIDCLR = 19
Private Const HS_DITHEREDCLR = 20
Private Const HS_SOLIDTEXTCLR = 21
Private Const HS_DITHEREDTEXTCLR = 22
Private Const HS_SOLIDBKCLR = 23
Private Const HS_DITHEREDBKCLR = 24
Private Const HS_API_MAX = 25

Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As String * 1
    lfUnderline As String * 1
    lfStrikeout As String * 1
    lfCharSet As String * 1
    lfOutPrecision As String * 1
    lfClipPrecision As String * 1
    lfQuality As String * 1
    lfPitchAndFamily As String * 1
    lfFaceName As String * 24
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

Private m_sRows() As String
Private m_lngItemData() As Long

Private m_rSpots() As RECT
Private m_lngCursorX As Long
Private m_lngCursorY As Long
Private m_lngScrollPosition As Long '//this gets passed in from scroller
Private m_lngNumRows As Long
Private m_lngMaxVisible As Long
Private m_lngRowLength As Long  '//length of a row in pixels
Private m_lngRowHeight As Long  '//height of a row in pixels
Private m_lngRowSeperation As Long
Private m_lngRowLeftPadding As Long
Private m_lngRowRightPadding As Long

Private m_lngLeft As Long
Private m_lngTop As Long
Private m_lngWidth As Long
Private m_lngHeight As Long

Private m_hFont As Long
Private m_hOldFont As Long
Private m_hBmp As Long
Private m_hOldBmp As Long
Private m_hDisplayHwnd As Long
Private m_hDisplayDC As Long
Private m_hBackBufferDC As Long
Private m_lngBkMode As Long
Private m_lngOldBkMode As Long
Private m_lngForeColor As Long
Private m_lngOldForeColor As Long
Private m_lngBackColor As Long
Private m_lngOldBackColor As Long
Private m_hBrush As Long
Private m_hOldBrush As Long
Private m_bInitialized As Boolean

Event ItemAdded(ByVal lngItemCount As Long, ByVal lngMaxVisible As Long)

Private Sub Class_Initialize()
    '//default values
vbwProfiler.vbwProcIn 513
vbwProfiler.vbwExecuteLine 10798
    m_lngRowLeftPadding = 3 'pixels
vbwProfiler.vbwExecuteLine 10799
    m_lngRowRightPadding = 3
vbwProfiler.vbwExecuteLine 10800
    m_lngRowSeperation = 2
vbwProfiler.vbwExecuteLine 10801
    m_lngScrollPosition = 1
vbwProfiler.vbwProcOut 513
vbwProfiler.vbwExecuteLine 10802
End Sub

Private Sub Class_Terminate()
    ' before we delete anything, we must restore them to their original states to keep VB happy
vbwProfiler.vbwProcIn 514
vbwProfiler.vbwExecuteLine 10803
    SelectObject m_hBackBufferDC, m_hOldBmp
vbwProfiler.vbwExecuteLine 10804
    SelectObject m_hBackBufferDC, m_hOldFont
vbwProfiler.vbwExecuteLine 10805
    SetBkMode m_hBackBufferDC, m_lngOldBkMode
vbwProfiler.vbwExecuteLine 10806
    SetTextColor m_hBackBufferDC, m_lngOldForeColor

vbwProfiler.vbwExecuteLine 10807
    SelectObject m_hBackBufferDC, m_hOldBrush
vbwProfiler.vbwExecuteLine 10808
    DeleteObject m_hBrush
vbwProfiler.vbwExecuteLine 10809
    DeleteObject m_hBmp
    ' DeleteObject m_hFont ' we dont delete this font since it belongs to the picturebox DC
vbwProfiler.vbwExecuteLine 10810
    DeleteDC m_hBackBufferDC
vbwProfiler.vbwProcOut 514
vbwProfiler.vbwExecuteLine 10811
End Sub


Public Sub initDisplay(ByVal hwnd As Long, ByVal lngBackColor As Long, ByVal lngTextColor As Long)
    ' todo: here we should select the font into our offscreen buffer?
vbwProfiler.vbwProcIn 515
    Dim lf As LOGFONT
    Dim hFont As Long
    Dim r As RECT
    Dim tm As TEXTMETRIC

    ' create a dummy font so that we can select "out" a handle to the font of the hdc passed in
vbwProfiler.vbwExecuteLine 10812
    hFont = CreateFontIndirect(lf)
vbwProfiler.vbwExecuteLine 10813
    m_hDisplayHwnd = hwnd
vbwProfiler.vbwExecuteLine 10814
    m_hDisplayDC = GetDC(hwnd)
vbwProfiler.vbwExecuteLine 10815
    m_hFont = SelectObject(m_hDisplayDC, hFont)
vbwProfiler.vbwExecuteLine 10816
    SelectObject m_hDisplayDC, m_hFont '<-- select the original back in now that we have a handle to it
vbwProfiler.vbwExecuteLine 10817
    DeleteObject hFont ' and delete the dummy font since we no longer need it

    ' create our backbuffer and select the same font as our picturebox into it
vbwProfiler.vbwExecuteLine 10818
    m_hBackBufferDC = CreateCompatibleDC(m_hDisplayDC)
vbwProfiler.vbwExecuteLine 10819
    m_hOldFont = SelectObject(m_hBackBufferDC, m_hFont)

    ' create a bitmap of the same dimensions as our picturebox to act as our backbuffer
vbwProfiler.vbwExecuteLine 10820
    GetClientRect hwnd, r
vbwProfiler.vbwExecuteLine 10821
    m_lngLeft = r.Left
vbwProfiler.vbwExecuteLine 10822
    m_lngTop = r.Top
vbwProfiler.vbwExecuteLine 10823
    m_lngWidth = r.Right
vbwProfiler.vbwExecuteLine 10824
    m_lngHeight = r.Bottom

vbwProfiler.vbwExecuteLine 10825
    m_hBmp = CreateCompatibleBitmap(m_hDisplayDC, r.Right, r.Bottom)
vbwProfiler.vbwExecuteLine 10826
    m_hOldBmp = SelectObject(m_hBackBufferDC, m_hBmp)

    ' get our row height
vbwProfiler.vbwExecuteLine 10827
    GetTextMetrics m_hDisplayDC, tm
vbwProfiler.vbwExecuteLine 10828
    m_lngRowHeight = tm.tmHeight
vbwProfiler.vbwExecuteLine 10829
    m_lngMaxVisible = m_lngHeight \ m_lngRowHeight
vbwProfiler.vbwExecuteLine 10830
    m_bInitialized = True

vbwProfiler.vbwExecuteLine 10831
    BackColor = lngBackColor
vbwProfiler.vbwExecuteLine 10832
    TextColor = lngTextColor
vbwProfiler.vbwProcOut 515
vbwProfiler.vbwExecuteLine 10833
End Sub

Public Sub setItemData(ByVal lngIndex As Long, ByVal lngValue As Long)
vbwProfiler.vbwProcIn 516
vbwProfiler.vbwExecuteLine 10834
    m_lngItemData(lngIndex) = lngValue
vbwProfiler.vbwProcOut 516
vbwProfiler.vbwExecuteLine 10835
End Sub
Public Function getItemData(ByVal lngIndex As Long) As Long
vbwProfiler.vbwProcIn 517
vbwProfiler.vbwExecuteLine 10836
    getItemData = m_lngItemData(lngIndex)
vbwProfiler.vbwProcOut 517
vbwProfiler.vbwExecuteLine 10837
End Function
Public Function Addline(ByVal sLine As String, ByVal bRenderImmediately As Boolean) As Long
vbwProfiler.vbwProcIn 518
vbwProfiler.vbwExecuteLine 10838
    If m_bInitialized Then
vbwProfiler.vbwExecuteLine 10839
        m_lngNumRows = m_lngNumRows + 1
vbwProfiler.vbwExecuteLine 10840
        ReDim Preserve m_sRows(1 To m_lngNumRows)
vbwProfiler.vbwExecuteLine 10841
        ReDim Preserve m_lngItemData(1 To m_lngNumRows)
vbwProfiler.vbwExecuteLine 10842
        ReDim Preserve m_rSpots(1 To m_lngNumRows)
vbwProfiler.vbwExecuteLine 10843
        m_sRows(m_lngNumRows) = sLine

        '//now that our line widths are properly setup, we can print the visible lines
vbwProfiler.vbwExecuteLine 10844
        If bRenderImmediately Then
vbwProfiler.vbwExecuteLine 10845
            Call RenderText
        End If
vbwProfiler.vbwExecuteLine 10846 'B
    End If
vbwProfiler.vbwExecuteLine 10847 'B
    'note: this returns the last row added index and NOT the number of rows.
    '      it only just so happens that this implementation the NumRows is the same
    '      as last added since we append it to the end of the array.  But if were to do
    '      some sort of "insert" add, we must return the index of the actual item inserted
vbwProfiler.vbwExecuteLine 10848
    RaiseEvent ItemAdded(m_lngNumRows, m_lngMaxVisible)
vbwProfiler.vbwExecuteLine 10849
    Addline = m_lngNumRows
vbwProfiler.vbwProcOut 518
vbwProfiler.vbwExecuteLine 10850
End Function

Public Sub RenderText()
    '//take the formatted text and print all visible rows
vbwProfiler.vbwProcIn 519
    Dim i As Long
    Dim j As Long
    Dim k As Long

    Dim pt As POINTAPI
    Dim iCursor As Long
    Dim lngLength As Long
    Dim xPos As Long
    Dim yPos As Long
    Dim s As String
    Dim sizeText As SIZE

vbwProfiler.vbwExecuteLine 10851
    Call Clear
vbwProfiler.vbwExecuteLine 10852
    yPos = 1
vbwProfiler.vbwExecuteLine 10853
    xPos = m_lngRowLeftPadding
vbwProfiler.vbwExecuteLine 10854
    iCursor = 1
vbwProfiler.vbwExecuteLine 10855
    For i = m_lngScrollPosition To m_lngNumRows
vbwProfiler.vbwExecuteLine 10856
        Do
vbwProfiler.vbwExecuteLine 10857
            s = m_sRows(i)
vbwProfiler.vbwExecuteLine 10858
            j = InStr(iCursor, s, "<u>")

vbwProfiler.vbwExecuteLine 10859
            If j Then
                ' draw the text up to the point of the <u>
vbwProfiler.vbwExecuteLine 10860
                lngLength = j - iCursor
vbwProfiler.vbwExecuteLine 10861
                s = Mid(m_sRows(i), iCursor, lngLength)
vbwProfiler.vbwExecuteLine 10862
                TextOut m_hBackBufferDC, xPos, yPos, s, Len(s)
vbwProfiler.vbwExecuteLine 10863
                iCursor = j + Len("<u>")
vbwProfiler.vbwExecuteLine 10864
                GetTextExtentPoint32 m_hBackBufferDC, s, Len(s), sizeText
vbwProfiler.vbwExecuteLine 10865
                xPos = xPos + sizeText.cx
vbwProfiler.vbwExecuteLine 10866
                k = InStr(iCursor, m_sRows(i), "</u>")
vbwProfiler.vbwExecuteLine 10867
                If k Then
                    ' draw text up to the close </u>
vbwProfiler.vbwExecuteLine 10868
                    lngLength = k - iCursor
vbwProfiler.vbwExecuteLine 10869
                    s = Mid(m_sRows(i), iCursor, lngLength)
vbwProfiler.vbwExecuteLine 10870
                    TextOut m_hBackBufferDC, xPos, yPos, s, Len(s)
vbwProfiler.vbwExecuteLine 10871
                    iCursor = k + Len("</u>")

                    ' underline it and register the hotspot
vbwProfiler.vbwExecuteLine 10872
                    MoveToEx m_hBackBufferDC, xPos, yPos + m_lngRowHeight, pt
vbwProfiler.vbwExecuteLine 10873
                    m_rSpots(i).Left = xPos
vbwProfiler.vbwExecuteLine 10874
                    m_rSpots(i).Right = yPos
vbwProfiler.vbwExecuteLine 10875
                    GetTextExtentPoint32 m_hBackBufferDC, s, Len(s), sizeText
vbwProfiler.vbwExecuteLine 10876
                    xPos = xPos + sizeText.cx
vbwProfiler.vbwExecuteLine 10877
                    m_rSpots(i).Right = xPos
vbwProfiler.vbwExecuteLine 10878
                    m_rSpots(i).Bottom = yPos + m_lngRowHeight
                    ' todo: should select in a diff color pen here... our underline is coming up as pen color and not as forecolor
                    ' plus maybe we want our underline a diff color than even the forecolor
vbwProfiler.vbwExecuteLine 10879
                    LineTo m_hBackBufferDC, xPos, yPos + m_lngRowHeight
                End If
vbwProfiler.vbwExecuteLine 10880 'B
            Else
vbwProfiler.vbwExecuteLine 10881 'B
                ' no more underline control characters, draw the rest of the line in one swoop
vbwProfiler.vbwExecuteLine 10882
                lngLength = Len(m_sRows(i)) - iCursor + 1
vbwProfiler.vbwExecuteLine 10883
                s = Mid(m_sRows(i), iCursor, lngLength)
vbwProfiler.vbwExecuteLine 10884
                iCursor = Len(m_sRows(i))
vbwProfiler.vbwExecuteLine 10885
                TextOut m_hBackBufferDC, xPos, yPos, s, Len(s)
            End If
vbwProfiler.vbwExecuteLine 10886 'B
vbwProfiler.vbwExecuteLine 10887
            If iCursor >= Len(m_sRows(i)) Then
vbwProfiler.vbwExecuteLine 10888
                 Exit Do
            End If
vbwProfiler.vbwExecuteLine 10889 'B
vbwProfiler.vbwExecuteLine 10890
        Loop
vbwProfiler.vbwExecuteLine 10891
        yPos = yPos + m_lngRowHeight '+ m_lngRowSeperation
vbwProfiler.vbwExecuteLine 10892
        xPos = m_lngRowLeftPadding
vbwProfiler.vbwExecuteLine 10893
        iCursor = 1
vbwProfiler.vbwExecuteLine 10894
    Next
vbwProfiler.vbwExecuteLine 10895
    Call Paint

vbwProfiler.vbwProcOut 519
vbwProfiler.vbwExecuteLine 10896
End Sub

Public Sub Paint()
vbwProfiler.vbwProcIn 520
vbwProfiler.vbwExecuteLine 10897
    If m_bInitialized Then
vbwProfiler.vbwExecuteLine 10898
        BitBlt m_hDisplayDC, 0, 0, m_lngWidth, m_lngHeight, m_hBackBufferDC, 0, 0, SRCCOPY
    End If
vbwProfiler.vbwExecuteLine 10899 'B
vbwProfiler.vbwProcOut 520
vbwProfiler.vbwExecuteLine 10900
End Sub

Public Sub RemoveAllItems()
     '//clears the entire buffer of text
vbwProfiler.vbwProcIn 521
vbwProfiler.vbwExecuteLine 10901
    Erase m_sRows()
vbwProfiler.vbwExecuteLine 10902
    Erase m_lngItemData()
vbwProfiler.vbwExecuteLine 10903
    m_lngNumRows = 0
vbwProfiler.vbwExecuteLine 10904
    m_lngCursorY = 0
vbwProfiler.vbwProcOut 521
vbwProfiler.vbwExecuteLine 10905
End Sub

Public Sub Clear()
vbwProfiler.vbwProcIn 522
vbwProfiler.vbwExecuteLine 10906
    PatBlt m_hBackBufferDC, 0, 0, m_lngWidth, m_lngHeight, PATCOPY
vbwProfiler.vbwProcOut 522
vbwProfiler.vbwExecuteLine 10907
End Sub

Public Function PointInHotSpot(ByVal x As Long, y As Long) As Long
    ' returns 0 if no, index value of row item if true
    ' NOTE: our hotspots are one per row... that is the limitation.
    ' So our array of hotspots is also mirrroed to our rows (e.g. even if there
    ' is no hotspot, its index exists with a rect of 0,0,0,0
vbwProfiler.vbwProcIn 523
    Dim i As Long
vbwProfiler.vbwExecuteLine 10908
    For i = 1 To m_lngNumRows
vbwProfiler.vbwExecuteLine 10909
        If PtInRect(m_rSpots(i), x, y) Then
vbwProfiler.vbwExecuteLine 10910
            PointInHotSpot = i
vbwProfiler.vbwProcOut 523
vbwProfiler.vbwExecuteLine 10911
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 10912 'B
vbwProfiler.vbwExecuteLine 10913
    Next
    'we're still here, it didnt exist
vbwProfiler.vbwExecuteLine 10914
    PointInHotSpot = 0
vbwProfiler.vbwProcOut 523
vbwProfiler.vbwExecuteLine 10915
End Function

Public Property Let BackColor(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 524
    Dim lRet As Long
    Dim LB As LOGBRUSH

vbwProfiler.vbwExecuteLine 10916
    If m_bInitialized Then
        ' back mode must be set to transparent
vbwProfiler.vbwExecuteLine 10917
        SetBkMode m_hBackBufferDC, TRANSPARENT
        ' create our flood filling brush which we use to clear the backbuffer
vbwProfiler.vbwExecuteLine 10918
        LB.lbColor = lngValue
vbwProfiler.vbwExecuteLine 10919
        LB.lbStyle = BS_SOLID
vbwProfiler.vbwExecuteLine 10920
        LB.lbHatch = HS_SOLIDBKCLR

        '//create the brush using those settings and select it into the dc
vbwProfiler.vbwExecuteLine 10921
        If m_hOldBrush = 0 Then
vbwProfiler.vbwExecuteLine 10922
            m_hBrush = CreateBrushIndirect(LB)
vbwProfiler.vbwExecuteLine 10923
            m_hOldBrush = SelectObject(m_hBackBufferDC, m_hBrush)
        Else
vbwProfiler.vbwExecuteLine 10924 'B
            ' if we've already created a brush the first time, we already have a handle
            ' to the original brush so delete the prvious "new" brush and create a (cough) new, new one
            ' and select it into our dc WITHOUT saving it over the m_hOldBrush
vbwProfiler.vbwExecuteLine 10925
            DeleteObject m_hBrush
vbwProfiler.vbwExecuteLine 10926
            m_hBrush = CreateBrushIndirect(LB)
vbwProfiler.vbwExecuteLine 10927
            lRet = SelectObject(m_hBackBufferDC, m_hBrush)
        End If
vbwProfiler.vbwExecuteLine 10928 'B

vbwProfiler.vbwExecuteLine 10929
        m_lngBackColor = lngValue

        ' back color has changed, we need to re-render the text
vbwProfiler.vbwExecuteLine 10930
        RenderText
    End If
vbwProfiler.vbwExecuteLine 10931 'B
vbwProfiler.vbwProcOut 524
vbwProfiler.vbwExecuteLine 10932
End Property

Public Property Let TextColor(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 525
    Dim lRet As Long
vbwProfiler.vbwExecuteLine 10933
    If m_bInitialized Then
vbwProfiler.vbwExecuteLine 10934
        m_lngForeColor = lngValue
vbwProfiler.vbwExecuteLine 10935
        lRet = SetTextColor(m_hBackBufferDC, m_lngForeColor)
vbwProfiler.vbwExecuteLine 10936
        If m_lngOldForeColor = 0 Then
vbwProfiler.vbwExecuteLine 10937
            m_lngOldForeColor = lRet
        End If
vbwProfiler.vbwExecuteLine 10938 'B
    End If
vbwProfiler.vbwExecuteLine 10939 'B
vbwProfiler.vbwProcOut 525
vbwProfiler.vbwExecuteLine 10940
End Property

Public Property Let scrollPosition(ByVal l As Long)
vbwProfiler.vbwProcIn 526
vbwProfiler.vbwExecuteLine 10941
    m_lngScrollPosition = l + 1
vbwProfiler.vbwProcOut 526
vbwProfiler.vbwExecuteLine 10942
End Property
Public Property Get maxVisibleRowCount() As Long
vbwProfiler.vbwProcIn 527
vbwProfiler.vbwExecuteLine 10943
    maxVisibleRowCount = m_lngMaxVisible
vbwProfiler.vbwProcOut 527
vbwProfiler.vbwExecuteLine 10944
End Property
Public Property Get rowCount() As Long
vbwProfiler.vbwProcIn 528
vbwProfiler.vbwExecuteLine 10945
    rowCount = m_lngNumRows
vbwProfiler.vbwProcOut 528
vbwProfiler.vbwExecuteLine 10946
End Property
Public Property Let RowSeperation(ByVal lngValue As Long)
    '//the amount of space between rows
vbwProfiler.vbwProcIn 529
vbwProfiler.vbwExecuteLine 10947
    m_lngRowSeperation = lngValue
vbwProfiler.vbwProcOut 529
vbwProfiler.vbwExecuteLine 10948
End Property

Public Property Let RowLeftPadding(ByVal lngValue As Long)
vbwProfiler.vbwProcIn 530
vbwProfiler.vbwExecuteLine 10949
    m_lngRowLeftPadding = lngValue
vbwProfiler.vbwProcOut 530
vbwProfiler.vbwExecuteLine 10950
End Property

Public Property Let RowRightPadding(ByVal lngValue As Long)
    '//the size of the right hand border
vbwProfiler.vbwProcIn 531
vbwProfiler.vbwExecuteLine 10951
    m_lngRowRightPadding = lngValue
vbwProfiler.vbwProcOut 531
vbwProfiler.vbwExecuteLine 10952
End Property


