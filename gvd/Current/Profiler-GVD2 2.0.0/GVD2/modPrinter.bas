Attribute VB_Name = "modPrinter"
Option Explicit
' Win32 Declarations for Print sub
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CharRange
    cpMin As Long     ' First character of range (0 for start of doc)
    cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
    hdc As Long       ' Actual DC to draw on
    hdcTarget As Long ' Target DC for determining text formatting
    rc As Rect        ' Region of the DC to draw to (in twips)
    rcPage As Rect    ' Region of the entire DC (page size) (in twips)
    chrg As CharRange ' Range of text to draw (see above declaration)
End Type

Const WM_USER = &H400
Const EM_FORMATRANGE As Long = WM_USER + 57
Const PHYSICALOFFSETX As Long = 112
Const PHYSICALOFFSETY As Long = 113

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, TopMarginHeight, RightMarginWidth, BottomMarginHeight)
    '** Description:
    '** Print the active document
vbwProfiler.vbwProcIn 282
vbwProfiler.vbwExecuteLine 8394
    On Error GoTo PrintError
    Dim LeftOffset As Long, TopOffset As Long
    Dim LeftMargin As Long, TopMargin As Long
    Dim RightMargin As Long, BottomMargin As Long
    Dim fr As FormatRange
    Dim rcDrawTo As Rect
    Dim rcPage As Rect
    Dim TextLength As Long
    Dim NextCharPosition As Long
    Dim r As Long

    ' Start a print job to get a valid Printer.hDC
vbwProfiler.vbwExecuteLine 8395
    Printer.Print Space(1)
vbwProfiler.vbwExecuteLine 8396
    Printer.ScaleMode = vbTwips

    ' Get the offsett to the printable area on the page in twips
vbwProfiler.vbwExecuteLine 8397
    LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
vbwProfiler.vbwExecuteLine 8398
    TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)

    ' Calculate the Left, Top, Right, and Bottom margins
vbwProfiler.vbwExecuteLine 8399
    LeftMargin = LeftMarginWidth - LeftOffset
vbwProfiler.vbwExecuteLine 8400
    TopMargin = TopMarginHeight - TopOffset
vbwProfiler.vbwExecuteLine 8401
    RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
vbwProfiler.vbwExecuteLine 8402
    BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset

    ' Set printable area rect
vbwProfiler.vbwExecuteLine 8403
    rcPage.Left = 0
vbwProfiler.vbwExecuteLine 8404
    rcPage.Top = 0
vbwProfiler.vbwExecuteLine 8405
    rcPage.Right = Printer.ScaleWidth
vbwProfiler.vbwExecuteLine 8406
    rcPage.Bottom = Printer.ScaleHeight

    ' Set rect in which to print (relative to printable area)
vbwProfiler.vbwExecuteLine 8407
    rcDrawTo.Left = LeftMargin
vbwProfiler.vbwExecuteLine 8408
    rcDrawTo.Top = TopMargin
vbwProfiler.vbwExecuteLine 8409
    rcDrawTo.Right = RightMargin
vbwProfiler.vbwExecuteLine 8410
    rcDrawTo.Bottom = BottomMargin

    ' Set up the print instructions
vbwProfiler.vbwExecuteLine 8411
    fr.hdc = Printer.hdc   ' Use the same DC for measuring and rendering
vbwProfiler.vbwExecuteLine 8412
    fr.hdcTarget = Printer.hdc  ' Point at printer hDC
vbwProfiler.vbwExecuteLine 8413
    fr.rc = rcDrawTo            ' Indicate the area on page to draw to
vbwProfiler.vbwExecuteLine 8414
    fr.rcPage = rcPage          ' Indicate entire size of page
vbwProfiler.vbwExecuteLine 8415
    fr.chrg.cpMin = 0           ' Indicate start of text through
vbwProfiler.vbwExecuteLine 8416
    fr.chrg.cpMax = -1          ' end of the text

    ' Get length of text in RTF
vbwProfiler.vbwExecuteLine 8417
    TextLength = Len(RTF.Text)

    ' Loop printing each page until done
vbwProfiler.vbwExecuteLine 8418
    Do
        ' Print the page by sending EM_FORMATRANGE message
vbwProfiler.vbwExecuteLine 8419
        NextCharPosition = SendMessage(RTF.hWnd, EM_FORMATRANGE, True, fr)
vbwProfiler.vbwExecuteLine 8420
        If NextCharPosition >= TextLength Then 'If done then exit
vbwProfiler.vbwExecuteLine 8421
             Exit Do
        End If
vbwProfiler.vbwExecuteLine 8422 'B
vbwProfiler.vbwExecuteLine 8423
        fr.chrg.cpMin = NextCharPosition ' Starting position for next page
vbwProfiler.vbwExecuteLine 8424
        Printer.NewPage                  ' Move on to next page
vbwProfiler.vbwExecuteLine 8425
        Printer.Print Space(1) ' Re-initialize hDC
vbwProfiler.vbwExecuteLine 8426
        fr.hdc = Printer.hdc
vbwProfiler.vbwExecuteLine 8427
        fr.hdcTarget = Printer.hdc
vbwProfiler.vbwExecuteLine 8428
    Loop

    ' Commit the print job
vbwProfiler.vbwExecuteLine 8429
    Printer.EndDoc

    ' Allow the RTF to free up memory
vbwProfiler.vbwExecuteLine 8430
    r = SendMessage(RTF.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))
PrintError:
    'ErrorLog "modAPI\PrintRTF"
vbwProfiler.vbwProcOut 282
vbwProfiler.vbwExecuteLine 8431
End Sub

