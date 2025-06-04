Option Strict Off
Option Explicit On
Module modPrinter
	' Win32 Declarations for Print sub
	Private Structure Rect
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	
	Private Structure CharRange
		Dim cpMin As Integer ' First character of range (0 for start of doc)
		Dim cpMax As Integer ' Last character of range (-1 for end of doc)
	End Structure
	
	Private Structure FormatRange
		Dim hdc As Integer ' Actual DC to draw on
		Dim hdcTarget As Integer ' Target DC for determining text formatting
		Dim rc As Rect ' Region of the DC to draw to (in twips)
		Dim rcPage As Rect ' Region of the entire DC (page size) (in twips)
		Dim chrg As CharRange ' Range of text to draw (see above declaration)
	End Structure
	
	Const WM_USER As Short = &H400s
	Const EM_FORMATRANGE As Integer = WM_USER + 57
	Const PHYSICALOFFSETX As Integer = 112
	Const PHYSICALOFFSETY As Integer = 113
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
	Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
	
	Public Sub PrintRTF(ByRef RTF As RichTextBox, ByRef LeftMarginWidth As Integer, ByRef TopMarginHeight As Object, ByRef RightMarginWidth As Object, ByRef BottomMarginHeight As Object)
		'** Description:
		'** Print the active document
		On Error GoTo PrintError
		Dim LeftOffset, TopOffset As Integer
		Dim LeftMargin, TopMargin As Integer
		Dim RightMargin, BottomMargin As Integer
		Dim fr As FormatRange
		Dim rcDrawTo As Rect
		Dim rcPage As Rect
		Dim TextLength As Integer
		Dim NextCharPosition As Integer
		Dim r As Integer
		
		' Start a print job to get a valid Printer.hDC
		'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.Print(Space(1))
		'UPGRADE_ISSUE: Constant vbTwips was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Printer property Printer.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.ScaleMode = vbTwips
		
		' Get the offsett to the printable area on the page in twips
		'UPGRADE_ISSUE: Constant vbTwips was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Printer property Printer.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_ISSUE: Printer method Printer.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
		'UPGRADE_ISSUE: Constant vbTwips was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Printer property Printer.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_ISSUE: Printer method Printer.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
		
		' Calculate the Left, Top, Right, and Bottom margins
		LeftMargin = LeftMarginWidth - LeftOffset
		'UPGRADE_WARNING: Couldn't resolve default property of object TopMarginHeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TopMargin = TopMarginHeight - TopOffset
		'UPGRADE_WARNING: Couldn't resolve default property of object RightMarginWidth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_ISSUE: Printer property Printer.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
		'UPGRADE_WARNING: Couldn't resolve default property of object BottomMarginHeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_ISSUE: Printer property Printer.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset
		
		' Set printable area rect
		rcPage.Left_Renamed = 0
		rcPage.Top = 0
		'UPGRADE_ISSUE: Printer property Printer.ScaleWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		rcPage.Right_Renamed = Printer.ScaleWidth
		'UPGRADE_ISSUE: Printer property Printer.ScaleHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		rcPage.Bottom = Printer.ScaleHeight
		
		' Set rect in which to print (relative to printable area)
		rcDrawTo.Left_Renamed = LeftMargin
		rcDrawTo.Top = TopMargin
		rcDrawTo.Right_Renamed = RightMargin
		rcDrawTo.Bottom = BottomMargin
		
		' Set up the print instructions
		'UPGRADE_ISSUE: Printer property Printer.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		fr.hdc = Printer.hDC ' Use the same DC for measuring and rendering
		'UPGRADE_ISSUE: Printer property Printer.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		fr.hdcTarget = Printer.hDC ' Point at printer hDC
		'UPGRADE_WARNING: Couldn't resolve default property of object fr.rc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fr.rc = rcDrawTo ' Indicate the area on page to draw to
		'UPGRADE_WARNING: Couldn't resolve default property of object fr.rcPage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fr.rcPage = rcPage ' Indicate entire size of page
		fr.chrg.cpMin = 0 ' Indicate start of text through
		fr.chrg.cpMax = -1 ' end of the text
		
		' Get length of text in RTF
		'UPGRADE_WARNING: Couldn't resolve default property of object RTF.Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TextLength = Len(RTF.Text)
		
		' Loop printing each page until done
		Do 
			' Print the page by sending EM_FORMATRANGE message
			'UPGRADE_WARNING: Couldn't resolve default property of object fr. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object RTF.hWnd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			NextCharPosition = SendMessage(RTF.hWnd, EM_FORMATRANGE, True, fr)
			If NextCharPosition >= TextLength Then Exit Do 'If done then exit
			fr.chrg.cpMin = NextCharPosition ' Starting position for next page
			'UPGRADE_ISSUE: Printer method Printer.NewPage was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Printer.NewPage() ' Move on to next page
			'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
			'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Printer.Print(Space(1)) ' Re-initialize hDC
			'UPGRADE_ISSUE: Printer property Printer.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			fr.hdc = Printer.hDC
			'UPGRADE_ISSUE: Printer property Printer.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			fr.hdcTarget = Printer.hDC
		Loop 
		
		' Commit the print job
		'UPGRADE_ISSUE: Printer method Printer.EndDoc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.EndDoc()
		
		' Allow the RTF to free up memory
		'UPGRADE_WARNING: Couldn't resolve default property of object RTF.hWnd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		r = SendMessage(RTF.hWnd, EM_FORMATRANGE, False, CInt(0))
PrintError: 
		'ErrorLog "modAPI\PrintRTF"
	End Sub
End Module