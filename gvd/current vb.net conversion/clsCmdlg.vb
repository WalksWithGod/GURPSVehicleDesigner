Option Strict Off
Option Explicit On
Friend Class clsCmdlg
	
	
	Private Declare Function GetParent Lib "user32" (ByVal hwnd As Integer) As Integer
	
	' Win32 Declarations for the Common Dialog
	
	'UPGRADE_WARNING: Structure OPENFILENAME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetOpenFileName Lib "comdlg32.dll"  Alias "GetOpenFileNameA"(ByRef pOpenfilename As OPENFILENAME) As Integer
	'UPGRADE_WARNING: Structure OPENFILENAME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetSaveFileName Lib "comdlg32.dll"  Alias "GetSaveFileNameA"(ByRef pOpenfilename As OPENFILENAME) As Integer
	'UPGRADE_WARNING: Structure CHOOSEFONT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CHOOSEFONT_Renamed Lib "comdlg32.dll"  Alias "ChooseFontA"(ByRef pChoosefont As CHOOSEFONT) As Integer
	'UPGRADE_WARNING: Structure ChooseColor may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function ChooseColor_Renamed Lib "comdlg32.dll"  Alias "ChooseColorA"(ByRef pChoosecolor As ChooseColor) As Integer
	'UPGRADE_WARNING: Structure PageSetupDlg may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function PageSetupDlg_Renamed Lib "comdlg32.dll"  Alias "PageSetupDlgA"(ByRef pPagesetupdlg As PageSetupDlg) As Integer
	'UPGRADE_WARNING: Structure PRINTDLG_TYPE may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function PrintDialog Lib "comdlg32.dll"  Alias "PrintDlgA"(ByRef pPrintdlg As PRINTDLG_TYPE) As Integer
	
	' Win32 Declarations for the ShowFont function
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef Destination As Any, ByRef Source As Any, ByVal Length As Integer)
	Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Integer, ByVal dwBytes As Integer) As Integer
	Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Integer) As Integer
	Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Integer) As Integer
	Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Integer) As Integer
	Private Const LF_FACESIZE As Short = 32
	Private Const CCHDEVICENAME As Short = 32
	Private Const CCHFORMNAME As Short = 32
	
	Private Structure OPENFILENAME
		Dim lStructSize As Integer
		Dim hwndOwner As Integer
		Dim hInstance As Integer
		Dim lpstrFilter As String
		Dim lpstrCustomFilter As String
		Dim nMaxCustFilter As Integer
		Dim nFilterIndex As Integer
		Dim lpstrFile As String
		Dim nMaxFile As Integer
		Dim lpstrFileTitle As String
		Dim nMaxFileTitle As Integer
		Dim lpstrInitialDir As String
		Dim lpstrTitle As String
		Dim flags As Integer
		Dim nFileOffset As Short
		Dim nFileExtension As Short
		Dim lpstrDefExt As String
		Dim lCustData As Integer
		Dim lpfnHook As Integer
		Dim lpTemplateName As String
	End Structure
	
	Private Structure CHOOSEFONT
		Dim lStructSize As Integer
		Dim hwndOwner As Integer
		Dim hdc As Integer
		Dim lpLogFont As Integer
		Dim iPointSize As Integer
		Dim flags As Integer
		Dim rgbColors As Integer
		Dim lCustData As Integer
		Dim lpfnHook As Integer
		Dim lpTemplateName As String
		Dim hInstance As Integer
		Dim lpszStyle As String
		Dim nFontType As Short
		Dim MISSING_ALIGNMENT As Short
		Dim nSizeMin As Integer
		Dim nSizeMax As Integer
	End Structure
	
	Private Structure ChooseColor
		Dim lStructSize As Integer
		Dim hwndOwner As Integer
		Dim hInstance As Integer
		Dim rgbResult As Integer
		Dim lpCustColors As Integer
		Dim flags As Integer
		Dim lCustData As Integer
		Dim lpfnHook As Integer
		Dim lpTemplateName As String
	End Structure
	
	Private Structure PRINTDLG_TYPE
		Dim lStructSize As Integer
		Dim hwndOwner As Integer
		Dim hDevMode As Integer
		Dim hDevNames As Integer
		Dim hdc As Integer
		Dim flags As Integer
		Dim nFromPage As Short
		Dim nToPage As Short
		Dim nMinPage As Short
		Dim nMaxPage As Short
		Dim nCopies As Short
		Dim hInstance As Integer
		Dim lCustData As Integer
		Dim lpfnPrintHook As Integer
		Dim lpfnSetupHook As Integer
		Dim lpPrintTemplateName As String
		Dim lpSetupTemplateName As String
		Dim hPrintTemplate As Integer
		Dim hSetupTemplate As Integer
	End Structure
	
	Private Structure DEVNAMES_TYPE
		Dim wDriverOffset As Short
		Dim wDeviceOffset As Short
		Dim wOutputOffset As Short
		Dim wDefault As Short
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(100),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=100)> Public extra() As Char
	End Structure
	
	Private Structure DEVMODE_TYPE
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(CCHDEVICENAME),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=CCHDEVICENAME)> Public dmDeviceName() As Char
		Dim dmSpecVersion As Short
		Dim dmDriverVersion As Short
		Dim dmSize As Short
		Dim dmDriverExtra As Short
		Dim dmFields As Integer
		Dim dmOrientation As Short
		Dim dmPaperSize As Short
		Dim dmPaperLength As Short
		Dim dmPaperWidth As Short
		Dim dmScale As Short
		Dim dmCopies As Short
		Dim dmDefaultSource As Short
		Dim dmPrintQuality As Short
		Dim dmColor As Short
		Dim dmDuplex As Short
		Dim dmYResolution As Short
		Dim dmTTOption As Short
		Dim dmCollate As Short
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(CCHFORMNAME),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=CCHFORMNAME)> Public dmFormName() As Char
		Dim dmUnusedPadding As Short
		Dim dmBitsPerPel As Short
		Dim dmPelsWidth As Integer
		Dim dmPelsHeight As Integer
		Dim dmDisplayFlags As Integer
		Dim dmDisplayFrequency As Integer
	End Structure
	
	Private Structure POINTAPI
		Dim x As Integer
		Dim y As Integer
	End Structure
	
	Private Structure Rect
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	
	Private Structure PageSetupDlg
		Dim lStructSize As Integer
		Dim hwndOwner As Integer
		Dim hDevMode As Integer
		Dim hDevNames As Integer
		Dim flags As Integer
		Dim ptPaperSize As POINTAPI
		Dim rtMinMargin As Rect
		Dim rtMargin As Rect
		Dim hInstance As Integer
		Dim lCustData As Integer
		Dim lpfnPageSetupHook As Integer
		Dim lpfnPagePaintHook As Integer
		Dim lpPageSetupTemplateName As String
		Dim hPageSetupTemplate As Integer
	End Structure
	
	Private Structure LOGFONT
		Dim lfHeight As Integer
		Dim lfWidth As Integer
		Dim lfEscapement As Integer
		Dim lfOrientation As Integer
		Dim lfWeight As Integer
		Dim lfItalic As Byte
		Dim lfUnderline As Byte
		Dim lfStrikeout As Byte
		Dim lfCharSet As Byte
		Dim lfOutPrecision As Byte
		Dim lfClipPrecision As Byte
		Dim lfQuality As Byte
		Dim lfPitchAndFamily As Byte
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(31),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=31)> Public lfFaceName() As Char
	End Structure
	
	' Constants for the common dialog
	Private Const OFN_ALLOWMULTISELECT As Short = &H200s 'Allow multi select (Open Dialog)
	Private Const OFN_EXPLORER As Integer = &H80000 'Set windows style explorer
	Private Const OFN_FILEMUSTEXIST As Short = &H1000s 'File must exist
	Private Const OFN_HIDEREADONLY As Short = &H4s 'Hide read-only check box (Open Dialog)
	Private Const OFN_OVERWRITEPROMPT As Short = &H2s 'Promt beafore overwritning file (Save Dialog)
	Private Const OFN_PATHMUSTEXIST As Short = &H800s 'Path must exist
	Private Const CF_PRINTERFONTS As Short = &H2s
	Private Const CF_SCREENFONTS As Short = &H1s
	Private Const CF_BOTH As Boolean = (CF_SCREENFONTS Or CF_PRINTERFONTS)
	Private Const CF_EFFECTS As Integer = &H100
	Private Const CF_FORCEFONTEXIST As Integer = &H10000
	Private Const CF_INITTOLOGFONTSTRUCT As Integer = &H40
	Private Const CF_LIMITSIZE As Integer = &H2000
	Private Const DEFAULT_CHARSET As Short = 1
	Private Const DEFAULT_PITCH As Short = 0
	Private Const DEFAULT_QUALITY As Short = 0
	Private Const FW_BOLD As Short = 700
	Private Const FF_ROMAN As Short = 16 '  Variable stroke width, serifed.
	Private Const FW_NORMAL As Short = 400
	Private Const CLIP_DEFAULT_PRECIS As Short = 0
	Private Const OUT_DEFAULT_PRECIS As Short = 0
	Private Const REGULAR_FONTTYPE As Short = &H400s
	Private Const DM_DUPLEX As Integer = &H1000
	Private Const DM_ORIENTATION As Integer = &H1
	
	' Constants for the GlobalAllocate
	Private Const GMEM_MOVEABLE As Short = &H2s
	Private Const GMEM_ZEROINIT As Short = &H40s
	
	Private Const MAX_PATH As Short = 260 'Constant for maximum path
	
	Public cFileName As Collection 'Filename collection
	Public cFileTitle As Collection 'Filetitle collection
	
	' Default Property Values:
	Const m_def_CancelError As Short = 0
	Const m_def_Filename As String = ""
	Const m_def_DialogTitle As String = ""
	Const m_def_InitialDir As String = ""
	Const m_def_Filter As String = ""
	Const m_def_FilterIndex As Short = 1
	Const m_def_MultiSelect As Short = 0
	Const m_def_FontName As String = "Arial"
	Const m_def_FontSize As Short = 10
	Const m_def_FontColor As Short = 0
	Const m_def_FontBold As Short = 0
	Const m_def_FontItalic As Short = 0
	Const m_def_FontUnderline As Short = 0
	Const m_def_FontStrikeThru As Short = 0
	
	' Property Variables:
	Dim m_CancelError As Boolean
	Dim m_Filename As String
	Dim m_DialogTitle As String
	Dim m_InitialDir As String
	Dim m_Filter As String
	Dim m_FilterIndex As Short
	Dim m_MultiSelect As Boolean
	Dim m_FontName As String
	Dim m_FontSize As Short
	Dim m_FontColor As Integer
	Dim m_FontBold As Boolean
	Dim m_FontItalic As Boolean
	Dim m_FontUnderline As Boolean
	Dim m_FontStrikeThru As Boolean
	
	'***** CANCEL ERROR
	Public Property CancelError() As Boolean
		Get
			CancelError = m_CancelError
		End Get
		Set(ByVal Value As Boolean)
			m_CancelError = Value
			
		End Set
	End Property
	'***** MULTI SELECT
	Public Property MultiSelect() As Boolean
		Get
			MultiSelect = m_MultiSelect
		End Get
		Set(ByVal Value As Boolean)
			m_MultiSelect = Value
		End Set
	End Property
	'***** DEFAULT FILENAME
	Public Property DefaultFilename() As String
		Get
			DefaultFilename = m_Filename
		End Get
		Set(ByVal Value As String)
			m_Filename = Value
		End Set
	End Property
	'***** DIALOG TITLE
	Public Property DialogTitle() As String
		Get
			DialogTitle = m_DialogTitle
		End Get
		Set(ByVal Value As String)
			m_DialogTitle = Value
		End Set
	End Property
	'***** INITIAL DIRECTORY
	Public Property InitialDir() As String
		Get
			InitialDir = m_InitialDir
		End Get
		Set(ByVal Value As String)
			m_InitialDir = Value
		End Set
	End Property
	'***** FILTER
	'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Property Filter_Renamed() As String
		Get
			Filter_Renamed = m_Filter
		End Get
		Set(ByVal Value As String)
			m_Filter = Value
		End Set
	End Property
	'***** FILTER INDEX
	Public Property FilterIndex() As Short
		Get
			FilterIndex = m_FilterIndex
		End Get
		Set(ByVal Value As Short)
			m_FilterIndex = Value
		End Set
	End Property
	'***** FONT NAME
	Public Property FontName() As String
		Get
			FontName = m_FontName
		End Get
		Set(ByVal Value As String)
			m_FontName = Value
		End Set
	End Property
	'***** FONT SIZE
	Public Property FontSize() As Short
		Get
			FontSize = m_FontSize
		End Get
		Set(ByVal Value As Short)
			m_FontSize = Value
		End Set
	End Property
	'***** FONT COLOR
	Public Property FontColor() As Integer
		Get
			FontColor = m_FontColor
		End Get
		Set(ByVal Value As Integer)
			m_FontColor = Value
		End Set
	End Property
	'***** FONT BOLD
	Public Property FontBold() As Boolean
		Get
			FontBold = m_FontBold
		End Get
		Set(ByVal Value As Boolean)
			m_FontBold = Value
		End Set
	End Property
	'***** FONT ITALIC
	Public Property FontItalic() As Boolean
		Get
			FontItalic = m_FontItalic
		End Get
		Set(ByVal Value As Boolean)
			m_FontItalic = Value
		End Set
	End Property
	'***** FONT UNDERLINE
	Public Property FontUnderline() As Boolean
		Get
			FontUnderline = m_FontUnderline
		End Get
		Set(ByVal Value As Boolean)
			m_FontUnderline = Value
		End Set
	End Property
	'***** FONT STRIKETHRU
	Public Property FontStrikeThru() As Boolean
		Get
			FontStrikeThru = m_FontStrikeThru
		End Get
		Set(ByVal Value As Boolean)
			m_FontStrikeThru = Value
		End Set
	End Property
	
	' Initialize Properties for class
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		m_CancelError = m_def_CancelError
		m_Filename = m_def_Filename
		m_DialogTitle = m_def_DialogTitle
		m_InitialDir = m_def_InitialDir
		m_Filter = m_def_Filter
		m_FilterIndex = m_def_FilterIndex
		m_MultiSelect = m_def_MultiSelect
		m_FontName = m_def_FontName
		m_FontSize = m_def_FontSize
		m_FontColor = m_def_FontColor
		m_FontBold = m_def_FontBold
		m_FontItalic = m_def_FontItalic
		m_FontUnderline = m_def_FontUnderline
		m_FontStrikeThru = m_def_FontStrikeThru
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	
	Public Function ShowOpen(ByVal hParent As Integer) As Boolean
		'** Description:
		'** Calls open dialog without OCX
		Dim epOFN As OPENFILENAME
		Dim lngRet As Integer
		With epOFN
			
			If MultiSelect Then 'If Multi Select then
				.flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
				.lpstrFile = DefaultFilename & Space(9999 - Len(DefaultFilename)) & vbNullChar
				.lpstrFileTitle = Space(9999) & vbNullChar
			Else
				.flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
				.lpstrFile = DefaultFilename & New String(Chr(0), MAX_PATH - Len(DefaultFilename)) & vbNullChar
				.lpstrFileTitle = New String(Chr(0), MAX_PATH) & vbNullChar
			End If
			
			.hwndOwner = hParent 'Handle to window
			.lpstrFilter = SetFilter(Filter_Renamed) & vbNullChar 'File filter
			.lpstrInitialDir = InitialDir & vbNullChar 'Initial directory
			.lpstrTitle = DialogTitle & vbNullChar 'Dialog title
			.lStructSize = Len(epOFN) 'Structure size in bytes
			.nFilterIndex = FilterIndex 'Filter index
			.nMaxFile = Len(.lpstrFile) 'Maximum file length
			.nMaxFileTitle = Len(.lpstrFileTitle) 'Maximum file title length
		End With
		
		lngRet = GetOpenFileName(epOFN) 'Call open dialog
		
		If lngRet <> 0 Then 'If there are no errors continue with opening file
			ParseFileName(epOFN.lpstrFile)
			ShowOpen = True
		Else
			If CancelError Then
				' For this to work you must check in Tools\Options\General
				' Break on Unhandled errors if it isn't already checked
				'err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
				ShowOpen = False
			End If
		End If
	End Function
	
	Public Function ShowSave(ByVal hParent As Integer) As Integer
		'** Description:
		'** Calls save dialog without OCX
		Dim epOFN As OPENFILENAME
		Dim lngRet As Integer
		With epOFN
			.hwndOwner = hParent 'Handle to parent window
			.flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
			.lpstrFile = DefaultFilename & New String(Chr(0), MAX_PATH - Len(DefaultFilename)) & vbNullChar
			.lpstrFileTitle = New String(Chr(0), MAX_PATH) & vbNullChar
			.lpstrFilter = SetFilter(Filter_Renamed) & vbNullChar 'File filter
			.lpstrInitialDir = InitialDir & vbNullChar 'Initial directory
			.lpstrTitle = DialogTitle & vbNullChar 'Dialog title
			.lStructSize = Len(epOFN) 'Structure size in bytes
			.nFilterIndex = FilterIndex 'Filter index
			.nMaxFile = Len(.lpstrFile) 'Maximum file length
			.nMaxFileTitle = Len(.lpstrFileTitle) 'Maximum file title length
		End With
		
		lngRet = GetSaveFileName(epOFN) 'Call save dialog
		
		If lngRet <> 0 Then 'If there are no errors continue with saving file
			ParseFileName(epOFN.lpstrFile)
			ShowSave = False
		Else
			If CancelError Then
				' For this to work you must check in Tools\Options\General
				' Break on Unhandled errors if it isn't already checked
				'err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
				ShowSave = True
			End If
		End If
	End Function
	
	Public Function ShowFont(ByVal hParent As Integer) As Integer
		'** Description:
		'** Call font dialog without OCX
		Dim CF As CHOOSEFONT
		Dim lf As LOGFONT
		Dim lMemHandle As Integer
		Dim lLogFont As Integer
		Dim lngRet As Integer
		
		With lf
			.lfCharSet = DEFAULT_CHARSET 'Default character set
			.lfClipPrecision = CLIP_DEFAULT_PRECIS 'Clipping precision
			.lfFaceName = "Arial" & vbNullChar 'Font name
			.lfHeight = 13 'Height
			.lfOutPrecision = OUT_DEFAULT_PRECIS 'Precision mapping
			.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN 'Default pitch
			.lfQuality = DEFAULT_QUALITY 'Default quality
			.lfWeight = FW_NORMAL 'Regular font type
		End With
		
		' Create the memory block
		lMemHandle = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lf))
		lLogFont = GlobalLock(lMemHandle)
		'UPGRADE_WARNING: Couldn't resolve default property of object lf. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(lLogFont, lf, Len(lf))
		
		With CF
			.flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
			'UPGRADE_ISSUE: Printer property Printer.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			.hdc = Printer.hDC 'Device context of default printer
			.hwndOwner = hParent 'Handle to window
			.iPointSize = 120 'Set font size to 12 size
			.lpLogFont = lLogFont 'Log font
			.lStructSize = Len(CF) 'Size of structure in bytes
			.nFontType = REGULAR_FONTTYPE 'Regular font type
			.nSizeMax = 72 'Maximum font size
			.nSizeMin = 10 'Minimum font size
			.rgbColors = RGB(0, 0, 0) 'Font color
		End With
		
		lngRet = CHOOSEFONT_Renamed(CF) 'Call font dialog
		If lngRet <> 0 Then 'If there are no errors continue with font
			'UPGRADE_WARNING: Couldn't resolve default property of object lf. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(lf, lLogFont, Len(lf))
			
			FontName = Left(lf.lfFaceName, InStr(lf.lfFaceName, vbNullChar) - 1)
			FontSize = CF.iPointSize / 10
			FontColor = CF.rgbColors
			If lf.lfWeight = FW_NORMAL Then
				FontBold = False
				FontItalic = False
				FontUnderline = False
				FontStrikeThru = False
			Else
				If lf.lfWeight = FW_BOLD Then FontBold = True
				If lf.lfItalic <> 0 Then FontItalic = True
				If lf.lfUnderline <> 0 Then FontUnderline = True
				If lf.lfStrikeout <> 0 Then FontStrikeThru = True
			End If
			ShowFont = False
		Else
			If CancelError Then
				' For this to work you must check in Tools\Options\General
				' Break on Unhandled errors if it isn't already checked
				'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				err.Raise(32755, My.Application.Info.AssemblyName, "Cancel was selected.", "cmdlg98.chm", 32755)
				ShowFont = True
			End If
		End If
		
		' Unlock and free the memory block
		' Note this must be done
		GlobalUnlock(lMemHandle)
		GlobalFree(lMemHandle)
	End Function
	
	Public Function ShowColor(ByVal hParent As Integer) As Integer
		'** Description:
		'** Call color dialog without OCX
		Dim epCC As ChooseColor
		Dim lngRet As Integer
		Dim CusCol(16) As Integer
		Dim i As Short
		
		' Fills custom colors with white
		For i = 0 To 15
			CusCol(i) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
		Next 
		
		With epCC
			.hwndOwner = hParent 'Handle to owner
			.lStructSize = Len(epCC) 'Structure size in bytes
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			.lpCustColors = VarPtr(CusCol(0)) 'Custom colors
			.rgbResult = 0 'RGB result
		End With
		
		lngRet = ChooseColor_Renamed(epCC) 'Call color dialog
		If lngRet <> 0 Then 'If there are no errors continue with color
			ShowColor = epCC.rgbResult
		Else
			If CancelError Then
				' For this to work you must check in Tools\Options\General
				' Break on Unhandled errors if it isn't already checked
				'err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
				ShowColor = True
			End If
		End If
	End Function
	
	Public Function ShowPageSetup(ByVal hParent As Integer) As Integer
		'** Description:
		'** Call page setup dialog without OCX
		Dim epPSD As PageSetupDlg
		Dim lngRet As Integer
		
		epPSD.lStructSize = Len(epPSD) 'Structure size in bytes
		epPSD.hwndOwner = hParent
		
		lngRet = PageSetupDlg_Renamed(epPSD) 'Call page setup dialog
		If lngRet <> 0 Then 'If there are no errors continue
			ShowPageSetup = False
		Else
			If CancelError Then
				' For this to work you must check in Tools\Options\General
				' Break on Unhandled errors if it isn't already checked
				' err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
				ShowPageSetup = True
			End If
		End If
	End Function
	
	Public Function ShowPrinter(ByVal hParent As Integer) As Integer
		'** Description:
		'** Call printer dialog without OCX
		'**
		'** Note:
		'** This is not my function it's from KPD-Team 1998 URL: http://www.allapi.net
		'** and i have modified it a little
		'-> Code by Donald Grover
		Dim PrintDlg As PRINTDLG_TYPE
		Dim DevMode As DEVMODE_TYPE
		Dim DevName As DEVNAMES_TYPE
		
		Dim lpDevMode, lpDevName As Integer
		Dim bReturn As Short
		'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim objPrinter As Printer
		Dim NewPrinterName As String
		
		' Use PrintDialog to get the handle to a memory
		' block with a DevMode and DevName structures
		
		PrintDlg.lStructSize = Len(PrintDlg)
		PrintDlg.hwndOwner = hParent 'Handle to window
		
		On Error Resume Next
		'Set the current orientation and duplex setting
		'UPGRADE_ISSUE: Printer property Printer.DeviceName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		DevMode.dmDeviceName = Printer.DeviceName
		DevMode.dmSize = Len(DevMode)
		DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
		'UPGRADE_ISSUE: Printer property Printer.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		DevMode.dmPaperWidth = Printer.Width
		'UPGRADE_ISSUE: Printer property Printer.Orientation was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		DevMode.dmOrientation = Printer.Orientation
		'UPGRADE_ISSUE: Printer property Printer.PaperSize was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		DevMode.dmPaperSize = Printer.PaperSize
		'UPGRADE_ISSUE: Printer property Printer.Duplex was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		DevMode.dmDuplex = Printer.Duplex
		On Error GoTo 0
		
		'Allocate memory for the initialization hDevMode structure
		'and copy the settings gathered above into this memory
		PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
		lpDevMode = GlobalLock(PrintDlg.hDevMode)
		If lpDevMode > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object DevMode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(lpDevMode, DevMode, Len(DevMode))
			bReturn = GlobalUnlock(PrintDlg.hDevMode)
		End If
		
		'Set the current driver, device, and port name strings
		With DevName
			.wDriverOffset = 8
			'UPGRADE_ISSUE: Printer property Printer.DriverName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			.wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
			'UPGRADE_ISSUE: Printer property Printer.Port was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			.wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
			.wDefault = 0
		End With
		
		'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		With Printer
			'UPGRADE_ISSUE: Printer property Printer.Port was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			'UPGRADE_ISSUE: Printer property Printer.DeviceName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			'UPGRADE_ISSUE: Printer property Printer.DriverName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
		End With
		
		'Allocate memory for the initial hDevName structure
		'and copy the settings gathered above into this memory
		PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
		lpDevName = GlobalLock(PrintDlg.hDevNames)
		If lpDevName > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object DevName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(lpDevName, DevName, Len(DevName))
			bReturn = GlobalUnlock(lpDevName)
		End If
		
		'Call the print dialog up and let the user make changes
		If PrintDialog(PrintDlg) <> 0 Then
			
			'First get the DevName structure.
			lpDevName = GlobalLock(PrintDlg.hDevNames)
			'UPGRADE_WARNING: Couldn't resolve default property of object DevName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(DevName, lpDevName, 45)
			bReturn = GlobalUnlock(lpDevName)
			GlobalFree(PrintDlg.hDevNames)
			
			'Next get the DevMode structure and set the printer
			'properties appropriately
			lpDevMode = GlobalLock(PrintDlg.hDevMode)
			'UPGRADE_WARNING: Couldn't resolve default property of object DevMode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(DevMode, lpDevMode, Len(DevMode))
			bReturn = GlobalUnlock(PrintDlg.hDevMode)
			GlobalFree(PrintDlg.hDevMode)
			NewPrinterName = UCase(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr(0)) - 1))
			'UPGRADE_ISSUE: Printer property Printer.DeviceName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			If Printer.DeviceName <> NewPrinterName Then
				'UPGRADE_ISSUE: Printers object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
				For	Each objPrinter In Printers
					'UPGRADE_ISSUE: Printer property objPrinter.DeviceName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					If UCase(objPrinter.DeviceName) = NewPrinterName Then
						'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
						Printer = objPrinter
						'set printer toolbar name at this point
					End If
				Next objPrinter
			End If
			
			On Error Resume Next
			'Set printer object properties according to selections made
			'by user
			'UPGRADE_ISSUE: Printer property Printer.Copies was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Printer.Copies = DevMode.dmCopies
			'UPGRADE_ISSUE: Printer property Printer.Duplex was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Printer.Duplex = DevMode.dmDuplex
			'UPGRADE_ISSUE: Printer property Printer.Orientation was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Printer.Orientation = DevMode.dmOrientation
			'UPGRADE_ISSUE: Printer property Printer.PaperSize was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Printer.PaperSize = DevMode.dmPaperSize
			'UPGRADE_ISSUE: Printer property Printer.PrintQuality was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Printer.PrintQuality = DevMode.dmPrintQuality
			'UPGRADE_ISSUE: Printer property Printer.ColorMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Printer.ColorMode = DevMode.dmColor
			'UPGRADE_ISSUE: Printer property Printer.PaperBin was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Printer.PaperBin = DevMode.dmDefaultSource
			ShowPrinter = False
			On Error GoTo 0
		Else
			If CancelError Then
				' For this to work you must check in Tools\Options\General
				' Break on Unhandled errors if it isn't already checked
				' err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
				ShowPrinter = True
			End If
		End If
	End Function
	
	Private Function ParseFileName(ByRef sFileName As String) As Object
		'** Description:
		'** Remove null chars from filename and parse multi filename
		'**
		'** Syntax:
		'** szFilename = ParseFileName(strFilename)
		'**
		'** Example:
		'** szFilename = ParseFileName("C:\Autoexec.bat||")
		Dim i As Integer
		Dim sPath As String
		Dim sFiles() As String
		Dim Pos As Short
		Dim sFile As String
		Dim sFileTitle As String
		
		' Create new collections
		cFileName = New Collection
		cFileTitle = New Collection
		' Found position of two last null chars
		Pos = InStr(sFileName, vbNullChar & vbNullChar)
		' Remove from filename last two chars
		sFile = Left(sFileName, Pos - 1)
		
		' Check to see if filename is single or multi
		If InStr(1, sFile, vbNullChar) <> 0 Then
			' Multi file
			sFile = Left(sFileName, Pos) & vbNullChar 'Add null char at end of filename
			sPath = Left(sFileName, InStr(1, sFileName, Chr(0)) - 1) 'Get file path
			sFiles = Split(sFile, Chr(0)) 'Split file where is nullchar
			
			' Add all filenames to collection
			For i = LBound(sFiles) To UBound(sFiles) - 2
				' If path doesent contain separator then add it
				If Right(sPath, 1) = "\" Then
					cFileName.Add(sPath & sFiles(i))
				Else
					cFileName.Add(sPath & "\" & sFiles(i))
				End If
				' Add file title
				cFileTitle.Add(sFiles(i))
				' Remove first item from collections
				If i = 1 Then cFileName.Remove(1) : cFileTitle.Remove(1)
			Next 
		Else ' Single file
			'Add file name to collection
			cFileName.Add(sFile)
			' Add file title
			cFileTitle.Add(Right(sFile, Len(sFile) - InStrRev(sFile, "\")))
		End If
	End Function
	
	Private Function SetFilter(ByRef sFlt As String) As String
		'** Description:
		'** Replace "|" with Null Character
		'**
		'** Syntax:
		'** szFilter = SetFilter(strFilter)
		'**
		'** Example:
		'** szFilter = SetFilter("Text Files (*.txt)|*.txt|All Files |*.*|")
		Dim sLen As Integer
		Dim Pos As Integer
		
		sLen = Len(sFlt) 'Get filter length
		Pos = InStr(1, sFlt, "|") 'Find first position of "|"
		
		' Loop while Pos > 0
		While Pos > 0
			' Replace "|" with null char
			sFlt = Left(sFlt, Pos - 1) & vbNullChar & Mid(sFlt, Pos + 1, sLen - Pos)
			' Find next position of "|"
			Pos = InStr(Pos + 1, sFlt, "|")
		End While
		SetFilter = sFlt ' Set filter
	End Function
End Class