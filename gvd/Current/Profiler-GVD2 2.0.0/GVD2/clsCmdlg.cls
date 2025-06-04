VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCmdlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

' Win32 Declarations for the Common Dialog

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PageSetupDlg) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long

' Win32 Declarations for the ShowFont function
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const LF_FACESIZE = 32
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long
        hdc As Long
        lpLogFont As Long
        iPointSize As Long
        flags As Long
        rgbColors As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
        hInstance As Long
        lpszStyle As String
        nFontType As Integer
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long
        nSizeMax As Long
End Type

Private Type ChooseColor
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Type PRINTDLG_TYPE
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hdc As Long
        flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long
        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
        extra As String * 100
End Type

Private Type DEVMODE_TYPE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

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

Private Type PageSetupDlg
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        flags As Long
        ptPaperSize As POINTAPI
        rtMinMargin As Rect
        rtMargin As Rect
        hInstance As Long
        lCustData As Long
        lpfnPageSetupHook As Long
        lpfnPagePaintHook As Long
        lpPageSetupTemplateName As String
        hPageSetupTemplate As Long
End Type

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeout As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 31
End Type

' Constants for the common dialog
Private Const OFN_ALLOWMULTISELECT = &H200  'Allow multi select (Open Dialog)
Private Const OFN_EXPLORER = &H80000        'Set windows style explorer
Private Const OFN_FILEMUSTEXIST = &H1000    'File must exist
Private Const OFN_HIDEREADONLY = &H4        'Hide read-only check box (Open Dialog)
Private Const OFN_OVERWRITEPROMPT = &H2     'Promt beafore overwritning file (Save Dialog)
Private Const OFN_PATHMUSTEXIST = &H800     'Path must exist
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const DEFAULT_CHARSET = 1
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_QUALITY = 0
Private Const FW_BOLD = 700
Private Const FF_ROMAN = 16      '  Variable stroke width, serifed.
Private Const FW_NORMAL = 400
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const OUT_DEFAULT_PRECIS = 0
Private Const REGULAR_FONTTYPE = &H400
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&

' Constants for the GlobalAllocate
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40

Private Const MAX_PATH = 260 'Constant for maximum path

Public cFileName As Collection   'Filename collection
Public cFileTitle As Collection  'Filetitle collection

 ' Default Property Values:
Const m_def_CancelError = 0
Const m_def_Filename = ""
Const m_def_DialogTitle = ""
Const m_def_InitialDir = ""
Const m_def_Filter = ""
Const m_def_FilterIndex = 1
Const m_def_MultiSelect = 0
Const m_def_FontName = "Arial"
Const m_def_FontSize = 10
Const m_def_FontColor = 0
Const m_def_FontBold = 0
Const m_def_FontItalic = 0
Const m_def_FontUnderline = 0
Const m_def_FontStrikeThru = 0

 ' Property Variables:
Dim m_CancelError As Boolean
Dim m_Filename As String
Dim m_DialogTitle As String
Dim m_InitialDir As String
Dim m_Filter As String
Dim m_FilterIndex As Integer
Dim m_MultiSelect As Boolean
Dim m_FontName As String
Dim m_FontSize As Integer
Dim m_FontColor As Long
Dim m_FontBold As Boolean
Dim m_FontItalic As Boolean
Dim m_FontUnderline As Boolean
Dim m_FontStrikeThru As Boolean

'***** CANCEL ERROR
Public Property Get CancelError() As Boolean
vbwProfiler.vbwProcIn 440
vbwProfiler.vbwExecuteLine 9807
    CancelError = m_CancelError
vbwProfiler.vbwProcOut 440
vbwProfiler.vbwExecuteLine 9808
End Property
Public Property Let CancelError(ByVal New_CancelError As Boolean)
vbwProfiler.vbwProcIn 441
vbwProfiler.vbwExecuteLine 9809
    m_CancelError = New_CancelError

vbwProfiler.vbwProcOut 441
vbwProfiler.vbwExecuteLine 9810
End Property
'***** MULTI SELECT
Public Property Get MultiSelect() As Boolean
vbwProfiler.vbwProcIn 442
vbwProfiler.vbwExecuteLine 9811
    MultiSelect = m_MultiSelect
vbwProfiler.vbwProcOut 442
vbwProfiler.vbwExecuteLine 9812
End Property
Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
vbwProfiler.vbwProcIn 443
vbwProfiler.vbwExecuteLine 9813
    m_MultiSelect = New_MultiSelect
vbwProfiler.vbwProcOut 443
vbwProfiler.vbwExecuteLine 9814
End Property
'***** DEFAULT FILENAME
Public Property Get DefaultFilename() As String
vbwProfiler.vbwProcIn 444
vbwProfiler.vbwExecuteLine 9815
    DefaultFilename = m_Filename
vbwProfiler.vbwProcOut 444
vbwProfiler.vbwExecuteLine 9816
End Property
Public Property Let DefaultFilename(ByVal New_Filename As String)
vbwProfiler.vbwProcIn 445
vbwProfiler.vbwExecuteLine 9817
    m_Filename = New_Filename
vbwProfiler.vbwProcOut 445
vbwProfiler.vbwExecuteLine 9818
End Property
'***** DIALOG TITLE
Public Property Get DialogTitle() As String
vbwProfiler.vbwProcIn 446
vbwProfiler.vbwExecuteLine 9819
    DialogTitle = m_DialogTitle
vbwProfiler.vbwProcOut 446
vbwProfiler.vbwExecuteLine 9820
End Property
Public Property Let DialogTitle(ByVal New_DialogTitle As String)
vbwProfiler.vbwProcIn 447
vbwProfiler.vbwExecuteLine 9821
    m_DialogTitle = New_DialogTitle
vbwProfiler.vbwProcOut 447
vbwProfiler.vbwExecuteLine 9822
End Property
'***** INITIAL DIRECTORY
Public Property Get InitialDir() As String
vbwProfiler.vbwProcIn 448
vbwProfiler.vbwExecuteLine 9823
    InitialDir = m_InitialDir
vbwProfiler.vbwProcOut 448
vbwProfiler.vbwExecuteLine 9824
End Property
Public Property Let InitialDir(ByVal New_InitialDir As String)
vbwProfiler.vbwProcIn 449
vbwProfiler.vbwExecuteLine 9825
    m_InitialDir = New_InitialDir
vbwProfiler.vbwProcOut 449
vbwProfiler.vbwExecuteLine 9826
End Property
'***** FILTER
Public Property Get Filter() As String
vbwProfiler.vbwProcIn 450
vbwProfiler.vbwExecuteLine 9827
    Filter = m_Filter
vbwProfiler.vbwProcOut 450
vbwProfiler.vbwExecuteLine 9828
End Property
Public Property Let Filter(ByVal New_Filter As String)
vbwProfiler.vbwProcIn 451
vbwProfiler.vbwExecuteLine 9829
    m_Filter = New_Filter
vbwProfiler.vbwProcOut 451
vbwProfiler.vbwExecuteLine 9830
End Property
'***** FILTER INDEX
Public Property Get FilterIndex() As Integer
vbwProfiler.vbwProcIn 452
vbwProfiler.vbwExecuteLine 9831
    FilterIndex = m_FilterIndex
vbwProfiler.vbwProcOut 452
vbwProfiler.vbwExecuteLine 9832
End Property
Public Property Let FilterIndex(ByVal New_FilterIndex As Integer)
vbwProfiler.vbwProcIn 453
vbwProfiler.vbwExecuteLine 9833
    m_FilterIndex = New_FilterIndex
vbwProfiler.vbwProcOut 453
vbwProfiler.vbwExecuteLine 9834
End Property
'***** FONT NAME
Public Property Get FontName() As String
vbwProfiler.vbwProcIn 454
vbwProfiler.vbwExecuteLine 9835
    FontName = m_FontName
vbwProfiler.vbwProcOut 454
vbwProfiler.vbwExecuteLine 9836
End Property
Public Property Let FontName(ByVal New_FontName As String)
vbwProfiler.vbwProcIn 455
vbwProfiler.vbwExecuteLine 9837
    m_FontName = New_FontName
vbwProfiler.vbwProcOut 455
vbwProfiler.vbwExecuteLine 9838
End Property
'***** FONT SIZE
Public Property Get FontSize() As Integer
vbwProfiler.vbwProcIn 456
vbwProfiler.vbwExecuteLine 9839
    FontSize = m_FontSize
vbwProfiler.vbwProcOut 456
vbwProfiler.vbwExecuteLine 9840
End Property
Public Property Let FontSize(ByVal New_FontSize As Integer)
vbwProfiler.vbwProcIn 457
vbwProfiler.vbwExecuteLine 9841
    m_FontSize = New_FontSize
vbwProfiler.vbwProcOut 457
vbwProfiler.vbwExecuteLine 9842
End Property
'***** FONT COLOR
Public Property Get FontColor() As Long
vbwProfiler.vbwProcIn 458
vbwProfiler.vbwExecuteLine 9843
    FontColor = m_FontColor
vbwProfiler.vbwProcOut 458
vbwProfiler.vbwExecuteLine 9844
End Property
Public Property Let FontColor(ByVal New_FontColor As Long)
vbwProfiler.vbwProcIn 459
vbwProfiler.vbwExecuteLine 9845
    m_FontColor = New_FontColor
vbwProfiler.vbwProcOut 459
vbwProfiler.vbwExecuteLine 9846
End Property
'***** FONT BOLD
Public Property Get FontBold() As Boolean
vbwProfiler.vbwProcIn 460
vbwProfiler.vbwExecuteLine 9847
    FontBold = m_FontBold
vbwProfiler.vbwProcOut 460
vbwProfiler.vbwExecuteLine 9848
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
vbwProfiler.vbwProcIn 461
vbwProfiler.vbwExecuteLine 9849
    m_FontBold = New_FontBold
vbwProfiler.vbwProcOut 461
vbwProfiler.vbwExecuteLine 9850
End Property
'***** FONT ITALIC
Public Property Get FontItalic() As Boolean
vbwProfiler.vbwProcIn 462
vbwProfiler.vbwExecuteLine 9851
    FontItalic = m_FontItalic
vbwProfiler.vbwProcOut 462
vbwProfiler.vbwExecuteLine 9852
End Property
Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
vbwProfiler.vbwProcIn 463
vbwProfiler.vbwExecuteLine 9853
    m_FontItalic = New_FontItalic
vbwProfiler.vbwProcOut 463
vbwProfiler.vbwExecuteLine 9854
End Property
'***** FONT UNDERLINE
Public Property Get FontUnderline() As Boolean
vbwProfiler.vbwProcIn 464
vbwProfiler.vbwExecuteLine 9855
    FontUnderline = m_FontUnderline
vbwProfiler.vbwProcOut 464
vbwProfiler.vbwExecuteLine 9856
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
vbwProfiler.vbwProcIn 465
vbwProfiler.vbwExecuteLine 9857
    m_FontUnderline = New_FontUnderline
vbwProfiler.vbwProcOut 465
vbwProfiler.vbwExecuteLine 9858
End Property
'***** FONT STRIKETHRU
Public Property Get FontStrikeThru() As Boolean
vbwProfiler.vbwProcIn 466
vbwProfiler.vbwExecuteLine 9859
    FontStrikeThru = m_FontStrikeThru
vbwProfiler.vbwProcOut 466
vbwProfiler.vbwExecuteLine 9860
End Property
Public Property Let FontStrikeThru(ByVal New_FontStrikeThru As Boolean)
vbwProfiler.vbwProcIn 467
vbwProfiler.vbwExecuteLine 9861
    m_FontStrikeThru = New_FontStrikeThru
vbwProfiler.vbwProcOut 467
vbwProfiler.vbwExecuteLine 9862
End Property

' Initialize Properties for class
Private Sub Class_Initialize()
vbwProfiler.vbwProcIn 468
vbwProfiler.vbwExecuteLine 9863
    m_CancelError = m_def_CancelError
vbwProfiler.vbwExecuteLine 9864
    m_Filename = m_def_Filename
vbwProfiler.vbwExecuteLine 9865
    m_DialogTitle = m_def_DialogTitle
vbwProfiler.vbwExecuteLine 9866
    m_InitialDir = m_def_InitialDir
vbwProfiler.vbwExecuteLine 9867
    m_Filter = m_def_Filter
vbwProfiler.vbwExecuteLine 9868
    m_FilterIndex = m_def_FilterIndex
vbwProfiler.vbwExecuteLine 9869
    m_MultiSelect = m_def_MultiSelect
vbwProfiler.vbwExecuteLine 9870
    m_FontName = m_def_FontName
vbwProfiler.vbwExecuteLine 9871
    m_FontSize = m_def_FontSize
vbwProfiler.vbwExecuteLine 9872
    m_FontColor = m_def_FontColor
vbwProfiler.vbwExecuteLine 9873
    m_FontBold = m_def_FontBold
vbwProfiler.vbwExecuteLine 9874
    m_FontItalic = m_def_FontItalic
vbwProfiler.vbwExecuteLine 9875
    m_FontUnderline = m_def_FontUnderline
vbwProfiler.vbwExecuteLine 9876
    m_FontStrikeThru = m_def_FontStrikeThru
vbwProfiler.vbwProcOut 468
vbwProfiler.vbwExecuteLine 9877
End Sub



Public Function ShowOpen(ByVal hParent As Long) As Boolean
    '** Description:
    '** Calls open dialog without OCX
vbwProfiler.vbwProcIn 469
    Dim epOFN As OPENFILENAME
    Dim lngRet As Long
vbwProfiler.vbwExecuteLine 9878
    With epOFN

vbwProfiler.vbwExecuteLine 9879
        If MultiSelect Then 'If Multi Select then
vbwProfiler.vbwExecuteLine 9880
            .flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
vbwProfiler.vbwExecuteLine 9881
            .lpstrFile = DefaultFilename & Space(9999 - Len(DefaultFilename)) & vbNullChar
vbwProfiler.vbwExecuteLine 9882
            .lpstrFileTitle = Space(9999) & vbNullChar
        Else
vbwProfiler.vbwExecuteLine 9883 'B
vbwProfiler.vbwExecuteLine 9884
            .flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
vbwProfiler.vbwExecuteLine 9885
            .lpstrFile = DefaultFilename & String(MAX_PATH - Len(DefaultFilename), 0) & vbNullChar
vbwProfiler.vbwExecuteLine 9886
            .lpstrFileTitle = String(MAX_PATH, 0) & vbNullChar
        End If
vbwProfiler.vbwExecuteLine 9887 'B

vbwProfiler.vbwExecuteLine 9888
        .hwndOwner = hParent   'Handle to window
vbwProfiler.vbwExecuteLine 9889
        .lpstrFilter = SetFilter(Filter) & vbNullChar 'File filter
vbwProfiler.vbwExecuteLine 9890
        .lpstrInitialDir = InitialDir & vbNullChar 'Initial directory
vbwProfiler.vbwExecuteLine 9891
        .lpstrTitle = DialogTitle & vbNullChar 'Dialog title
vbwProfiler.vbwExecuteLine 9892
        .lStructSize = Len(epOFN) 'Structure size in bytes
vbwProfiler.vbwExecuteLine 9893
        .nFilterIndex = FilterIndex 'Filter index
vbwProfiler.vbwExecuteLine 9894
        .nMaxFile = Len(.lpstrFile) 'Maximum file length
vbwProfiler.vbwExecuteLine 9895
        .nMaxFileTitle = Len(.lpstrFileTitle) 'Maximum file title length
vbwProfiler.vbwExecuteLine 9896
    End With

vbwProfiler.vbwExecuteLine 9897
    lngRet = GetOpenFileName(epOFN) 'Call open dialog

vbwProfiler.vbwExecuteLine 9898
    If lngRet <> 0 Then 'If there are no errors continue with opening file
vbwProfiler.vbwExecuteLine 9899
        ParseFileName epOFN.lpstrFile
vbwProfiler.vbwExecuteLine 9900
        ShowOpen = True
    Else
vbwProfiler.vbwExecuteLine 9901 'B
vbwProfiler.vbwExecuteLine 9902
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            'err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
vbwProfiler.vbwExecuteLine 9903
            ShowOpen = False
        End If
vbwProfiler.vbwExecuteLine 9904 'B
    End If
vbwProfiler.vbwExecuteLine 9905 'B
vbwProfiler.vbwProcOut 469
vbwProfiler.vbwExecuteLine 9906
End Function

Public Function ShowSave(ByVal hParent As Long) As Long
    '** Description:
    '** Calls save dialog without OCX
vbwProfiler.vbwProcIn 470
    Dim epOFN As OPENFILENAME
    Dim lngRet As Long
vbwProfiler.vbwExecuteLine 9907
    With epOFN
vbwProfiler.vbwExecuteLine 9908
        .hwndOwner = hParent  'Handle to parent window
vbwProfiler.vbwExecuteLine 9909
        .flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
vbwProfiler.vbwExecuteLine 9910
        .lpstrFile = DefaultFilename & String(MAX_PATH - Len(DefaultFilename), 0) & vbNullChar
vbwProfiler.vbwExecuteLine 9911
        .lpstrFileTitle = String(MAX_PATH, 0) & vbNullChar
vbwProfiler.vbwExecuteLine 9912
        .lpstrFilter = SetFilter(Filter) & vbNullChar 'File filter
vbwProfiler.vbwExecuteLine 9913
        .lpstrInitialDir = InitialDir & vbNullChar 'Initial directory
vbwProfiler.vbwExecuteLine 9914
        .lpstrTitle = DialogTitle & vbNullChar 'Dialog title
vbwProfiler.vbwExecuteLine 9915
        .lStructSize = Len(epOFN) 'Structure size in bytes
vbwProfiler.vbwExecuteLine 9916
        .nFilterIndex = FilterIndex 'Filter index
vbwProfiler.vbwExecuteLine 9917
        .nMaxFile = Len(.lpstrFile) 'Maximum file length
vbwProfiler.vbwExecuteLine 9918
        .nMaxFileTitle = Len(.lpstrFileTitle) 'Maximum file title length
vbwProfiler.vbwExecuteLine 9919
    End With

vbwProfiler.vbwExecuteLine 9920
    lngRet = GetSaveFileName(epOFN) 'Call save dialog

vbwProfiler.vbwExecuteLine 9921
    If lngRet <> 0 Then 'If there are no errors continue with saving file
vbwProfiler.vbwExecuteLine 9922
        ParseFileName epOFN.lpstrFile
vbwProfiler.vbwExecuteLine 9923
        ShowSave = False
    Else
vbwProfiler.vbwExecuteLine 9924 'B
vbwProfiler.vbwExecuteLine 9925
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            'err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
vbwProfiler.vbwExecuteLine 9926
            ShowSave = True
        End If
vbwProfiler.vbwExecuteLine 9927 'B
    End If
vbwProfiler.vbwExecuteLine 9928 'B
vbwProfiler.vbwProcOut 470
vbwProfiler.vbwExecuteLine 9929
End Function

Public Function ShowFont(ByVal hParent As Long) As Long
    '** Description:
    '** Call font dialog without OCX
vbwProfiler.vbwProcIn 471
    Dim CF As CHOOSEFONT
    Dim lf As LOGFONT
    Dim lMemHandle As Long
    Dim lLogFont As Long
    Dim lngRet As Long

vbwProfiler.vbwExecuteLine 9930
    With lf
vbwProfiler.vbwExecuteLine 9931
        .lfCharSet = DEFAULT_CHARSET 'Default character set
vbwProfiler.vbwExecuteLine 9932
        .lfClipPrecision = CLIP_DEFAULT_PRECIS 'Clipping precision
vbwProfiler.vbwExecuteLine 9933
        .lfFaceName = "Arial" & vbNullChar 'Font name
vbwProfiler.vbwExecuteLine 9934
        .lfHeight = 13 'Height
vbwProfiler.vbwExecuteLine 9935
        .lfOutPrecision = OUT_DEFAULT_PRECIS 'Precision mapping
vbwProfiler.vbwExecuteLine 9936
        .lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN 'Default pitch
vbwProfiler.vbwExecuteLine 9937
        .lfQuality = DEFAULT_QUALITY 'Default quality
vbwProfiler.vbwExecuteLine 9938
        .lfWeight = FW_NORMAL 'Regular font type
vbwProfiler.vbwExecuteLine 9939
    End With

    ' Create the memory block
vbwProfiler.vbwExecuteLine 9940
    lMemHandle = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lf))
vbwProfiler.vbwExecuteLine 9941
    lLogFont = GlobalLock(lMemHandle)
vbwProfiler.vbwExecuteLine 9942
    CopyMemory ByVal lLogFont, lf, Len(lf)

vbwProfiler.vbwExecuteLine 9943
    With CF
vbwProfiler.vbwExecuteLine 9944
        .flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
vbwProfiler.vbwExecuteLine 9945
        .hdc = Printer.hdc 'Device context of default printer
vbwProfiler.vbwExecuteLine 9946
        .hwndOwner = hParent  'Handle to window
vbwProfiler.vbwExecuteLine 9947
        .iPointSize = 120 'Set font size to 12 size
vbwProfiler.vbwExecuteLine 9948
        .lpLogFont = lLogFont 'Log font
vbwProfiler.vbwExecuteLine 9949
        .lStructSize = Len(CF) 'Size of structure in bytes
vbwProfiler.vbwExecuteLine 9950
        .nFontType = REGULAR_FONTTYPE 'Regular font type
vbwProfiler.vbwExecuteLine 9951
        .nSizeMax = 72 'Maximum font size
vbwProfiler.vbwExecuteLine 9952
        .nSizeMin = 10 'Minimum font size
vbwProfiler.vbwExecuteLine 9953
        .rgbColors = RGB(0, 0, 0) 'Font color
vbwProfiler.vbwExecuteLine 9954
    End With

vbwProfiler.vbwExecuteLine 9955
    lngRet = CHOOSEFONT(CF) 'Call font dialog
vbwProfiler.vbwExecuteLine 9956
    If lngRet <> 0 Then 'If there are no errors continue with font
vbwProfiler.vbwExecuteLine 9957
        CopyMemory lf, ByVal lLogFont, Len(lf)

vbwProfiler.vbwExecuteLine 9958
        FontName = Left(lf.lfFaceName, InStr(lf.lfFaceName, vbNullChar) - 1)
vbwProfiler.vbwExecuteLine 9959
        FontSize = CF.iPointSize / 10
vbwProfiler.vbwExecuteLine 9960
        FontColor = CF.rgbColors
vbwProfiler.vbwExecuteLine 9961
        If lf.lfWeight = FW_NORMAL Then
vbwProfiler.vbwExecuteLine 9962
            FontBold = False
vbwProfiler.vbwExecuteLine 9963
            FontItalic = False
vbwProfiler.vbwExecuteLine 9964
            FontUnderline = False
vbwProfiler.vbwExecuteLine 9965
            FontStrikeThru = False
        Else
vbwProfiler.vbwExecuteLine 9966 'B
vbwProfiler.vbwExecuteLine 9967
            If lf.lfWeight = FW_BOLD Then
vbwProfiler.vbwExecuteLine 9968
                 FontBold = True
            End If
vbwProfiler.vbwExecuteLine 9969 'B
vbwProfiler.vbwExecuteLine 9970
            If lf.lfItalic <> 0 Then
vbwProfiler.vbwExecuteLine 9971
                 FontItalic = True
            End If
vbwProfiler.vbwExecuteLine 9972 'B
vbwProfiler.vbwExecuteLine 9973
            If lf.lfUnderline <> 0 Then
vbwProfiler.vbwExecuteLine 9974
                 FontUnderline = True
            End If
vbwProfiler.vbwExecuteLine 9975 'B
vbwProfiler.vbwExecuteLine 9976
            If lf.lfStrikeout <> 0 Then
vbwProfiler.vbwExecuteLine 9977
                 FontStrikeThru = True
            End If
vbwProfiler.vbwExecuteLine 9978 'B
        End If
vbwProfiler.vbwExecuteLine 9979 'B
vbwProfiler.vbwExecuteLine 9980
        ShowFont = False
    Else
vbwProfiler.vbwExecuteLine 9981 'B
vbwProfiler.vbwExecuteLine 9982
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
vbwProfiler.vbwExecuteLine 9983
            err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
vbwProfiler.vbwExecuteLine 9984
            ShowFont = True
        End If
vbwProfiler.vbwExecuteLine 9985 'B
    End If
vbwProfiler.vbwExecuteLine 9986 'B

    ' Unlock and free the memory block
    ' Note this must be done
vbwProfiler.vbwExecuteLine 9987
    GlobalUnlock lMemHandle
vbwProfiler.vbwExecuteLine 9988
    GlobalFree lMemHandle
vbwProfiler.vbwProcOut 471
vbwProfiler.vbwExecuteLine 9989
End Function

Public Function ShowColor(ByVal hParent As Long) As Long
    '** Description:
    '** Call color dialog without OCX
vbwProfiler.vbwProcIn 472
    Dim epCC As ChooseColor
    Dim lngRet As Long
    Dim CusCol(0 To 16) As Long
    Dim i As Integer

    ' Fills custom colors with white
vbwProfiler.vbwExecuteLine 9990
    For i = 0 To 15
vbwProfiler.vbwExecuteLine 9991
        CusCol(i) = vbWhite
vbwProfiler.vbwExecuteLine 9992
    Next

vbwProfiler.vbwExecuteLine 9993
    With epCC
vbwProfiler.vbwExecuteLine 9994
        .hwndOwner = hParent  'Handle to owner
vbwProfiler.vbwExecuteLine 9995
        .lStructSize = Len(epCC) 'Structure size in bytes
vbwProfiler.vbwExecuteLine 9996
        .lpCustColors = VarPtr(CusCol(0)) 'Custom colors
vbwProfiler.vbwExecuteLine 9997
        .rgbResult = 0 'RGB result
vbwProfiler.vbwExecuteLine 9998
    End With

vbwProfiler.vbwExecuteLine 9999
    lngRet = ChooseColor(epCC) 'Call color dialog
vbwProfiler.vbwExecuteLine 10000
    If lngRet <> 0 Then 'If there are no errors continue with color
vbwProfiler.vbwExecuteLine 10001
        ShowColor = epCC.rgbResult
    Else
vbwProfiler.vbwExecuteLine 10002 'B
vbwProfiler.vbwExecuteLine 10003
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            'err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
vbwProfiler.vbwExecuteLine 10004
            ShowColor = True
        End If
vbwProfiler.vbwExecuteLine 10005 'B
    End If
vbwProfiler.vbwExecuteLine 10006 'B
vbwProfiler.vbwProcOut 472
vbwProfiler.vbwExecuteLine 10007
End Function

Public Function ShowPageSetup(ByVal hParent As Long) As Long
    '** Description:
    '** Call page setup dialog without OCX
vbwProfiler.vbwProcIn 473
    Dim epPSD As PageSetupDlg
    Dim lngRet As Long

vbwProfiler.vbwExecuteLine 10008
    epPSD.lStructSize = Len(epPSD) 'Structure size in bytes
vbwProfiler.vbwExecuteLine 10009
    epPSD.hwndOwner = hParent

vbwProfiler.vbwExecuteLine 10010
    lngRet = PageSetupDlg(epPSD) 'Call page setup dialog
vbwProfiler.vbwExecuteLine 10011
    If lngRet <> 0 Then 'If there are no errors continue
vbwProfiler.vbwExecuteLine 10012
        ShowPageSetup = False
    Else
vbwProfiler.vbwExecuteLine 10013 'B
vbwProfiler.vbwExecuteLine 10014
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
           ' err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
vbwProfiler.vbwExecuteLine 10015
            ShowPageSetup = True
        End If
vbwProfiler.vbwExecuteLine 10016 'B
    End If
vbwProfiler.vbwExecuteLine 10017 'B
vbwProfiler.vbwProcOut 473
vbwProfiler.vbwExecuteLine 10018
End Function

Public Function ShowPrinter(ByVal hParent As Long) As Long
    '** Description:
    '** Call printer dialog without OCX
    '**
    '** Note:
    '** This is not my function it's from KPD-Team 1998 URL: http://www.allapi.net
    '** and i have modified it a little
    '-> Code by Donald Grover
vbwProfiler.vbwProcIn 474
    Dim PrintDlg As PRINTDLG_TYPE
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE

    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String

    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName structures

vbwProfiler.vbwExecuteLine 10019
    PrintDlg.lStructSize = Len(PrintDlg)
vbwProfiler.vbwExecuteLine 10020
    PrintDlg.hwndOwner = hParent  'Handle to window

vbwProfiler.vbwExecuteLine 10021
    On Error Resume Next
    'Set the current orientation and duplex setting
vbwProfiler.vbwExecuteLine 10022
    DevMode.dmDeviceName = Printer.DeviceName
vbwProfiler.vbwExecuteLine 10023
    DevMode.dmSize = Len(DevMode)
vbwProfiler.vbwExecuteLine 10024
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
vbwProfiler.vbwExecuteLine 10025
    DevMode.dmPaperWidth = Printer.Width
vbwProfiler.vbwExecuteLine 10026
    DevMode.dmOrientation = Printer.Orientation
vbwProfiler.vbwExecuteLine 10027
    DevMode.dmPaperSize = Printer.PaperSize
vbwProfiler.vbwExecuteLine 10028
    DevMode.dmDuplex = Printer.Duplex
vbwProfiler.vbwExecuteLine 10029
    On Error GoTo 0

    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
vbwProfiler.vbwExecuteLine 10030
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
vbwProfiler.vbwExecuteLine 10031
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
vbwProfiler.vbwExecuteLine 10032
    If lpDevMode > 0 Then
vbwProfiler.vbwExecuteLine 10033
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
vbwProfiler.vbwExecuteLine 10034
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If
vbwProfiler.vbwExecuteLine 10035 'B

    'Set the current driver, device, and port name strings
vbwProfiler.vbwExecuteLine 10036
    With DevName
vbwProfiler.vbwExecuteLine 10037
        .wDriverOffset = 8
vbwProfiler.vbwExecuteLine 10038
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
vbwProfiler.vbwExecuteLine 10039
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
vbwProfiler.vbwExecuteLine 10040
        .wDefault = 0
vbwProfiler.vbwExecuteLine 10041
    End With

vbwProfiler.vbwExecuteLine 10042
    With Printer
vbwProfiler.vbwExecuteLine 10043
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
vbwProfiler.vbwExecuteLine 10044
    End With

    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
vbwProfiler.vbwExecuteLine 10045
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
vbwProfiler.vbwExecuteLine 10046
    lpDevName = GlobalLock(PrintDlg.hDevNames)
vbwProfiler.vbwExecuteLine 10047
    If lpDevName > 0 Then
vbwProfiler.vbwExecuteLine 10048
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
vbwProfiler.vbwExecuteLine 10049
        bReturn = GlobalUnlock(lpDevName)
    End If
vbwProfiler.vbwExecuteLine 10050 'B

    'Call the print dialog up and let the user make changes
vbwProfiler.vbwExecuteLine 10051
    If PrintDialog(PrintDlg) <> 0 Then

        'First get the DevName structure.
vbwProfiler.vbwExecuteLine 10052
        lpDevName = GlobalLock(PrintDlg.hDevNames)
vbwProfiler.vbwExecuteLine 10053
        CopyMemory DevName, ByVal lpDevName, 45
vbwProfiler.vbwExecuteLine 10054
        bReturn = GlobalUnlock(lpDevName)
vbwProfiler.vbwExecuteLine 10055
        GlobalFree PrintDlg.hDevNames

        'Next get the DevMode structure and set the printer
        'properties appropriately
vbwProfiler.vbwExecuteLine 10056
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
vbwProfiler.vbwExecuteLine 10057
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
vbwProfiler.vbwExecuteLine 10058
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
vbwProfiler.vbwExecuteLine 10059
        GlobalFree PrintDlg.hDevMode
vbwProfiler.vbwExecuteLine 10060
        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
vbwProfiler.vbwExecuteLine 10061
        If Printer.DeviceName <> NewPrinterName Then
vbwProfiler.vbwExecuteLine 10062
            For Each objPrinter In Printers
vbwProfiler.vbwExecuteLine 10063
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
vbwProfiler.vbwExecuteLine 10064
                    Set Printer = objPrinter
                    'set printer toolbar name at this point
                End If
vbwProfiler.vbwExecuteLine 10065 'B
vbwProfiler.vbwExecuteLine 10066
            Next
        End If
vbwProfiler.vbwExecuteLine 10067 'B

vbwProfiler.vbwExecuteLine 10068
        On Error Resume Next
        'Set printer object properties according to selections made
        'by user
vbwProfiler.vbwExecuteLine 10069
        Printer.Copies = DevMode.dmCopies
vbwProfiler.vbwExecuteLine 10070
        Printer.Duplex = DevMode.dmDuplex
vbwProfiler.vbwExecuteLine 10071
        Printer.Orientation = DevMode.dmOrientation
vbwProfiler.vbwExecuteLine 10072
        Printer.PaperSize = DevMode.dmPaperSize
vbwProfiler.vbwExecuteLine 10073
        Printer.PrintQuality = DevMode.dmPrintQuality
vbwProfiler.vbwExecuteLine 10074
        Printer.ColorMode = DevMode.dmColor
vbwProfiler.vbwExecuteLine 10075
        Printer.PaperBin = DevMode.dmDefaultSource
vbwProfiler.vbwExecuteLine 10076
        ShowPrinter = False
vbwProfiler.vbwExecuteLine 10077
        On Error GoTo 0
    Else
vbwProfiler.vbwExecuteLine 10078 'B
vbwProfiler.vbwExecuteLine 10079
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
           ' err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
vbwProfiler.vbwExecuteLine 10080
            ShowPrinter = True
        End If
vbwProfiler.vbwExecuteLine 10081 'B
    End If
vbwProfiler.vbwExecuteLine 10082 'B
vbwProfiler.vbwProcOut 474
vbwProfiler.vbwExecuteLine 10083
End Function

Private Function ParseFileName(sFileName As String)
    '** Description:
    '** Remove null chars from filename and parse multi filename
    '**
    '** Syntax:
    '** szFilename = ParseFileName(strFilename)
    '**
    '** Example:
    '** szFilename = ParseFileName("C:\Autoexec.bat||")
vbwProfiler.vbwProcIn 475
    Dim i As Long
    Dim sPath As String
    Dim sFiles() As String
    Dim Pos As Integer
    Dim sFile As String
    Dim sFileTitle As String

    ' Create new collections
vbwProfiler.vbwExecuteLine 10084
    Set cFileName = New Collection
vbwProfiler.vbwExecuteLine 10085
    Set cFileTitle = New Collection
    ' Found position of two last null chars
vbwProfiler.vbwExecuteLine 10086
    Pos = InStr(sFileName, vbNullChar & vbNullChar)
    ' Remove from filename last two chars
vbwProfiler.vbwExecuteLine 10087
    sFile = Left(sFileName, Pos - 1)

    ' Check to see if filename is single or multi
vbwProfiler.vbwExecuteLine 10088
    If InStr(1, sFile, vbNullChar) <> 0 Then
    ' Multi file
vbwProfiler.vbwExecuteLine 10089
        sFile = Left(sFileName, Pos) & vbNullChar 'Add null char at end of filename
vbwProfiler.vbwExecuteLine 10090
        sPath = Left(sFileName, InStr(1, sFileName, Chr(0)) - 1) 'Get file path
vbwProfiler.vbwExecuteLine 10091
        sFiles = Split(sFile, Chr(0)) 'Split file where is nullchar

        ' Add all filenames to collection
vbwProfiler.vbwExecuteLine 10092
        For i = LBound(sFiles) To UBound(sFiles) - 2
            ' If path doesent contain separator then add it
vbwProfiler.vbwExecuteLine 10093
            If Right(sPath, 1) = "\" Then
vbwProfiler.vbwExecuteLine 10094
                cFileName.Add sPath & sFiles(i)
            Else
vbwProfiler.vbwExecuteLine 10095 'B
vbwProfiler.vbwExecuteLine 10096
                cFileName.Add sPath & "\" & sFiles(i)
            End If
vbwProfiler.vbwExecuteLine 10097 'B
            ' Add file title
vbwProfiler.vbwExecuteLine 10098
            cFileTitle.Add sFiles(i)
            ' Remove first item from collections
vbwProfiler.vbwExecuteLine 10099
            If i = 1 Then
vbwProfiler.vbwExecuteLine 10100
                 cFileName.Remove 1
vbwProfiler.vbwExecuteLine 10101
                 cFileTitle.Remove 1
            End If
vbwProfiler.vbwExecuteLine 10102 'B
vbwProfiler.vbwExecuteLine 10103
        Next
    Else ' Single file
vbwProfiler.vbwExecuteLine 10104 'B
        'Add file name to collection
vbwProfiler.vbwExecuteLine 10105
        cFileName.Add sFile
        ' Add file title
vbwProfiler.vbwExecuteLine 10106
        cFileTitle.Add Right(sFile, Len(sFile) - InStrRev(sFile, "\"))
    End If
vbwProfiler.vbwExecuteLine 10107 'B
vbwProfiler.vbwProcOut 475
vbwProfiler.vbwExecuteLine 10108
End Function

Private Function SetFilter(sFlt As String) As String
    '** Description:
    '** Replace "|" with Null Character
    '**
    '** Syntax:
    '** szFilter = SetFilter(strFilter)
    '**
    '** Example:
    '** szFilter = SetFilter("Text Files (*.txt)|*.txt|All Files |*.*|")
vbwProfiler.vbwProcIn 476
    Dim sLen As Long
    Dim Pos As Long

vbwProfiler.vbwExecuteLine 10109
    sLen = Len(sFlt) 'Get filter length
vbwProfiler.vbwExecuteLine 10110
    Pos = InStr(1, sFlt, "|") 'Find first position of "|"

    ' Loop while Pos > 0
'vbwLine 10111:    While Pos > 0
    While vbwProfiler.vbwExecuteLine(10111) Or Pos > 0
        ' Replace "|" with null char
vbwProfiler.vbwExecuteLine 10112
        sFlt = Left(sFlt, Pos - 1) & vbNullChar & Mid(sFlt, Pos + 1, sLen - Pos)
        ' Find next position of "|"
vbwProfiler.vbwExecuteLine 10113
        Pos = InStr(Pos + 1, sFlt, "|")
vbwProfiler.vbwExecuteLine 10114
    Wend
vbwProfiler.vbwExecuteLine 10115
    SetFilter = sFlt ' Set filter
vbwProfiler.vbwProcOut 476
vbwProfiler.vbwExecuteLine 10116
End Function



