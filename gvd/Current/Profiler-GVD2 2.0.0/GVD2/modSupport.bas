Attribute VB_Name = "modSupport"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long

' pen styles
Public Const PS_SOLID = 0

' Color Types
Public Const CTLCOLOR_MSGBOX = 0
Public Const CTLCOLOR_EDIT = 1
Public Const CTLCOLOR_LISTBOX = 2
Public Const CTLCOLOR_BTN = 3
Public Const CTLCOLOR_DLG = 4
Public Const CTLCOLOR_SCROLLBAR = 5
Public Const CTLCOLOR_STATIC = 6
Public Const CTLCOLOR_MAX = 8                                  '  three bits max

Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_MSGBOX = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_MSGBOXTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20

Public Const COLOR_3DDKSHADOW = 21
Public Const COLOR_3DLIGHT = 22
Public Const COLOR_INFOTEXT = 23
Public Const COLOR_INFOBK = 24

Public Const COLOR_DESKTOP = COLOR_BACKGROUND
Public Const COLOR_3DFACE = COLOR_BTNFACE
Public Const COLOR_3DSHADOW = COLOR_BTNSHADOW
Public Const COLOR_3DHIGHLIGHT = COLOR_BTNHIGHLIGHT
Public Const COLOR_3DHILIGHT = COLOR_BTNHIGHLIGHT
Public Const COLOR_BTNHILIGHT = COLOR_BTNHIGHLIGHT

Public Type SIZE
        cx As Long
        cy As Long
End Type

Public Type TEXTMETRIC
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

Public Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type




'Add a units items to the Custom Power cells selection. Instead of just
'putting in the power as kWs, allow selecting kWs, MWs (megawatts), GWs
'(Gigawatts), TWs (terrawatts).  This makes sense, some users dont
' want to have these huge numbers which dont even fit on screen properly.

Public Enum EXPRESSION_TYPE
    CONDITION = 0
    EXCEPTION = 1
End Enum

Public Enum EVALUATOR_TYPE
    EQUAL_TO = 0
    GREATER_THAN = 1
    GREATER_THAN_OR_EQUAL_TO = 2
    LESS_THAN = 3
    LESS_THAN_EQUAL_TO = 4
    NOT_EQUAL_TO = 5
End Enum

Public Enum ROUND_OPTION
    ROUND_NONE = 0
    DECIMAL_PLACES = 1
    SCIENTIFIC = 2
End Enum

Public Const EXPRESSION_EQUAL_TO = "if equal to "
Public Const EXPRESSION_GREATER_THAN = "if greater than "
Public Const EXPRESSION_GREATER_THAN_OR_EQUAL_TO = "if greater than or equal to "
Public Const EXPRESSION_LESS_THAN = "if less than "
Public Const EXPRESSION_LESS_THAN_OR_EQUAL_TO = "if less than or equal to "
Public Const EXPRESSION_NOT_EQUAL_TO = "if not equal to "

' hotspots are either going to be negative value or 0+ to indicate an expression index
Public Const HS_CONVERSION = -1
Public Const HS_ROUND = -2

' rounding & sig digit constants
Public Const MAX_ROUND = 6
Public Const MAX_SIGNIFICANT_DIGITS_1 = 100

Public Const ROUND_NUMBER_OF_PLACES = "Number of places after the decimal point to round to = "
Public Const ROUND_NUMBER_SIG_DIGITS = "Numer of significant digits to display = "


' since we are using rtlmovememory on this UDT, its important to remember that VB will
' dword align every element that is less than 4 bytes (e.g. an integer or byte) but NOT if its
' the last element in the UDT.  So we will keep the uExpression.Type element last.  Our lenB(uExpression)
' will be = 9 bytes as expected.
Public Type uExpression
    value As Double
    evaluator As EVALUATOR_TYPE
    type As EXPRESSION_TYPE
End Type

Private Function getEvaluatorStringFromCode(ByVal eType As EVALUATOR_TYPE) As String
vbwProfiler.vbwProcIn 630
    Dim s As String

vbwProfiler.vbwExecuteLine 11313
    Select Case eType
'vbwLine 11314:        Case EQUAL_TO
        Case IIf(vbwProfiler.vbwExecuteLine(11314), VBWPROFILER_EMPTY, _
        EQUAL_TO)
vbwProfiler.vbwExecuteLine 11315
            s = EXPRESSION_EQUAL_TO
'vbwLine 11316:        Case GREATER_THAN
        Case IIf(vbwProfiler.vbwExecuteLine(11316), VBWPROFILER_EMPTY, _
        GREATER_THAN)
vbwProfiler.vbwExecuteLine 11317
            s = EXPRESSION_GREATER_THAN
'vbwLine 11318:        Case GREATER_THAN_OR_EQUAL_TO
        Case IIf(vbwProfiler.vbwExecuteLine(11318), VBWPROFILER_EMPTY, _
        GREATER_THAN_OR_EQUAL_TO)
vbwProfiler.vbwExecuteLine 11319
            s = EXPRESSION_GREATER_THAN_OR_EQUAL_TO
'vbwLine 11320:        Case LESS_THAN
        Case IIf(vbwProfiler.vbwExecuteLine(11320), VBWPROFILER_EMPTY, _
        LESS_THAN)
vbwProfiler.vbwExecuteLine 11321
            s = EXPRESSION_LESS_THAN
'vbwLine 11322:        Case LESS_THAN_EQUAL_TO
        Case IIf(vbwProfiler.vbwExecuteLine(11322), VBWPROFILER_EMPTY, _
        LESS_THAN_EQUAL_TO)
vbwProfiler.vbwExecuteLine 11323
            s = EXPRESSION_LESS_THAN_OR_EQUAL_TO
'vbwLine 11324:        Case NOT_EQUAL_TO
        Case IIf(vbwProfiler.vbwExecuteLine(11324), VBWPROFILER_EMPTY, _
        NOT_EQUAL_TO)
vbwProfiler.vbwExecuteLine 11325
            s = EXPRESSION_NOT_EQUAL_TO
    End Select
vbwProfiler.vbwExecuteLine 11326 'B

vbwProfiler.vbwExecuteLine 11327
    getEvaluatorStringFromCode = s
vbwProfiler.vbwProcOut 630
vbwProfiler.vbwExecuteLine 11328
End Function

Public Sub renderRule(ByRef oRule As cRule, ByRef oLB As cCustomListBox)
vbwProfiler.vbwProcIn 631
    Dim uExp As uExpression
    Dim arrExpressions() As uExpression
    Dim lngCount As Long
    Dim i As Long
    Dim j As Long
    Dim oConvert As cUnitConverter
    Dim s As String
    Dim lngNextSpot As Long
    Dim index As Long

vbwProfiler.vbwExecuteLine 11329
    Set oConvert = New cUnitConverter

    ' load up all our expressions from our rules objects
    ' we dont need to sort them... just copy them over locally
vbwProfiler.vbwExecuteLine 11330
    If oRule Is Nothing Then
vbwProfiler.vbwProcOut 631
vbwProfiler.vbwExecuteLine 11331
         Exit Sub
    End If
vbwProfiler.vbwExecuteLine 11332 'B
vbwProfiler.vbwExecuteLine 11333
    With oRule
vbwProfiler.vbwExecuteLine 11334
        lngCount = .expressionCount
vbwProfiler.vbwExecuteLine 11335
        If lngCount >= 1 Then
vbwProfiler.vbwExecuteLine 11336
            ReDim arrExpressions(0 To lngCount - 1)
vbwProfiler.vbwExecuteLine 11337
            For i = 0 To lngCount - 1
vbwProfiler.vbwExecuteLine 11338
                CopyMemory uExp, ByVal .getExpression(i), LenB(uExp)
vbwProfiler.vbwExecuteLine 11339
                arrExpressions(i) = uExp
vbwProfiler.vbwExecuteLine 11340
            Next
        End If
vbwProfiler.vbwExecuteLine 11341 'B
vbwProfiler.vbwExecuteLine 11342
    End With

    ' so the size of our hotspot array needs to be lngCount + 1 for convert and +1 IF round is set
vbwProfiler.vbwExecuteLine 11343
    lngCount = oRule.expressionCount + 1
vbwProfiler.vbwExecuteLine 11344
    If oRule.RoundType > ROUND_NONE Then
vbwProfiler.vbwExecuteLine 11345
        lngCount = lngCount + 1
    End If
vbwProfiler.vbwExecuteLine 11346 'B

   ' we have access to all we need. Lets create our lines of text and our hotspots
vbwProfiler.vbwExecuteLine 11347
    oLB.RemoveAllItems
    'NOTE: The <u></u> tags define which parts of the line become a "hotspot".  LIMIT only one hotspot per line tho!
    's = "Convert " & oConvert.unitDescription(oRule.convertFrom) & " to <u>" & oConvert.unitDescription(oRule.convertTo) & "</u>"  '<-- add unit type to the rule, will make it easier to call function to convert code to string"
    'mpj 09/03/2003 -- decided to not allow user to modify the convert To or convert From via hotspot, instead they must go into "modify" wizard
vbwProfiler.vbwExecuteLine 11348
    s = "Convert " & oConvert.unitDescription(oRule.convertFrom) & " to " & oConvert.unitDescription(oRule.convertTo)

vbwProfiler.vbwExecuteLine 11349
    index = oLB.Addline(s, False)
vbwProfiler.vbwExecuteLine 11350
    oLB.setItemData index, HS_CONVERSION

vbwProfiler.vbwExecuteLine 11351
    lngNextSpot = 2
vbwProfiler.vbwExecuteLine 11352
    lngCount = oRule.expressionCount

vbwProfiler.vbwExecuteLine 11353
    s = ""

    Dim bFirstCondition As Boolean
    Dim bFirstException As Boolean
vbwProfiler.vbwExecuteLine 11354
    bFirstCondition = True
vbwProfiler.vbwExecuteLine 11355
    bFirstException = True
    ' add the expressions in order with conditions first,then all the exceptions
vbwProfiler.vbwExecuteLine 11356
    For j = CONDITION To EXCEPTION
vbwProfiler.vbwExecuteLine 11357
        For i = 0 To lngCount - 1
vbwProfiler.vbwExecuteLine 11358
            If arrExpressions(i).type = j Then

vbwProfiler.vbwExecuteLine 11359
                If j = CONDITION Then
vbwProfiler.vbwExecuteLine 11360
                    If bFirstCondition Then
vbwProfiler.vbwExecuteLine 11361
                       bFirstCondition = False
                    Else
vbwProfiler.vbwExecuteLine 11362 'B
vbwProfiler.vbwExecuteLine 11363
                        s = "   and "
                    End If
vbwProfiler.vbwExecuteLine 11364 'B
                Else
vbwProfiler.vbwExecuteLine 11365 'B
vbwProfiler.vbwExecuteLine 11366
                    If bFirstException Then
vbwProfiler.vbwExecuteLine 11367
                        s = "except "
vbwProfiler.vbwExecuteLine 11368
                        bFirstException = False
                    Else
vbwProfiler.vbwExecuteLine 11369 'B
vbwProfiler.vbwExecuteLine 11370
                        s = "   or "
                    End If
vbwProfiler.vbwExecuteLine 11371 'B
                End If
vbwProfiler.vbwExecuteLine 11372 'B

vbwProfiler.vbwExecuteLine 11373
                s = s & getEvaluatorStringFromCode(arrExpressions(i).evaluator) & "<u>" & arrExpressions(i).value & "</u>"
vbwProfiler.vbwExecuteLine 11374
                index = oLB.Addline(s, False)
vbwProfiler.vbwExecuteLine 11375
                oLB.setItemData index, i
vbwProfiler.vbwExecuteLine 11376
                lngNextSpot = lngNextSpot + 1
            End If
vbwProfiler.vbwExecuteLine 11377 'B
vbwProfiler.vbwExecuteLine 11378
        Next
vbwProfiler.vbwExecuteLine 11379
    Next

    ' the formatting options
vbwProfiler.vbwExecuteLine 11380
    Select Case oRule.RoundType
'vbwLine 11381:        Case ROUND_NONE
        Case IIf(vbwProfiler.vbwExecuteLine(11381), VBWPROFILER_EMPTY, _
        ROUND_NONE)
            ' do nothing
'vbwLine 11382:        Case DECIMAL_PLACES
        Case IIf(vbwProfiler.vbwExecuteLine(11382), VBWPROFILER_EMPTY, _
        DECIMAL_PLACES)
vbwProfiler.vbwExecuteLine 11383
            s = "Round result to <u>" & oRule.roundDigits & "</u> digits"
vbwProfiler.vbwExecuteLine 11384
            index = oLB.Addline(s, False)
vbwProfiler.vbwExecuteLine 11385
            oLB.setItemData index, HS_ROUND
'vbwLine 11386:        Case SCIENTIFIC
        Case IIf(vbwProfiler.vbwExecuteLine(11386), VBWPROFILER_EMPTY, _
        SCIENTIFIC)
vbwProfiler.vbwExecuteLine 11387
            s = "Display result in scientific notation"
vbwProfiler.vbwExecuteLine 11388
            index = oLB.Addline(s, False)
vbwProfiler.vbwExecuteLine 11389
            oLB.setItemData index, HS_ROUND
        Case Else
vbwProfiler.vbwExecuteLine 11390 'B
vbwProfiler.vbwExecuteLine 11391
            Debug.Print "modSupport:renderRule() -- Invalid Round Option Enum"
    End Select
vbwProfiler.vbwExecuteLine 11392 'B

    ' thousand seperators setting
vbwProfiler.vbwExecuteLine 11393
    If oRule.useThousandSeperators = True Then
vbwProfiler.vbwExecuteLine 11394
        s = "Use thousand seperators."
vbwProfiler.vbwExecuteLine 11395
        index = oLB.Addline(s, False)
vbwProfiler.vbwExecuteLine 11396
        oLB.setItemData index, HS_ROUND 'HS_ROUND just tells it that its not an expression
    End If
vbwProfiler.vbwExecuteLine 11397 'B

    ' finally "append prefix" setting
vbwProfiler.vbwExecuteLine 11398
    If oRule.appendPostfix = True Then
vbwProfiler.vbwExecuteLine 11399
        s = oConvert.unitAbbrev(oRule.convertTo)
vbwProfiler.vbwExecuteLine 11400
        If s <> "" Then
vbwProfiler.vbwExecuteLine 11401
            s = "Append suffix '" & s & "'."
vbwProfiler.vbwExecuteLine 11402
            index = oLB.Addline(s, False)
vbwProfiler.vbwExecuteLine 11403
            oLB.setItemData index, HS_ROUND 'HS_ROUND just tells it that its not an expression
        End If
vbwProfiler.vbwExecuteLine 11404 'B
    Else
vbwProfiler.vbwExecuteLine 11405 'B
vbwProfiler.vbwExecuteLine 11406
        s = "Do not append suffix."
vbwProfiler.vbwExecuteLine 11407
        index = oLB.Addline(s, False)
vbwProfiler.vbwExecuteLine 11408
        oLB.setItemData index, HS_ROUND
    End If
vbwProfiler.vbwExecuteLine 11409 'B

vbwProfiler.vbwExecuteLine 11410
    oLB.RenderText
vbwProfiler.vbwExecuteLine 11411
    Set oConvert = Nothing
vbwProfiler.vbwProcOut 631
vbwProfiler.vbwExecuteLine 11412
End Sub

Public Sub displayItemClick(ByVal lngItem As Long, ByVal hObject As Long, oLB As cCustomListBox)
vbwProfiler.vbwProcIn 632
    Dim oRule As cRule
    Dim oNode As cINode

vbwProfiler.vbwExecuteLine 11413
    On Error Resume Next '<-- ensures that we can release our object ptr if an error results from user input in the InputBox()
    ' get a reference to the current node and make sure its a "cRule" type
vbwProfiler.vbwExecuteLine 11414
    CopyMemory oNode, hObject, 4
vbwProfiler.vbwExecuteLine 11415
    Set oRule = oNode
vbwProfiler.vbwExecuteLine 11416
    If oNode.Classname = "cRule" Then
vbwProfiler.vbwExecuteLine 11417
        Debug.Print "Hotpot " & lngItem & " clicked"
vbwProfiler.vbwExecuteLine 11418
        If oLB.getItemData(lngItem) = HS_ROUND Then
            ' allow the user to select a new round code
            Dim l As Long
vbwProfiler.vbwExecuteLine 11419
            l = Val(InputBox("Enter the number of digits after the decimal to round to:", "Edit number of digits", oRule.roundDigits))
vbwProfiler.vbwExecuteLine 11420
            oRule.roundDigits = l
'vbwLine 11421:        ElseIf oLB.getItemData(lngItem) = HS_CONVERSION Then
        ElseIf vbwProfiler.vbwExecuteLine(11421) Or oLB.getItemData(lngItem) = HS_CONVERSION Then
            ' allow the user to select a new conversion

        Else
vbwProfiler.vbwExecuteLine 11422 'B
vbwProfiler.vbwExecuteLine 11423
            Debug.Assert oLB.getItemData(lngItem) >= 0
vbwProfiler.vbwExecuteLine 11424
            Call oRule.setExpressionValue(oLB.getItemData(lngItem), Val(InputBox("Enter new value for this expression.", "Edit expression value")))
        End If
vbwProfiler.vbwExecuteLine 11425 'B
    End If
vbwProfiler.vbwExecuteLine 11426 'B
vbwProfiler.vbwExecuteLine 11427
    CopyMemory oNode, 0&, 4
vbwProfiler.vbwExecuteLine 11428
    renderRule oRule, oLB
vbwProfiler.vbwExecuteLine 11429
    Set oRule = Nothing
vbwProfiler.vbwProcOut 632
vbwProfiler.vbwExecuteLine 11430
End Sub

'Function AddIconToImageList(ByRef s As String) As Long
'    AddIconToImageList = 0
'End Function

Public Function IsNotReservedName(ByRef s As String) As Boolean
vbwProfiler.vbwProcIn 633

    Dim oConvert As cUnitConverter
    Dim lCount As Long
    Dim i As Long
    Dim bRet As Boolean

vbwProfiler.vbwExecuteLine 11431
    Set oConvert = New cUnitConverter

vbwProfiler.vbwExecuteLine 11432
    lCount = oConvert.tableCount
vbwProfiler.vbwExecuteLine 11433
    bRet = True
vbwProfiler.vbwExecuteLine 11434
    For i = 0 To lCount - 1
vbwProfiler.vbwExecuteLine 11435
        If oConvert.tableName(i) = s Then
vbwProfiler.vbwExecuteLine 11436
             bRet = False
        End If
vbwProfiler.vbwExecuteLine 11437 'B
vbwProfiler.vbwExecuteLine 11438
    Next

vbwProfiler.vbwExecuteLine 11439
    Set oConvert = Nothing
vbwProfiler.vbwExecuteLine 11440
    IsNotReservedName = bRet
vbwProfiler.vbwProcOut 633
vbwProfiler.vbwExecuteLine 11441
End Function

