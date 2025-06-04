VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' cXML - a wrapper for MSXML 4.0 to simplify dealing with XML documents
'
'===============================================================================
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' DECLARATIONS - START
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Option Explicit

' The DOMDocument object which holds the XML Document
Dim m_xDoc As MSXML2.DOMDocument40


' Enumeration of possible node relationships the user may select
' with the .InsertNode method
Public Enum NodeRelationshipConstant
    norSIBLING = 0
    norCHILD = 1
End Enum

' Enumeration of possible XML document information the user may select
' with the .m_xDocumentInfo property
Public Enum m_xDocumentInfoConstant
    xdiVERSION = 0
    xdiENCODING = 1
    xdiSTANDALONE = 2
End Enum

' Enumeration of possible Constants used
' with the .OpenXML method
Private Enum OpenXMLConstant
    oxFILE = 0
    oxSTRING = 1
End Enum

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' DECLARATIONS - END
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'-------------------------------------------------------------------------------

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' PROPERTIES (PUBLIC) - START
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Property Get XML() As String
vbwProfiler.vbwProcIn 413
vbwProfiler.vbwExecuteLine 9541
    XML = m_xDoc.XML
vbwProfiler.vbwProcOut 413
vbwProfiler.vbwExecuteLine 9542
End Property


'.------------------------------------------------------------------------------
'.  Function  : Public Property Get m_xDocumentInfo
'.
'.  Parameters: xdiDesiredInfo As m_xDocumentInfoConstant
'.
'.  Returns   : Variant (NULL = Error)
'.
'.  Comments  : Returns information on the XML document's "version", "encoding" or
'.              "standalone" attributes
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Property Get m_xDocumentInfo(xdiDesiredInfo As m_xDocumentInfoConstant) As Variant
vbwProfiler.vbwProcIn 414
vbwProfiler.vbwExecuteLine 9543
    On Error GoTo error

    Dim strDesiredInfo As String
vbwProfiler.vbwExecuteLine 9544
    If xdiDesiredInfo = xdiVERSION Then
vbwProfiler.vbwExecuteLine 9545
        strDesiredInfo = "version"
'vbwLine 9546:    ElseIf xdiDesiredInfo = xdiENCODING Then
    ElseIf vbwProfiler.vbwExecuteLine(9546) Or xdiDesiredInfo = xdiENCODING Then
vbwProfiler.vbwExecuteLine 9547
        strDesiredInfo = "encoding"
'vbwLine 9548:    ElseIf xdiDesiredInfo = xdiSTANDALONE Then
    ElseIf vbwProfiler.vbwExecuteLine(9548) Or xdiDesiredInfo = xdiSTANDALONE Then
vbwProfiler.vbwExecuteLine 9549
        strDesiredInfo = "standalone"
    End If
vbwProfiler.vbwExecuteLine 9550 'B

    Dim oXNode As MSXML2.IXMLDOMNode
vbwProfiler.vbwExecuteLine 9551
    Set oXNode = m_xDoc.childNodes(0).Attributes.getNamedItem(strDesiredInfo)
vbwProfiler.vbwExecuteLine 9552
    m_xDocumentInfo = oXNode.Text
vbwProfiler.vbwExecuteLine 9553
    Set oXNode = Nothing

vbwProfiler.vbwProcOut 414
vbwProfiler.vbwExecuteLine 9554
    Exit Property
error:
vbwProfiler.vbwExecuteLine 9555
    m_xDocumentInfo = Null
vbwProfiler.vbwExecuteLine 9556
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 414
vbwProfiler.vbwExecuteLine 9557
End Property


'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' PROPERTIES (PUBLIC) - END
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'-------------------------------------------------------------------------------

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' METHODS (PUBLIC) - START
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'.------------------------------------------------------------------------------
'.  Function  : Public Function ReadNodeXML
'.
'.  Parameters: ByVal strQueryString As String
'.
'.  Returns   : String (Error = "")
'.
'.  Comments  : Return the supplied node's XML value
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function ReadNodeXML(ByVal strQueryString As String) As String
vbwProfiler.vbwProcIn 415
vbwProfiler.vbwExecuteLine 9558
    On Error GoTo errorhandler

    Dim oXNode As MSXML2.IXMLDOMNode
vbwProfiler.vbwExecuteLine 9559
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 9560
    ReadNodeXML = oXNode.XML

vbwProfiler.vbwExecuteLine 9561
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 415
vbwProfiler.vbwExecuteLine 9562
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 9563
    ReadNodeXML = ""
vbwProfiler.vbwExecuteLine 9564
    Set oXNode = Nothing

vbwProfiler.vbwProcOut 415
vbwProfiler.vbwExecuteLine 9565
End Function

'.------------------------------------------------------------------------------
'.  Function  : Private Function OpenXML
'.
'.  Parameters: ByVal oxFileOrString As OpenXMLConstant
'.              ByVal strSource As String
'.              Optional ByVal bolVerboseError As Boolean = False
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Tries to open and parse strSource,
'.              either from file or from string.
'.              Option: Display MS XML Parser Error in MsgBox?
'.
'.

'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Private Function OpenXML(ByVal oxFileOrString As OpenXMLConstant, ByVal strSource As String, Optional ByVal bolVerboseError As Boolean = False) As Boolean
vbwProfiler.vbwProcIn 416

    ' Create the DOMDocument
vbwProfiler.vbwExecuteLine 9566
    Set m_xDoc = New MSXML2.DOMDocument40

vbwProfiler.vbwExecuteLine 9567
    If m_xDoc Is Nothing Then
vbwProfiler.vbwExecuteLine 9568
        OpenXML = False
vbwProfiler.vbwProcOut 416
vbwProfiler.vbwExecuteLine 9569
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 9570 'B


vbwProfiler.vbwExecuteLine 9571
    m_xDoc.resolveExternals = False
vbwProfiler.vbwExecuteLine 9572
    m_xDoc.validateOnParse = True
vbwProfiler.vbwExecuteLine 9573
    m_xDoc.async = False

    Dim bolLoadSuccess As Boolean
vbwProfiler.vbwExecuteLine 9574
    Select Case oxFileOrString
'vbwLine 9575:        Case oxFILE:
        Case IIf(vbwProfiler.vbwExecuteLine(9575), VBWPROFILER_EMPTY, _
        oxFILE):
vbwProfiler.vbwExecuteLine 9576
            bolLoadSuccess = m_xDoc.Load(strSource) ' .Load = Load from File
'vbwLine 9577:        Case oxSTRING:
        Case IIf(vbwProfiler.vbwExecuteLine(9577), VBWPROFILER_EMPTY, _
        oxSTRING):
vbwProfiler.vbwExecuteLine 9578
            bolLoadSuccess = m_xDoc.loadXML(strSource) ' .LoadXML = Load from String
    End Select
vbwProfiler.vbwExecuteLine 9579 'B

vbwProfiler.vbwExecuteLine 9580
    Select Case bolLoadSuccess ' Loading successful?
'vbwLine 9581:        Case True:
        Case IIf(vbwProfiler.vbwExecuteLine(9581), VBWPROFILER_EMPTY, _
        True):
vbwProfiler.vbwExecuteLine 9582
            OpenXML = True
'vbwLine 9583:        Case False:
        Case IIf(vbwProfiler.vbwExecuteLine(9583), VBWPROFILER_EMPTY, _
        False):
vbwProfiler.vbwExecuteLine 9584
            On Error Resume Next
            Dim xmlParseError As Object
vbwProfiler.vbwExecuteLine 9585
            Set xmlParseError = m_xDoc.parseError


vbwProfiler.vbwExecuteLine 9586
                If bolVerboseError = True Then
vbwProfiler.vbwExecuteLine 9587
                     Call DisplayParseError
                End If
vbwProfiler.vbwExecuteLine 9588 'B
vbwProfiler.vbwExecuteLine 9589
                OpenXML = False
vbwProfiler.vbwExecuteLine 9590
                Set m_xDoc = Nothing


vbwProfiler.vbwExecuteLine 9591
            Set xmlParseError = Nothing
    End Select
vbwProfiler.vbwExecuteLine 9592 'B
vbwProfiler.vbwProcOut 416
vbwProfiler.vbwExecuteLine 9593
End Function

'.------------------------------------------------------------------------------
'.  Function  : Public Function OpenFromFile
'.
'.  Parameters: ByVal strFileName As String
'.              Optional ByVal bolVerboseError As Boolean = False
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Open XML document from file
'.              Option: Display MS XML Parser Error in MsgBox
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function OpenFromFile(ByVal strFileName As String, Optional ByVal bolVerboseError As Boolean = False) As Boolean
vbwProfiler.vbwProcIn 417
vbwProfiler.vbwExecuteLine 9594
    OpenFromFile = OpenXML(oxFILE, strFileName, bolVerboseError)
vbwProfiler.vbwProcOut 417
vbwProfiler.vbwExecuteLine 9595
End Function

'.------------------------------------------------------------------------------
'.  Function  : Public Function OpenFromString
'.
'.  Parameters: ByVal strSource As String
'.              Optional ByVal bolVerboseError As Boolean = False
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Open XML document from file
'.              Option: Display MS XML Parser Error in MsgBox
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   :
'.------------------------------------------------------------------------------

Public Function OpenFromString(ByVal strSource As String, Optional ByVal bolVerboseError As Boolean = False) As Boolean
vbwProfiler.vbwProcIn 418
vbwProfiler.vbwExecuteLine 9596
    OpenFromString = OpenXML(oxSTRING, strSource, bolVerboseError)
vbwProfiler.vbwProcOut 418
vbwProfiler.vbwExecuteLine 9597
End Function


'.------------------------------------------------------------------------------
'.  Function  : Public Function NodeCount
'.
'.  Parameters: ByVal strQueryString As String
'.
'.  Returns   : Long (Error = -1)
'.
'.  Comments  : Return the number of nodes found
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function NodeCount(ByVal strQueryString As String) As Long
vbwProfiler.vbwProcIn 419
vbwProfiler.vbwExecuteLine 9598
    On Error GoTo errorhandler

    Dim oXNodeList As MSXML2.IXMLDOMNodeList
vbwProfiler.vbwExecuteLine 9599
    Set oXNodeList = m_xDoc.documentElement.selectNodes(strQueryString)
vbwProfiler.vbwExecuteLine 9600
    NodeCount = oXNodeList.Length
vbwProfiler.vbwExecuteLine 9601
    Set oXNodeList = Nothing

vbwProfiler.vbwProcOut 419
vbwProfiler.vbwExecuteLine 9602
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 9603
    NodeCount = -1
vbwProfiler.vbwExecuteLine 9604
    Set oXNodeList = Nothing
vbwProfiler.vbwProcOut 419
vbwProfiler.vbwExecuteLine 9605
End Function

'.------------------------------------------------------------------------------
'.  Function  : Public Function GetRootNode
'.
'.  Parameters:
'.  Returns   : Object
'.
'.  Comments  : Returns first Element with querystring
'.
'.  Author    : Mike Joseph - 15.09.2002
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function GetRootNode() As Object
vbwProfiler.vbwProcIn 420
    Dim oXNode As MSXML2.IXMLDOMNode

vbwProfiler.vbwExecuteLine 9606
    Set oXNode = m_xDoc.firstChild.nextSibling
vbwProfiler.vbwExecuteLine 9607
    Set GetRootNode = oXNode
vbwProfiler.vbwExecuteLine 9608
    Set oXNode = Nothing

vbwProfiler.vbwProcOut 420
vbwProfiler.vbwExecuteLine 9609
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function GetChildByName
'.
'.  Parameters: ByVal strNodeName As String
'.              ByRef oXNode as Object
'.
'.  Returns   : Long (Error = -1)
'.
'.  Comments  : Returns first Object with querystring
'.
'.  Author    : Mike Joseph - 15.09.2002
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function GetChildByName(strNodeName As String, oXNode As MSXML2.IXMLDOMNode) As Object
vbwProfiler.vbwProcIn 421
    Dim xTempNode As MSXML2.IXMLDOMNode

vbwProfiler.vbwExecuteLine 9610
    Set xTempNode = oXNode.firstChild

vbwProfiler.vbwExecuteLine 9611
    If Not xTempNode Is Nothing Then
'vbwLine 9612:        Do While xTempNode.nodeName <> strNodeName
        Do While vbwProfiler.vbwExecuteLine(9612) Or xTempNode.nodeName <> strNodeName
vbwProfiler.vbwExecuteLine 9613
            If Not xTempNode.nextSibling Is Nothing Then
vbwProfiler.vbwExecuteLine 9614
                Set xTempNode = xTempNode.nextSibling
            Else
vbwProfiler.vbwExecuteLine 9615 'B
vbwProfiler.vbwExecuteLine 9616
                Exit Do
            End If
vbwProfiler.vbwExecuteLine 9617 'B
vbwProfiler.vbwExecuteLine 9618
        Loop
    End If
vbwProfiler.vbwExecuteLine 9619 'B

vbwProfiler.vbwExecuteLine 9620
    Set GetChildByName = xTempNode
vbwProfiler.vbwExecuteLine 9621
    Set xTempNode = Nothing

vbwProfiler.vbwProcOut 421
vbwProfiler.vbwExecuteLine 9622
End Function

'.------------------------------------------------------------------------------
'.  Function  : Public Function GetXPath(o as object) as string
'.
'.  Parameters: ooXNode as object
'.
'.  Returns   : string)
'.
'.  Comments  : Returns the node's XPath
'.
'.  Author    : Michael Joseph - 01.Oct.2002
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function GetXPath(oXNode As MSXML2.IXMLDOMNode) As String
vbwProfiler.vbwProcIn 422
    Dim oXParent As MSXML2.IXMLDOMNode
    Dim sXPath As String

vbwProfiler.vbwExecuteLine 9623
    If Not oXNode Is Nothing Then
vbwProfiler.vbwExecuteLine 9624
        sXPath = "/" & oXNode.nodeName

vbwProfiler.vbwExecuteLine 9625
        Set oXParent = oXNode.parentNode
'vbwLine 9626:        Do While Not oXParent Is Nothing
        Do While vbwProfiler.vbwExecuteLine(9626) Or Not oXParent Is Nothing
            ' we dont want to add root path, first level children is as far as we go
vbwProfiler.vbwExecuteLine 9627
            If Not oXParent.parentNode Is Nothing Then
vbwProfiler.vbwExecuteLine 9628
                sXPath = "/" & oXParent.nodeName & sXPath
            End If
vbwProfiler.vbwExecuteLine 9629 'B
vbwProfiler.vbwExecuteLine 9630
            Set oXParent = oXParent.parentNode
vbwProfiler.vbwExecuteLine 9631
        Loop
    End If
vbwProfiler.vbwExecuteLine 9632 'B

vbwProfiler.vbwExecuteLine 9633
    GetXPath = sXPath

vbwProfiler.vbwProcOut 422
vbwProfiler.vbwExecuteLine 9634
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function ReadNode
'.
'.  Parameters: ByVal strQueryString As String
'.
'.  Returns   : Variant (Error = NULL)
'.
'.  Comments  : Return the node text found
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function ReadNode(ByVal strQueryString As String) As Variant
vbwProfiler.vbwProcIn 423
vbwProfiler.vbwExecuteLine 9635
    On Error GoTo errorhandler

    Dim oXNode As MSXML2.IXMLDOMNode
vbwProfiler.vbwExecuteLine 9636
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 9637
    ReadNode = oXNode.Text
vbwProfiler.vbwExecuteLine 9638
    Debug.Print "cXML:ReadNode() -- Query = " & strQueryString & " Result = " & oXNode.Text

vbwProfiler.vbwExecuteLine 9639
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 423
vbwProfiler.vbwExecuteLine 9640
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 9641
    ReadNode = Null
vbwProfiler.vbwExecuteLine 9642
    Set oXNode = Nothing
vbwProfiler.vbwExecuteLine 9643
    Debug.Print "cXML:ReadNode() -- ERROR #" & err.Number & " Description: " & err.Description & " QueryString = " & strQueryString

vbwProfiler.vbwProcOut 423
vbwProfiler.vbwExecuteLine 9644
End Function

'.------------------------------------------------------------------------------
'.  Function  : Public Function NextSibling
'.
'.  Parameters: ByVal strQueryString As String
'.              ByVal strAttributeName As String
'.
'.  Returns   : Variant (Error = NULL)
'.
'.  Comments  : Itterator to get next sibling node
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
'Public Function NextSibling() as MSXML2.IXMLDOMNode
    'todo: need to modify class to track a module level "m_oCurrentNode"
    '      also needs to reset m_oCurrentNode to Nothing when loading a new XML Document since it invalidates the current node
'End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function ReadAttribute
'.
'.  Parameters: ByVal strQueryString As String
'.              ByVal strAttributeName As String
'.
'.  Returns   : Variant (Error = NULL)
'.
'.  Comments  : Return supplied attribute's value from supplied node
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function ReadAttribute(ByVal strQueryString As String, ByVal strAttributeName As String) As Variant
vbwProfiler.vbwProcIn 424
vbwProfiler.vbwExecuteLine 9645
    On Error GoTo errorhandler

    Dim xmlAttributeNode As MSXML2.IXMLDOMAttribute
vbwProfiler.vbwExecuteLine 9646
    Set xmlAttributeNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 9647
    ReadAttribute = xmlAttributeNode.getAttribute(strAttributeName)

vbwProfiler.vbwExecuteLine 9648
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwProcOut 424
vbwProfiler.vbwExecuteLine 9649
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 9650
    ReadAttribute = Null
vbwProfiler.vbwExecuteLine 9651
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwProcOut 424
vbwProfiler.vbwExecuteLine 9652
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function DeleteNode
'.
'.  Parameters: ByVal strQueryString As String
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Delete node
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function DeleteNode(ByVal strQueryString As String) As Boolean
vbwProfiler.vbwProcIn 425
vbwProfiler.vbwExecuteLine 9653
    On Error GoTo errorhandler

    Dim oXNode As MSXML2.IXMLDOMNode
vbwProfiler.vbwExecuteLine 9654
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 9655
    oXNode.parentNode.removeChild oXNode ' This one was tricky ;-)

vbwProfiler.vbwExecuteLine 9656
    Set oXNode = Nothing
vbwProfiler.vbwExecuteLine 9657
    DeleteNode = True
vbwProfiler.vbwProcOut 425
vbwProfiler.vbwExecuteLine 9658
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 9659
    DeleteNode = False
vbwProfiler.vbwExecuteLine 9660
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 425
vbwProfiler.vbwExecuteLine 9661
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function WriteNode
'.
'.  Parameters: ByVal strQueryString As String
'.              ByVal vntNewNodeValue As Variant
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Write a supplied node's text value (as variant)
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function WriteNode(ByVal strQueryString As String, ByVal vntNewNodeValue As Variant) As Boolean
vbwProfiler.vbwProcIn 426
vbwProfiler.vbwExecuteLine 9662
    On Error GoTo errorhandler

    Dim oXNode As MSXML2.IXMLDOMNode
vbwProfiler.vbwExecuteLine 9663
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 9664
    oXNode.Text = vntNewNodeValue

vbwProfiler.vbwExecuteLine 9665
    Set oXNode = Nothing
vbwProfiler.vbwExecuteLine 9666
    WriteNode = True
vbwProfiler.vbwProcOut 426
vbwProfiler.vbwExecuteLine 9667
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 9668
    WriteNode = False
vbwProfiler.vbwExecuteLine 9669
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 426
vbwProfiler.vbwExecuteLine 9670
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function WriteAttribute
'.
'.  Parameters: ByVal strQueryString As String
'.              ByVal strAttributeName As String
'.              ByVal vntNewAttributeValue As Variant
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Write supplied node's supplied attribute's value
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function WriteAttribute(ByVal strQueryString As String, ByVal strAttributeName As String, ByVal vntNewAttributeValue As Variant) As Boolean
vbwProfiler.vbwProcIn 427
vbwProfiler.vbwExecuteLine 9671
    On Error GoTo errorhandler

    Dim xmlAttributeNode As MSXML2.IXMLDOMAttribute
vbwProfiler.vbwExecuteLine 9672
    Set xmlAttributeNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 9673
    xmlAttributeNode.setAttribute strAttributeName, vntNewAttributeValue

vbwProfiler.vbwExecuteLine 9674
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwExecuteLine 9675
    WriteAttribute = True
vbwProfiler.vbwProcOut 427
vbwProfiler.vbwExecuteLine 9676
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 9677
    WriteAttribute = False
vbwProfiler.vbwExecuteLine 9678
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwProcOut 427
vbwProfiler.vbwExecuteLine 9679
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function WriteCDataSection
'.
'.  Parameters: ByVal strQueryString As String
'.              ByVal vntNewCDataValue As Variant
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Write supplied node's CDATA section
'.              WILL TRY TO OVERWRITE AN EXISTING CDATA SECTION,
'.              INSTEAD OF APPENDING TO IT.
'.              Personally, I can't think of any good reason to have
'.              more than one CDATA section, but I may just as well be wrong...
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function WriteCDataSection(ByVal strQueryString As String, ByVal vntNewCDataValue As Variant) As Boolean
vbwProfiler.vbwProcIn 428
vbwProfiler.vbwExecuteLine 9680
    On Error GoTo errorhandler

    Dim oXNode As MSXML2.IXMLDOMNode
    Dim xmlCDATASection As Object
vbwProfiler.vbwExecuteLine 9681
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 9682
    Select Case oXNode.childNodes.Length
'vbwLine 9683:        Case 0: ' No Problem, go ahead
        Case IIf(vbwProfiler.vbwExecuteLine(9683), VBWPROFILER_EMPTY, _
        0): ' No Problem, go ahead
vbwProfiler.vbwExecuteLine 9684
            Set xmlCDATASection = m_xDoc.createCDATASection(vntNewCDataValue)
vbwProfiler.vbwExecuteLine 9685
            oXNode.appendChild xmlCDATASection
'vbwLine 9686:        Case 1: ' UHOH, better check, if there is already a CData section in form of a CHILD Node present
        Case IIf(vbwProfiler.vbwExecuteLine(9686), VBWPROFILER_EMPTY, _
        1): ' UHOH, better check, if there is already a CData section in form of a CHILD Node present
vbwProfiler.vbwExecuteLine 9687
            Select Case oXNode.childNodes(0).nodeTypeString
'vbwLine 9688:                Case "cdatasection": ' Yes, CDATA section already exists, change Text-Property only!
                Case IIf(vbwProfiler.vbwExecuteLine(9688), VBWPROFILER_EMPTY, _
        "cdatasection"): ' Yes, CDATA section already exists, change Text-Property only!
vbwProfiler.vbwExecuteLine 9689
                    oXNode.childNodes.Item(0).Text = vntNewCDataValue
                Case Else: ' Probably a Text section, delete Text and replace with CData section
vbwProfiler.vbwExecuteLine 9690 'B
vbwProfiler.vbwExecuteLine 9691
                    oXNode.Text = ""
vbwProfiler.vbwExecuteLine 9692
                    Set xmlCDATASection = m_xDoc.createCDATASection(vntNewCDataValue)
vbwProfiler.vbwExecuteLine 9693
                    oXNode.appendChild xmlCDATASection
            End Select
vbwProfiler.vbwExecuteLine 9694 'B
        Case Else: ' Raise Error, if more than 1 child node
vbwProfiler.vbwExecuteLine 9695 'B
vbwProfiler.vbwExecuteLine 9696
            GoTo errorhandler
    End Select
vbwProfiler.vbwExecuteLine 9697 'B

vbwProfiler.vbwExecuteLine 9698
    Set xmlCDATASection = Nothing
vbwProfiler.vbwExecuteLine 9699
    Set oXNode = Nothing
vbwProfiler.vbwExecuteLine 9700
    WriteCDataSection = True
vbwProfiler.vbwProcOut 428
vbwProfiler.vbwExecuteLine 9701
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 9702
    WriteCDataSection = False
vbwProfiler.vbwExecuteLine 9703
    Set xmlCDATASection = Nothing
vbwProfiler.vbwExecuteLine 9704
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 428
vbwProfiler.vbwExecuteLine 9705
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function InsertNode
'.
'.  Parameters: ByVal strQueryString As String
'.              ByVal strNewNodeName As String
'.              Optional ByVal vntNewNodeValue As Variant
'.              Optional ByVal strNewAttributeName As String
'.              Optional ByVal vntNewAttributeValue As Variant
'.              Optional ByVal norRelationship As NodeRelationshipConstant = norCHILD
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Lets you insert a new Child OR SIBLING node, along
'.              with ONE attribute / attribute value pair
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function InsertNode(ByVal strQueryString As String, ByVal strNewNodeName As String, Optional ByVal vntNewNodeValue As Variant, Optional ByVal strNewAttributeName As String, Optional ByVal vntNewAttributeValue As Variant, Optional ByVal norRelationship As NodeRelationshipConstant = norCHILD) As Boolean
vbwProfiler.vbwProcIn 429
vbwProfiler.vbwExecuteLine 9706
    On Error GoTo errorhandler

    Dim oXNode As MSXML2.IXMLDOMNode
    Dim xmlChildNode As MSXML2.IXMLDOMNode
    Dim xmlAttributeNode As MSXML2.IXMLDOMAttribute

vbwProfiler.vbwExecuteLine 9707
    Select Case norRelationship
'vbwLine 9708:        Case norCHILD:
        Case IIf(vbwProfiler.vbwExecuteLine(9708), VBWPROFILER_EMPTY, _
        norCHILD):
vbwProfiler.vbwExecuteLine 9709
            Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)
'vbwLine 9710:        Case norSIBLING:
        Case IIf(vbwProfiler.vbwExecuteLine(9710), VBWPROFILER_EMPTY, _
        norSIBLING):
vbwProfiler.vbwExecuteLine 9711
            Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString).parentNode
    End Select
vbwProfiler.vbwExecuteLine 9712 'B

vbwProfiler.vbwExecuteLine 9713
    Set xmlChildNode = m_xDoc.createElement(strNewNodeName)

vbwProfiler.vbwExecuteLine 9714
    If Not IsMissing(vntNewNodeValue) Then
vbwProfiler.vbwExecuteLine 9715
        xmlChildNode.Text = vntNewNodeValue
    End If
vbwProfiler.vbwExecuteLine 9716 'B

vbwProfiler.vbwExecuteLine 9717
    Set xmlAttributeNode = oXNode.appendChild(xmlChildNode)

vbwProfiler.vbwExecuteLine 9718
    If Not IsMissing(strNewAttributeName) And Not IsMissing(vntNewAttributeValue) Then
vbwProfiler.vbwExecuteLine 9719
        xmlAttributeNode.setAttribute strNewAttributeName, vntNewAttributeValue
    End If
vbwProfiler.vbwExecuteLine 9720 'B

vbwProfiler.vbwExecuteLine 9721
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwExecuteLine 9722
    Set xmlChildNode = Nothing
vbwProfiler.vbwExecuteLine 9723
    Set oXNode = Nothing
vbwProfiler.vbwExecuteLine 9724
    InsertNode = True
vbwProfiler.vbwProcOut 429
vbwProfiler.vbwExecuteLine 9725
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 9726
    InsertNode = False
vbwProfiler.vbwExecuteLine 9727
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwExecuteLine 9728
    Set xmlChildNode = Nothing
vbwProfiler.vbwExecuteLine 9729
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 429
vbwProfiler.vbwExecuteLine 9730
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function DeleteAttribute
'.
'.  Parameters: ByVal strQueryString As String
'.              ByVal strAttributeName As String
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Delete attribute / attribute value pair form supplied node
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function DeleteAttribute(ByVal strQueryString As String, ByVal strAttributeName As String) As Boolean
vbwProfiler.vbwProcIn 430
vbwProfiler.vbwExecuteLine 9731
    On Error GoTo errorhandler

    Dim xmlAttributeNode As MSXML2.IXMLDOMAttribute
vbwProfiler.vbwExecuteLine 9732
    Set xmlAttributeNode = m_xDoc.documentElement.selectSingleNode(strQueryString).Attributes.removeNamedItem(strAttributeName)

vbwProfiler.vbwExecuteLine 9733
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwExecuteLine 9734
    DeleteAttribute = True
vbwProfiler.vbwProcOut 430
vbwProfiler.vbwExecuteLine 9735
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 9736
    DeleteAttribute = False
vbwProfiler.vbwExecuteLine 9737
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwProcOut 430
vbwProfiler.vbwExecuteLine 9738
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function Reparse
'.
'.  Parameters: -
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Recheck if current XML document is still valid XML
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function Reparse() As Boolean
vbwProfiler.vbwProcIn 431
vbwProfiler.vbwExecuteLine 9739
    On Error GoTo errorhandler

    ' Create the DOMDocument
    Dim m_xDocReparse As Object
vbwProfiler.vbwExecuteLine 9740
    Set m_xDocReparse = InitializeObject(CONST_DOMDOCUMENT)

vbwProfiler.vbwExecuteLine 9741
    If m_xDocReparse.loadXML(m_xDoc.XML) = True Then
vbwProfiler.vbwExecuteLine 9742
        Reparse = True
    Else
vbwProfiler.vbwExecuteLine 9743 'B
vbwProfiler.vbwExecuteLine 9744
        Reparse = False
    End If
vbwProfiler.vbwExecuteLine 9745 'B

vbwProfiler.vbwExecuteLine 9746
    Set m_xDocReparse = Nothing

vbwProfiler.vbwProcOut 431
vbwProfiler.vbwExecuteLine 9747
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 9748
    Reparse = False
vbwProfiler.vbwExecuteLine 9749
    Set m_xDocReparse = Nothing
vbwProfiler.vbwProcOut 431
vbwProfiler.vbwExecuteLine 9750
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function Save
'.
'.  Parameters: ByVal strFileName As String
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Save current XML document to file
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function Save(ByVal strFileName As String) As Boolean
vbwProfiler.vbwProcIn 432
vbwProfiler.vbwExecuteLine 9751
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 9752
    m_xDoc.Save (strFileName)
vbwProfiler.vbwExecuteLine 9753
    Save = True

vbwProfiler.vbwProcOut 432
vbwProfiler.vbwExecuteLine 9754
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 9755
    Save = False
vbwProfiler.vbwProcOut 432
vbwProfiler.vbwExecuteLine 9756
End Function

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' METHODS (PUBLIC) - END
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'-------------------------------------------------------------------------------

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' METHODS (PRIVATE) - START
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'.------------------------------------------------------------------------------
'.  Function  : Private Sub DisplayParseError
'.
'.  Parameters: -
'.
'.  Returns   : -
'.
'.  Comments  : Display the latest ParseError from MS XML Parser
'.
'.  Author    : Roman Kehr - 26.10.2001
'.  Changed   : -
'.------------------------------------------------------------------------------
Private Sub DisplayParseError()
vbwProfiler.vbwProcIn 433
    Dim xmlParseError As Object
vbwProfiler.vbwExecuteLine 9757
    Set xmlParseError = m_xDoc.parseError

vbwProfiler.vbwExecuteLine 9758
    MsgBox "ERROR loading XML" & vbCrLf & vbCrLf & _
        xmlParseError.errorCode & vbCrLf & _
        xmlParseError.reason & "Line " & xmlParseError.Line & _
        ", Position " & xmlParseError.linepos & vbCrLf & vbCrLf & _
        xmlParseError.srcText & vbCrLf _
        , vbExclamation, "XML Parse Error"

vbwProfiler.vbwExecuteLine 9759
    Set xmlParseError = Nothing
vbwProfiler.vbwProcOut 433
vbwProfiler.vbwExecuteLine 9760
End Sub

Private Sub Class_Terminate()
vbwProfiler.vbwProcIn 434
vbwProfiler.vbwExecuteLine 9761
    Set m_xDoc = Nothing
vbwProfiler.vbwProcOut 434
vbwProfiler.vbwExecuteLine 9762
End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' METHODS (PRIVATE) - END
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<











