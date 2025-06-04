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
vbwProfiler.vbwProcIn 1
vbwProfiler.vbwExecuteLine 1
    XML = m_xDoc.XML
vbwProfiler.vbwProcOut 1
vbwProfiler.vbwExecuteLine 2
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
vbwProfiler.vbwProcIn 2
vbwProfiler.vbwExecuteLine 3
    On Error GoTo error

    Dim strDesiredInfo As String
vbwProfiler.vbwExecuteLine 4
    If xdiDesiredInfo = xdiVERSION Then
vbwProfiler.vbwExecuteLine 5
        strDesiredInfo = "version"
'vbwLine 6:    ElseIf xdiDesiredInfo = xdiENCODING Then
    ElseIf vbwProfiler.vbwExecuteLine(6) Or xdiDesiredInfo = xdiENCODING Then
vbwProfiler.vbwExecuteLine 7
        strDesiredInfo = "encoding"
'vbwLine 8:    ElseIf xdiDesiredInfo = xdiSTANDALONE Then
    ElseIf vbwProfiler.vbwExecuteLine(8) Or xdiDesiredInfo = xdiSTANDALONE Then
vbwProfiler.vbwExecuteLine 9
        strDesiredInfo = "standalone"
    End If
vbwProfiler.vbwExecuteLine 10 'B

    Dim oXNode As MSXML2.IXMLDOMNode
vbwProfiler.vbwExecuteLine 11
    Set oXNode = m_xDoc.childNodes(0).Attributes.getNamedItem(strDesiredInfo)
vbwProfiler.vbwExecuteLine 12
    m_xDocumentInfo = oXNode.Text
vbwProfiler.vbwExecuteLine 13
    Set oXNode = Nothing

vbwProfiler.vbwProcOut 2
vbwProfiler.vbwExecuteLine 14
    Exit Property
error:
vbwProfiler.vbwExecuteLine 15
    m_xDocumentInfo = Null
vbwProfiler.vbwExecuteLine 16
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 2
vbwProfiler.vbwExecuteLine 17
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
vbwProfiler.vbwProcIn 3
vbwProfiler.vbwExecuteLine 18
    On Error GoTo errorhandler

    Dim oXNode As MSXML2.IXMLDOMNode
vbwProfiler.vbwExecuteLine 19
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 20
    ReadNodeXML = oXNode.XML

vbwProfiler.vbwExecuteLine 21
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 3
vbwProfiler.vbwExecuteLine 22
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 23
    ReadNodeXML = ""
vbwProfiler.vbwExecuteLine 24
    Set oXNode = Nothing

vbwProfiler.vbwProcOut 3
vbwProfiler.vbwExecuteLine 25
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
vbwProfiler.vbwProcIn 4

    ' Create the DOMDocument
vbwProfiler.vbwExecuteLine 26
    Set m_xDoc = New MSXML2.DOMDocument40

vbwProfiler.vbwExecuteLine 27
    If m_xDoc Is Nothing Then
vbwProfiler.vbwExecuteLine 28
        OpenXML = False
vbwProfiler.vbwProcOut 4
vbwProfiler.vbwExecuteLine 29
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 30 'B


vbwProfiler.vbwExecuteLine 31
    m_xDoc.resolveExternals = False
vbwProfiler.vbwExecuteLine 32
    m_xDoc.validateOnParse = True
vbwProfiler.vbwExecuteLine 33
    m_xDoc.async = False

    Dim bolLoadSuccess As Boolean
vbwProfiler.vbwExecuteLine 34
    Select Case oxFileOrString
'vbwLine 35:        Case oxFILE:
        Case IIf(vbwProfiler.vbwExecuteLine(35), VBWPROFILER_EMPTY, _
        oxFILE):
vbwProfiler.vbwExecuteLine 36
            bolLoadSuccess = m_xDoc.Load(strSource) ' .Load = Load from File
'vbwLine 37:        Case oxSTRING:
        Case IIf(vbwProfiler.vbwExecuteLine(37), VBWPROFILER_EMPTY, _
        oxSTRING):
vbwProfiler.vbwExecuteLine 38
            bolLoadSuccess = m_xDoc.loadXML(strSource) ' .LoadXML = Load from String
    End Select
vbwProfiler.vbwExecuteLine 39 'B

vbwProfiler.vbwExecuteLine 40
    Select Case bolLoadSuccess ' Loading successful?
'vbwLine 41:        Case True:
        Case IIf(vbwProfiler.vbwExecuteLine(41), VBWPROFILER_EMPTY, _
        True):
vbwProfiler.vbwExecuteLine 42
            OpenXML = True
'vbwLine 43:        Case False:
        Case IIf(vbwProfiler.vbwExecuteLine(43), VBWPROFILER_EMPTY, _
        False):
vbwProfiler.vbwExecuteLine 44
            On Error Resume Next
            Dim xmlParseError As Object
vbwProfiler.vbwExecuteLine 45
            Set xmlParseError = m_xDoc.parseError


vbwProfiler.vbwExecuteLine 46
                If bolVerboseError = True Then
vbwProfiler.vbwExecuteLine 47
                     Call DisplayParseError
                End If
vbwProfiler.vbwExecuteLine 48 'B
vbwProfiler.vbwExecuteLine 49
                OpenXML = False
vbwProfiler.vbwExecuteLine 50
                Set m_xDoc = Nothing


vbwProfiler.vbwExecuteLine 51
            Set xmlParseError = Nothing
    End Select
vbwProfiler.vbwExecuteLine 52 'B
vbwProfiler.vbwProcOut 4
vbwProfiler.vbwExecuteLine 53
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
vbwProfiler.vbwProcIn 5
vbwProfiler.vbwExecuteLine 54
    OpenFromFile = OpenXML(oxFILE, strFileName, bolVerboseError)
vbwProfiler.vbwProcOut 5
vbwProfiler.vbwExecuteLine 55
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
vbwProfiler.vbwProcIn 6
vbwProfiler.vbwExecuteLine 56
    OpenFromString = OpenXML(oxSTRING, strSource, bolVerboseError)
vbwProfiler.vbwProcOut 6
vbwProfiler.vbwExecuteLine 57
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
vbwProfiler.vbwProcIn 7
vbwProfiler.vbwExecuteLine 58
    On Error GoTo errorhandler

    Dim oXNodeList As MSXML2.IXMLDOMNodeList
vbwProfiler.vbwExecuteLine 59
    Set oXNodeList = m_xDoc.documentElement.selectNodes(strQueryString)
vbwProfiler.vbwExecuteLine 60
    NodeCount = oXNodeList.Length
vbwProfiler.vbwExecuteLine 61
    Set oXNodeList = Nothing

vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 62
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 63
    NodeCount = -1
vbwProfiler.vbwExecuteLine 64
    Set oXNodeList = Nothing
vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 65
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
vbwProfiler.vbwProcIn 8
    Dim oXNode As MSXML2.IXMLDOMNode

vbwProfiler.vbwExecuteLine 66
    Set oXNode = m_xDoc.firstChild.nextSibling
vbwProfiler.vbwExecuteLine 67
    Set GetRootNode = oXNode
vbwProfiler.vbwExecuteLine 68
    Set oXNode = Nothing

vbwProfiler.vbwProcOut 8
vbwProfiler.vbwExecuteLine 69
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
vbwProfiler.vbwProcIn 9
    Dim xTempNode As MSXML2.IXMLDOMNode

vbwProfiler.vbwExecuteLine 70
    Set xTempNode = oXNode.firstChild

vbwProfiler.vbwExecuteLine 71
    If Not xTempNode Is Nothing Then
'vbwLine 72:        Do While xTempNode.nodeName <> strNodeName
        Do While vbwProfiler.vbwExecuteLine(72) Or xTempNode.nodeName <> strNodeName
vbwProfiler.vbwExecuteLine 73
            If Not xTempNode.nextSibling Is Nothing Then
vbwProfiler.vbwExecuteLine 74
                Set xTempNode = xTempNode.nextSibling
            Else
vbwProfiler.vbwExecuteLine 75 'B
vbwProfiler.vbwExecuteLine 76
                Exit Do
            End If
vbwProfiler.vbwExecuteLine 77 'B
vbwProfiler.vbwExecuteLine 78
        Loop
    End If
vbwProfiler.vbwExecuteLine 79 'B

vbwProfiler.vbwExecuteLine 80
    Set GetChildByName = xTempNode
vbwProfiler.vbwExecuteLine 81
    Set xTempNode = Nothing

vbwProfiler.vbwProcOut 9
vbwProfiler.vbwExecuteLine 82
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
vbwProfiler.vbwProcIn 10
    Dim oXParent As MSXML2.IXMLDOMNode
    Dim sXPath As String

vbwProfiler.vbwExecuteLine 83
    If Not oXNode Is Nothing Then
vbwProfiler.vbwExecuteLine 84
        sXPath = "/" & oXNode.nodeName

vbwProfiler.vbwExecuteLine 85
        Set oXParent = oXNode.parentNode
'vbwLine 86:        Do While Not oXParent Is Nothing
        Do While vbwProfiler.vbwExecuteLine(86) Or Not oXParent Is Nothing
            ' we dont want to add root path, first level children is as far as we go
vbwProfiler.vbwExecuteLine 87
            If Not oXParent.parentNode Is Nothing Then
vbwProfiler.vbwExecuteLine 88
                sXPath = "/" & oXParent.nodeName & sXPath
            End If
vbwProfiler.vbwExecuteLine 89 'B
vbwProfiler.vbwExecuteLine 90
            Set oXParent = oXParent.parentNode
vbwProfiler.vbwExecuteLine 91
        Loop
    End If
vbwProfiler.vbwExecuteLine 92 'B

vbwProfiler.vbwExecuteLine 93
    GetXPath = sXPath

vbwProfiler.vbwProcOut 10
vbwProfiler.vbwExecuteLine 94
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
vbwProfiler.vbwProcIn 11
vbwProfiler.vbwExecuteLine 95
    On Error GoTo errorhandler

    Dim oXNode As MSXML2.IXMLDOMNode
vbwProfiler.vbwExecuteLine 96
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 97
    ReadNode = oXNode.Text
vbwProfiler.vbwExecuteLine 98
    Debug.Print "cXML:ReadNode() -- Query = " & strQueryString & " Result = " & oXNode.Text

vbwProfiler.vbwExecuteLine 99
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 11
vbwProfiler.vbwExecuteLine 100
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 101
    ReadNode = Null
vbwProfiler.vbwExecuteLine 102
    Set oXNode = Nothing
vbwProfiler.vbwExecuteLine 103
    Debug.Print "cXML:ReadNode() -- ERROR #" & err.Number & " Description: " & err.Description & " QueryString = " & strQueryString

vbwProfiler.vbwProcOut 11
vbwProfiler.vbwExecuteLine 104
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
vbwProfiler.vbwProcIn 12
vbwProfiler.vbwExecuteLine 105
    On Error GoTo errorhandler

    Dim xmlAttributeNode As MSXML2.IXMLDOMAttribute
vbwProfiler.vbwExecuteLine 106
    Set xmlAttributeNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 107
    ReadAttribute = xmlAttributeNode.getAttribute(strAttributeName)

vbwProfiler.vbwExecuteLine 108
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwProcOut 12
vbwProfiler.vbwExecuteLine 109
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 110
    ReadAttribute = Null
vbwProfiler.vbwExecuteLine 111
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwProcOut 12
vbwProfiler.vbwExecuteLine 112
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
vbwProfiler.vbwProcIn 13
vbwProfiler.vbwExecuteLine 113
    On Error GoTo errorhandler

    Dim oXNode As MSXML2.IXMLDOMNode
vbwProfiler.vbwExecuteLine 114
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 115
    oXNode.parentNode.removeChild oXNode ' This one was tricky ;-)

vbwProfiler.vbwExecuteLine 116
    Set oXNode = Nothing
vbwProfiler.vbwExecuteLine 117
    DeleteNode = True
vbwProfiler.vbwProcOut 13
vbwProfiler.vbwExecuteLine 118
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 119
    DeleteNode = False
vbwProfiler.vbwExecuteLine 120
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 13
vbwProfiler.vbwExecuteLine 121
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
vbwProfiler.vbwProcIn 14
vbwProfiler.vbwExecuteLine 122
    On Error GoTo errorhandler

    Dim oXNode As MSXML2.IXMLDOMNode
vbwProfiler.vbwExecuteLine 123
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 124
    oXNode.Text = vntNewNodeValue

vbwProfiler.vbwExecuteLine 125
    Set oXNode = Nothing
vbwProfiler.vbwExecuteLine 126
    WriteNode = True
vbwProfiler.vbwProcOut 14
vbwProfiler.vbwExecuteLine 127
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 128
    WriteNode = False
vbwProfiler.vbwExecuteLine 129
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 14
vbwProfiler.vbwExecuteLine 130
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
vbwProfiler.vbwProcIn 15
vbwProfiler.vbwExecuteLine 131
    On Error GoTo errorhandler

    Dim xmlAttributeNode As MSXML2.IXMLDOMAttribute
vbwProfiler.vbwExecuteLine 132
    Set xmlAttributeNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 133
    xmlAttributeNode.setAttribute strAttributeName, vntNewAttributeValue

vbwProfiler.vbwExecuteLine 134
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwExecuteLine 135
    WriteAttribute = True
vbwProfiler.vbwProcOut 15
vbwProfiler.vbwExecuteLine 136
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 137
    WriteAttribute = False
vbwProfiler.vbwExecuteLine 138
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwProcOut 15
vbwProfiler.vbwExecuteLine 139
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
vbwProfiler.vbwProcIn 16
vbwProfiler.vbwExecuteLine 140
    On Error GoTo errorhandler

    Dim oXNode As MSXML2.IXMLDOMNode
    Dim xmlCDATASection As Object
vbwProfiler.vbwExecuteLine 141
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

vbwProfiler.vbwExecuteLine 142
    Select Case oXNode.childNodes.Length
'vbwLine 143:        Case 0: ' No Problem, go ahead
        Case IIf(vbwProfiler.vbwExecuteLine(143), VBWPROFILER_EMPTY, _
        0): ' No Problem, go ahead
vbwProfiler.vbwExecuteLine 144
            Set xmlCDATASection = m_xDoc.createCDATASection(vntNewCDataValue)
vbwProfiler.vbwExecuteLine 145
            oXNode.appendChild xmlCDATASection
'vbwLine 146:        Case 1: ' UHOH, better check, if there is already a CData section in form of a CHILD Node present
        Case IIf(vbwProfiler.vbwExecuteLine(146), VBWPROFILER_EMPTY, _
        1): ' UHOH, better check, if there is already a CData section in form of a CHILD Node present
vbwProfiler.vbwExecuteLine 147
            Select Case oXNode.childNodes(0).nodeTypeString
'vbwLine 148:                Case "cdatasection": ' Yes, CDATA section already exists, change Text-Property only!
                Case IIf(vbwProfiler.vbwExecuteLine(148), VBWPROFILER_EMPTY, _
        "cdatasection"): ' Yes, CDATA section already exists, change Text-Property only!
vbwProfiler.vbwExecuteLine 149
                    oXNode.childNodes.Item(0).Text = vntNewCDataValue
                Case Else: ' Probably a Text section, delete Text and replace with CData section
vbwProfiler.vbwExecuteLine 150 'B
vbwProfiler.vbwExecuteLine 151
                    oXNode.Text = ""
vbwProfiler.vbwExecuteLine 152
                    Set xmlCDATASection = m_xDoc.createCDATASection(vntNewCDataValue)
vbwProfiler.vbwExecuteLine 153
                    oXNode.appendChild xmlCDATASection
            End Select
vbwProfiler.vbwExecuteLine 154 'B
        Case Else: ' Raise Error, if more than 1 child node
vbwProfiler.vbwExecuteLine 155 'B
vbwProfiler.vbwExecuteLine 156
            GoTo errorhandler
    End Select
vbwProfiler.vbwExecuteLine 157 'B

vbwProfiler.vbwExecuteLine 158
    Set xmlCDATASection = Nothing
vbwProfiler.vbwExecuteLine 159
    Set oXNode = Nothing
vbwProfiler.vbwExecuteLine 160
    WriteCDataSection = True
vbwProfiler.vbwProcOut 16
vbwProfiler.vbwExecuteLine 161
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 162
    WriteCDataSection = False
vbwProfiler.vbwExecuteLine 163
    Set xmlCDATASection = Nothing
vbwProfiler.vbwExecuteLine 164
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 16
vbwProfiler.vbwExecuteLine 165
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
vbwProfiler.vbwProcIn 17
vbwProfiler.vbwExecuteLine 166
    On Error GoTo errorhandler

    Dim oXNode As MSXML2.IXMLDOMNode
    Dim xmlChildNode As MSXML2.IXMLDOMNode
    Dim xmlAttributeNode As MSXML2.IXMLDOMAttribute

vbwProfiler.vbwExecuteLine 167
    Select Case norRelationship
'vbwLine 168:        Case norCHILD:
        Case IIf(vbwProfiler.vbwExecuteLine(168), VBWPROFILER_EMPTY, _
        norCHILD):
vbwProfiler.vbwExecuteLine 169
            Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)
'vbwLine 170:        Case norSIBLING:
        Case IIf(vbwProfiler.vbwExecuteLine(170), VBWPROFILER_EMPTY, _
        norSIBLING):
vbwProfiler.vbwExecuteLine 171
            Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString).parentNode
    End Select
vbwProfiler.vbwExecuteLine 172 'B

vbwProfiler.vbwExecuteLine 173
    Set xmlChildNode = m_xDoc.createElement(strNewNodeName)

vbwProfiler.vbwExecuteLine 174
    If Not IsMissing(vntNewNodeValue) Then
vbwProfiler.vbwExecuteLine 175
        xmlChildNode.Text = vntNewNodeValue
    End If
vbwProfiler.vbwExecuteLine 176 'B

vbwProfiler.vbwExecuteLine 177
    Set xmlAttributeNode = oXNode.appendChild(xmlChildNode)

vbwProfiler.vbwExecuteLine 178
    If Not IsMissing(strNewAttributeName) And Not IsMissing(vntNewAttributeValue) Then
vbwProfiler.vbwExecuteLine 179
        xmlAttributeNode.setAttribute strNewAttributeName, vntNewAttributeValue
    End If
vbwProfiler.vbwExecuteLine 180 'B

vbwProfiler.vbwExecuteLine 181
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwExecuteLine 182
    Set xmlChildNode = Nothing
vbwProfiler.vbwExecuteLine 183
    Set oXNode = Nothing
vbwProfiler.vbwExecuteLine 184
    InsertNode = True
vbwProfiler.vbwProcOut 17
vbwProfiler.vbwExecuteLine 185
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 186
    InsertNode = False
vbwProfiler.vbwExecuteLine 187
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwExecuteLine 188
    Set xmlChildNode = Nothing
vbwProfiler.vbwExecuteLine 189
    Set oXNode = Nothing
vbwProfiler.vbwProcOut 17
vbwProfiler.vbwExecuteLine 190
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
vbwProfiler.vbwProcIn 18
vbwProfiler.vbwExecuteLine 191
    On Error GoTo errorhandler

    Dim xmlAttributeNode As MSXML2.IXMLDOMAttribute
vbwProfiler.vbwExecuteLine 192
    Set xmlAttributeNode = m_xDoc.documentElement.selectSingleNode(strQueryString).Attributes.removeNamedItem(strAttributeName)

vbwProfiler.vbwExecuteLine 193
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwExecuteLine 194
    DeleteAttribute = True
vbwProfiler.vbwProcOut 18
vbwProfiler.vbwExecuteLine 195
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 196
    DeleteAttribute = False
vbwProfiler.vbwExecuteLine 197
    Set xmlAttributeNode = Nothing
vbwProfiler.vbwProcOut 18
vbwProfiler.vbwExecuteLine 198
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
vbwProfiler.vbwProcIn 19
vbwProfiler.vbwExecuteLine 199
    On Error GoTo errorhandler

    ' Create the DOMDocument
    Dim m_xDocReparse As Object
vbwProfiler.vbwExecuteLine 200
    Set m_xDocReparse = InitializeObject(CONST_DOMDOCUMENT)

vbwProfiler.vbwExecuteLine 201
    If m_xDocReparse.loadXML(m_xDoc.XML) = True Then
vbwProfiler.vbwExecuteLine 202
        Reparse = True
    Else
vbwProfiler.vbwExecuteLine 203 'B
vbwProfiler.vbwExecuteLine 204
        Reparse = False
    End If
vbwProfiler.vbwExecuteLine 205 'B

vbwProfiler.vbwExecuteLine 206
    Set m_xDocReparse = Nothing

vbwProfiler.vbwProcOut 19
vbwProfiler.vbwExecuteLine 207
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 208
    Reparse = False
vbwProfiler.vbwExecuteLine 209
    Set m_xDocReparse = Nothing
vbwProfiler.vbwProcOut 19
vbwProfiler.vbwExecuteLine 210
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
vbwProfiler.vbwProcIn 20
vbwProfiler.vbwExecuteLine 211
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 212
    m_xDoc.Save (strFileName)
vbwProfiler.vbwExecuteLine 213
    Save = True

vbwProfiler.vbwProcOut 20
vbwProfiler.vbwExecuteLine 214
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 215
    Save = False
vbwProfiler.vbwProcOut 20
vbwProfiler.vbwExecuteLine 216
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
vbwProfiler.vbwProcIn 21
    Dim xmlParseError As Object
vbwProfiler.vbwExecuteLine 217
    Set xmlParseError = m_xDoc.parseError

vbwProfiler.vbwExecuteLine 218
    MsgBox "ERROR loading XML" & vbCrLf & vbCrLf & _
        xmlParseError.errorCode & vbCrLf & _
        xmlParseError.reason & "Line " & xmlParseError.Line & _
        ", Position " & xmlParseError.linepos & vbCrLf & vbCrLf & _
        xmlParseError.srcText & vbCrLf _
        , vbExclamation, "XML Parse Error"

vbwProfiler.vbwExecuteLine 219
    Set xmlParseError = Nothing
vbwProfiler.vbwProcOut 21
vbwProfiler.vbwExecuteLine 220
End Sub

Private Sub Class_Terminate()
vbwProfiler.vbwProcIn 22
vbwProfiler.vbwExecuteLine 221
    Set m_xDoc = Nothing
vbwProfiler.vbwProcOut 22
vbwProfiler.vbwExecuteLine 222
End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' METHODS (PRIVATE) - END
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<











