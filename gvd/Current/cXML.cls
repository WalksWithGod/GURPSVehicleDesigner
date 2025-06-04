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
    XML = m_xDoc.XML
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
    On Error GoTo error
    
    Dim strDesiredInfo As String
    If xdiDesiredInfo = xdiVERSION Then
        strDesiredInfo = "version"
    ElseIf xdiDesiredInfo = xdiENCODING Then
        strDesiredInfo = "encoding"
    ElseIf xdiDesiredInfo = xdiSTANDALONE Then
        strDesiredInfo = "standalone"
    End If
    
    Dim oXNode As MSXML2.IXMLDOMNode
    Set oXNode = m_xDoc.childNodes(0).Attributes.getNamedItem(strDesiredInfo)
    m_xDocumentInfo = oXNode.Text
    Set oXNode = Nothing
    
    Exit Property
error:
    m_xDocumentInfo = Null
    Set oXNode = Nothing
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
    On Error GoTo errorhandler
    
    Dim oXNode As MSXML2.IXMLDOMNode
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)
  
    ReadNodeXML = oXNode.XML
    
    Set oXNode = Nothing
    Exit Function
errorhandler:
    ReadNodeXML = ""
    Set oXNode = Nothing
   
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
    
    ' Create the DOMDocument
    Set m_xDoc = New MSXML2.DOMDocument40

    If m_xDoc Is Nothing Then
        OpenXML = False
        Exit Function
    End If
    
    
    m_xDoc.resolveExternals = False
    m_xDoc.validateOnParse = True
    m_xDoc.async = False
    
    Dim bolLoadSuccess As Boolean
    Select Case oxFileOrString
        Case oxFILE:
            bolLoadSuccess = m_xDoc.Load(strSource) ' .Load = Load from File
        Case oxSTRING:
            bolLoadSuccess = m_xDoc.loadXML(strSource) ' .LoadXML = Load from String
    End Select
        
    Select Case bolLoadSuccess ' Loading successful?
        Case True:
            OpenXML = True
        Case False:
            On Error Resume Next
            Dim xmlParseError As Object
            Set xmlParseError = m_xDoc.parseError
            
         
                If bolVerboseError = True Then Call DisplayParseError
                OpenXML = False
                Set m_xDoc = Nothing
            
            
            Set xmlParseError = Nothing
    End Select
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
    OpenFromFile = OpenXML(oxFILE, strFileName, bolVerboseError)
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
    OpenFromString = OpenXML(oxSTRING, strSource, bolVerboseError)
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
    On Error GoTo errorhandler

    Dim oXNodeList As MSXML2.IXMLDOMNodeList
    Set oXNodeList = m_xDoc.documentElement.selectNodes(strQueryString)
    NodeCount = oXNodeList.Length
    Set oXNodeList = Nothing

    Exit Function
errorhandler:
    NodeCount = -1
    Set oXNodeList = Nothing
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
    Dim oXNode As MSXML2.IXMLDOMNode
    
    Set oXNode = m_xDoc.firstChild.nextSibling
    Set GetRootNode = oXNode
    Set oXNode = Nothing
    
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
    Dim xTempNode As MSXML2.IXMLDOMNode
    
    Set xTempNode = oXNode.firstChild
        
    If Not xTempNode Is Nothing Then
        Do While xTempNode.nodeName <> strNodeName
            If Not xTempNode.nextSibling Is Nothing Then
                Set xTempNode = xTempNode.nextSibling
            Else
                Exit Do
            End If
        Loop
    End If
    
    Set GetChildByName = xTempNode
    Set xTempNode = Nothing
    
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
    Dim oXParent As MSXML2.IXMLDOMNode
    Dim sXPath As String
    
    If Not oXNode Is Nothing Then
        sXPath = "/" & oXNode.nodeName
        
        Set oXParent = oXNode.parentNode
        Do While Not oXParent Is Nothing
            ' we dont want to add root path, first level children is as far as we go
            If Not oXParent.parentNode Is Nothing Then
                sXPath = "/" & oXParent.nodeName & sXPath
            End If
            Set oXParent = oXParent.parentNode
        Loop
    End If
    
    GetXPath = sXPath

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
    On Error GoTo errorhandler
    
    Dim oXNode As MSXML2.IXMLDOMNode
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)
  
    ReadNode = oXNode.Text
    Debug.Print "cXML:ReadNode() -- Query = " & strQueryString & " Result = " & oXNode.Text
    
    Set oXNode = Nothing
    Exit Function
errorhandler:
    ReadNode = Null
    Set oXNode = Nothing
    Debug.Print "cXML:ReadNode() -- ERROR #" & err.Number & " Description: " & err.Description & " QueryString = " & strQueryString
    
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
    On Error GoTo errorhandler
    
    Dim xmlAttributeNode As MSXML2.IXMLDOMAttribute
    Set xmlAttributeNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

    ReadAttribute = xmlAttributeNode.getAttribute(strAttributeName)

    Set xmlAttributeNode = Nothing
    Exit Function
errorhandler:
    ReadAttribute = Null
    Set xmlAttributeNode = Nothing
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
    On Error GoTo errorhandler
    
    Dim oXNode As MSXML2.IXMLDOMNode
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)
    
    oXNode.parentNode.removeChild oXNode ' This one was tricky ;-)
    
    Set oXNode = Nothing
    DeleteNode = True
    Exit Function
errorhandler:
    DeleteNode = False
    Set oXNode = Nothing
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
    On Error GoTo errorhandler
   
    Dim oXNode As MSXML2.IXMLDOMNode
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)
    
    oXNode.Text = vntNewNodeValue
    
    Set oXNode = Nothing
    WriteNode = True
    Exit Function
errorhandler:
    WriteNode = False
    Set oXNode = Nothing
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
    On Error GoTo errorhandler

    Dim xmlAttributeNode As MSXML2.IXMLDOMAttribute
    Set xmlAttributeNode = m_xDoc.documentElement.selectSingleNode(strQueryString)
    
    xmlAttributeNode.setAttribute strAttributeName, vntNewAttributeValue
    
    Set xmlAttributeNode = Nothing
    WriteAttribute = True
    Exit Function
errorhandler:
    WriteAttribute = False
    Set xmlAttributeNode = Nothing
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
    On Error GoTo errorhandler
    
    Dim oXNode As MSXML2.IXMLDOMNode
    Dim xmlCDATASection As Object
    Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)

    Select Case oXNode.childNodes.Length
        Case 0: ' No Problem, go ahead
            Set xmlCDATASection = m_xDoc.createCDATASection(vntNewCDataValue)
            oXNode.appendChild xmlCDATASection
        Case 1: ' UHOH, better check, if there is already a CData section in form of a CHILD Node present
            Select Case oXNode.childNodes(0).nodeTypeString
                Case "cdatasection": ' Yes, CDATA section already exists, change Text-Property only!
                    oXNode.childNodes.Item(0).Text = vntNewCDataValue
                Case Else: ' Probably a Text section, delete Text and replace with CData section
                    oXNode.Text = ""
                    Set xmlCDATASection = m_xDoc.createCDATASection(vntNewCDataValue)
                    oXNode.appendChild xmlCDATASection
            End Select
        Case Else: ' Raise Error, if more than 1 child node
            GoTo errorhandler
    End Select

    Set xmlCDATASection = Nothing
    Set oXNode = Nothing
    WriteCDataSection = True
    Exit Function
errorhandler:
    WriteCDataSection = False
    Set xmlCDATASection = Nothing
    Set oXNode = Nothing
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
    On Error GoTo errorhandler
    
    Dim oXNode As MSXML2.IXMLDOMNode
    Dim xmlChildNode As MSXML2.IXMLDOMNode
    Dim xmlAttributeNode As MSXML2.IXMLDOMAttribute
    
    Select Case norRelationship
        Case norCHILD:
            Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString)
        Case norSIBLING:
            Set oXNode = m_xDoc.documentElement.selectSingleNode(strQueryString).parentNode
    End Select
   
    Set xmlChildNode = m_xDoc.createElement(strNewNodeName)
            
    If Not IsMissing(vntNewNodeValue) Then
        xmlChildNode.Text = vntNewNodeValue
    End If

    Set xmlAttributeNode = oXNode.appendChild(xmlChildNode)
            
    If Not IsMissing(strNewAttributeName) And Not IsMissing(vntNewAttributeValue) Then
        xmlAttributeNode.setAttribute strNewAttributeName, vntNewAttributeValue
    End If
            
    Set xmlAttributeNode = Nothing
    Set xmlChildNode = Nothing
    Set oXNode = Nothing
    InsertNode = True
    Exit Function
errorhandler:
    InsertNode = False
    Set xmlAttributeNode = Nothing
    Set xmlChildNode = Nothing
    Set oXNode = Nothing
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
    On Error GoTo errorhandler
    
    Dim xmlAttributeNode As MSXML2.IXMLDOMAttribute
    Set xmlAttributeNode = m_xDoc.documentElement.selectSingleNode(strQueryString).Attributes.removeNamedItem(strAttributeName)
    
    Set xmlAttributeNode = Nothing
    DeleteAttribute = True
    Exit Function
errorhandler:
    DeleteAttribute = False
    Set xmlAttributeNode = Nothing
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
    On Error GoTo errorhandler
    
    ' Create the DOMDocument
    Dim m_xDocReparse As Object
    Set m_xDocReparse = InitializeObject(CONST_DOMDOCUMENT)
    
    If m_xDocReparse.loadXML(m_xDoc.XML) = True Then
        Reparse = True
    Else
        Reparse = False
    End If
    
    Set m_xDocReparse = Nothing
    
    Exit Function
errorhandler:
    Reparse = False
    Set m_xDocReparse = Nothing
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
    On Error GoTo errorhandler
    
    m_xDoc.Save (strFileName)
    Save = True
    
    Exit Function
errorhandler:
    Save = False
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
    Dim xmlParseError As Object
    Set xmlParseError = m_xDoc.parseError
    
    MsgBox "ERROR loading XML" & vbCrLf & vbCrLf & _
        xmlParseError.errorCode & vbCrLf & _
        xmlParseError.reason & "Line " & xmlParseError.Line & _
        ", Position " & xmlParseError.linepos & vbCrLf & vbCrLf & _
        xmlParseError.srcText & vbCrLf _
        , vbExclamation, "XML Parse Error"
    
    Set xmlParseError = Nothing
End Sub

Private Sub Class_Terminate()
    Set m_xDoc = Nothing
End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' METHODS (PRIVATE) - END
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<










