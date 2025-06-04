Attribute VB_Name = "XMLProcess"
Option Explicit

Const NODE_TEXT = 3
Const NODE_CDATA_SECTION = 4



' -------------------------------------------------------------------------
' This function recurses through the XML tree given in the nodeIn parameter
' generally this would be the root node, at least that's how I call it but
' it would work to format any subtree as well
Public Function FormatNodeAsString(nodeIn As MSXML2.IXMLDOMNode, Optional intIndent As Integer = 0) As String
vbwProfiler.vbwProcIn 23
    Dim I As Integer
    Dim strXML As String
    Dim bChildIsTextOrCDATA As Boolean
    Dim intDent As Integer
    Dim nodemapAttributes As MSXML2.IXMLDOMNamedNodeMap
    Dim nodeTempChild As MSXML2.IXMLDOMNode
    Dim lngNodeType As Long
    Dim strNodeVal As String
    Dim bHasChildren As Boolean


vbwProfiler.vbwExecuteLine 223
    intDent = intIndent ' tracks the level of indentation required based on how deeply we've recursed
vbwProfiler.vbwExecuteLine 224
    lngNodeType = nodeIn.nodeType
vbwProfiler.vbwExecuteLine 225
    strNodeVal = nodeIn.nodeTypedValue
    ' in MSXML parser, i think its known that cdata and text nodes can't have children.  We could simplify this code using this assumption
vbwProfiler.vbwExecuteLine 226
    bHasChildren = NodeHasChildren(nodeIn)

    ' 1) get the start tag e.g. <string
vbwProfiler.vbwExecuteLine 227
    If lngNodeType = NODE_CDATA_SECTION Then
vbwProfiler.vbwExecuteLine 228
        strXML = "<![CDATA["
'vbwLine 229:    ElseIf lngNodeType <> NODE_TEXT Then
    ElseIf vbwProfiler.vbwExecuteLine(229) Or lngNodeType <> NODE_TEXT Then
vbwProfiler.vbwExecuteLine 230
        For I = 1 To intDent
vbwProfiler.vbwExecuteLine 231
            strXML = strXML & vbTab
vbwProfiler.vbwExecuteLine 232
        Next
vbwProfiler.vbwExecuteLine 233
        strXML = strXML & "<" & nodeIn.nodeName
    End If
vbwProfiler.vbwExecuteLine 234 'B

    ' 2) fill in the attribtes e.g. name = "Mike"
vbwProfiler.vbwExecuteLine 235
    Set nodemapAttributes = nodeIn.Attributes
vbwProfiler.vbwExecuteLine 236
    If Not (nodemapAttributes Is Nothing) Then
vbwProfiler.vbwExecuteLine 237
        For I = 0 To nodemapAttributes.length - 1
vbwProfiler.vbwExecuteLine 238
            Set nodeTempChild = nodemapAttributes.Item(I)
vbwProfiler.vbwExecuteLine 239
            strXML = strXML & " " & nodeTempChild.nodeName & "=" & """" & XMLEscapeString(nodeTempChild.nodeTypedValue) & """"
vbwProfiler.vbwExecuteLine 240
        Next
vbwProfiler.vbwExecuteLine 241
        Set nodeTempChild = Nothing
    End If
vbwProfiler.vbwExecuteLine 242 'B
vbwProfiler.vbwExecuteLine 243
    Set nodemapAttributes = Nothing

    ' 3) close the start tag
vbwProfiler.vbwExecuteLine 244
    If (lngNodeType <> NODE_CDATA_SECTION) And (lngNodeType <> NODE_TEXT) Then
vbwProfiler.vbwExecuteLine 245
        If bHasChildren Then
vbwProfiler.vbwExecuteLine 246
            strXML = strXML & ">"
        Else
vbwProfiler.vbwExecuteLine 247 'B
            ' if this tag has no children, we can actually close it right here and exit
            ' and without using the nodeName  e.g. /> instead of </string>
vbwProfiler.vbwExecuteLine 248
            strXML = strXML & "/>" & vbCrLf
vbwProfiler.vbwExecuteLine 249
            FormatNodeAsString = strXML
        End If
vbwProfiler.vbwExecuteLine 250 'B
    End If
vbwProfiler.vbwExecuteLine 251 'B

vbwProfiler.vbwExecuteLine 252
    If bHasChildren Then
vbwProfiler.vbwExecuteLine 253
        Set nodeTempChild = nodeIn.childNodes(0)
vbwProfiler.vbwExecuteLine 254
        If (nodeTempChild.nodeType = NODE_CDATA_SECTION) Or (nodeTempChild.nodeType = NODE_TEXT) Then
vbwProfiler.vbwExecuteLine 255
            bChildIsTextOrCDATA = True
        Else
vbwProfiler.vbwExecuteLine 256 'B
vbwProfiler.vbwExecuteLine 257
            strXML = strXML & vbCrLf
vbwProfiler.vbwExecuteLine 258
            bChildIsTextOrCDATA = False
        End If
vbwProfiler.vbwExecuteLine 259 'B

        ' for each child node of this current node we recursively call this function
'vbwLine 260:        Do Until nodeTempChild Is Nothing
        Do Until vbwProfiler.vbwExecuteLine(260) Or nodeTempChild Is Nothing
            ' every time we recurse, we increase the indent level by +1
vbwProfiler.vbwExecuteLine 261
            strXML = strXML & FormatNodeAsString(nodeTempChild, intDent + 1)
vbwProfiler.vbwExecuteLine 262
            Set nodeTempChild = nodeTempChild.nextSibling
vbwProfiler.vbwExecuteLine 263
        Loop
vbwProfiler.vbwExecuteLine 264
        Set nodeTempChild = Nothing
vbwProfiler.vbwExecuteLine 265
        If Not bChildIsTextOrCDATA Then
vbwProfiler.vbwExecuteLine 266
            For I = 1 To intDent
vbwProfiler.vbwExecuteLine 267
                strXML = strXML & vbTab
vbwProfiler.vbwExecuteLine 268
            Next
        End If
vbwProfiler.vbwExecuteLine 269 'B
vbwProfiler.vbwExecuteLine 270
        FormatNodeAsString = strXML & "</" & nodeIn.nodeName & ">" & vbCrLf
    Else
vbwProfiler.vbwExecuteLine 271 'B
vbwProfiler.vbwExecuteLine 272
        strXML = strXML & strNodeVal
vbwProfiler.vbwExecuteLine 273
        If lngNodeType = NODE_CDATA_SECTION Then
vbwProfiler.vbwExecuteLine 274
            FormatNodeAsString = strXML & "]]>"
'vbwLine 275:        ElseIf lngNodeType = NODE_TEXT Then
        ElseIf vbwProfiler.vbwExecuteLine(275) Or lngNodeType = NODE_TEXT Then
vbwProfiler.vbwExecuteLine 276
            FormatNodeAsString = strXML
        Else
vbwProfiler.vbwExecuteLine 277 'B
vbwProfiler.vbwExecuteLine 278
            FormatNodeAsString = strXML & vbCrLf
        End If
vbwProfiler.vbwExecuteLine 279 'B
    End If
vbwProfiler.vbwExecuteLine 280 'B
vbwProfiler.vbwProcOut 23
vbwProfiler.vbwExecuteLine 281
End Function
            
Private Function XMLEscapeString(ByVal XMLString As String) As String
vbwProfiler.vbwProcIn 24
vbwProfiler.vbwExecuteLine 282
    XMLString = Replace(XMLString, "&", "&amp;")
vbwProfiler.vbwExecuteLine 283
    XMLString = Replace(XMLString, Chr(34), "&quot;")
vbwProfiler.vbwExecuteLine 284
    XMLString = Replace(XMLString, "'", "&apos;")
vbwProfiler.vbwExecuteLine 285
    XMLString = Replace(XMLString, ">", "&gt;")
vbwProfiler.vbwExecuteLine 286
    XMLString = Replace(XMLString, "<", "&lt;")
vbwProfiler.vbwExecuteLine 287
    XMLEscapeString = XMLString
vbwProfiler.vbwProcOut 24
vbwProfiler.vbwExecuteLine 288
End Function

Private Function NodeHasChildren(nodeIn As MSXML2.IXMLDOMNode) As Boolean
vbwProfiler.vbwProcIn 25
    Dim nodelistChildren As MSXML2.IXMLDOMNodeList

vbwProfiler.vbwExecuteLine 289
    Set nodelistChildren = nodeIn.childNodes
vbwProfiler.vbwExecuteLine 290
    If nodelistChildren.length > 0 Then
vbwProfiler.vbwExecuteLine 291
        NodeHasChildren = True
    End If
vbwProfiler.vbwExecuteLine 292 'B
vbwProfiler.vbwProcOut 25
vbwProfiler.vbwExecuteLine 293
End Function

