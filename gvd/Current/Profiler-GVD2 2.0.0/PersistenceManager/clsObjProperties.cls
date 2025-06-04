VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

 'todo: i borrowed this code for Persistence Manager... i need to rename all the public interfaces so that its not as obvious
' and delete this todo of course
Private m_os As ObjectStore
Private m_objNode As MSXML2.IXMLDOMNode



Public Function Load(ByRef propertyName As String) As Variant
vbwProfiler.vbwProcIn 11
    Dim propertyValue As Variant

vbwProfiler.vbwExecuteLine 138
    If LoadProperty(m_objNode, propertyName, propertyValue) Then
vbwProfiler.vbwExecuteLine 139
        If IsObject(propertyValue) Then
vbwProfiler.vbwExecuteLine 140
            If Not propertyValue Is Nothing Then
vbwProfiler.vbwExecuteLine 141
                Set Load = propertyValue
            Else
vbwProfiler.vbwExecuteLine 142 'B
vbwProfiler.vbwExecuteLine 143
                Set Load = propertyValue
                'todo we need to raise an error here?  We certainly cant let it assign Nothing and possibly the
                ' calling app doesnt realize it til it later tries to access the object which is not initialized
            End If
vbwProfiler.vbwExecuteLine 144 'B
        Else
vbwProfiler.vbwExecuteLine 145 'B
           ' Debug.Print ; "Value = " & propertyValue
vbwProfiler.vbwExecuteLine 146
            Load = propertyValue
        End If
vbwProfiler.vbwExecuteLine 147 'B
    End If
vbwProfiler.vbwExecuteLine 148 'B
vbwProfiler.vbwProcOut 11
vbwProfiler.vbwExecuteLine 149
End Function

Public Sub Store(ByRef propertyName As String, ByVal value As Variant)
vbwProfiler.vbwProcIn 12
vbwProfiler.vbwExecuteLine 150
    StoreProperty m_objNode, propertyName, value
vbwProfiler.vbwProcOut 12
vbwProfiler.vbwExecuteLine 151
End Sub


'---------------------------
Friend Sub Constructor(ByRef os As ObjectStore, ByRef oNode As MSXML2.IXMLDOMNode)
    ' a reference to the ObjectStore is needed ONLY if this object requires us to
    ' deserialize OR serialize  a member/child object or to
vbwProfiler.vbwProcIn 13
vbwProfiler.vbwExecuteLine 152
    Set m_os = os

    ' when deserializing, this node contains all the XML about this object
    ' when serializing, this is the node where we store everything so that eventually, the calling
    ' function can take the node and store it to file.
vbwProfiler.vbwExecuteLine 153
    Set m_objNode = oNode
vbwProfiler.vbwProcOut 13
vbwProfiler.vbwExecuteLine 154
End Sub


'---------------------------
Private Sub StoreProperty(ByRef parentNode As MSXML2.IXMLDOMNode, ByRef propertyName As String, ByVal value As Variant)
vbwProfiler.vbwProcIn 14
vbwProfiler.vbwExecuteLine 155
    If InStr(propertyName, ".") = 0 Then
        Dim p As MSXML2.IXMLDOMNode
        Dim nodeName As String
        Dim oAttribute As MSXML2.IXMLDOMAttribute

vbwProfiler.vbwExecuteLine 156
        Const TYPE_STRING = "String"
vbwProfiler.vbwExecuteLine 157
        Const TYPE_BYTE = "Byte"
vbwProfiler.vbwExecuteLine 158
        Const TYPE_INTEGER = "Integer"
vbwProfiler.vbwExecuteLine 159
        Const TYPE_LONG = "Long"
vbwProfiler.vbwExecuteLine 160
        Const TYPE_SINGLE = "Single"
vbwProfiler.vbwExecuteLine 161
        Const TYPE_DOUBLE = "Double"
vbwProfiler.vbwExecuteLine 162
        Const TYPE_CURRENCY = "Currency"
vbwProfiler.vbwExecuteLine 163
        Const TYPE_DECIMAL = "Decimal"
vbwProfiler.vbwExecuteLine 164
        Const TYPE_BOOLEAN = "Boolean"
vbwProfiler.vbwExecuteLine 165
        Const TYPE_DATE = "Date"

vbwProfiler.vbwExecuteLine 166
        Select Case TypeName(value)
'vbwLine 167:            Case TYPE_STRING
            Case IIf(vbwProfiler.vbwExecuteLine(167), VBWPROFILER_EMPTY, _
        TYPE_STRING)
vbwProfiler.vbwExecuteLine 168
                nodeName = XML_NODETYPE_STRING
vbwProfiler.vbwExecuteLine 169
                GoSub StoreSimpleTypeValue

vbwProfiler.vbwExecuteLine 170
                Set oAttribute = parentNode.ownerDocument.createAttribute(XML_ATTRIB_XMLSPACE)
vbwProfiler.vbwExecuteLine 171
                oAttribute.nodeValue = XML_ATTRIB_PRESERVE
vbwProfiler.vbwExecuteLine 172
                p.Attributes.setNamedItem oAttribute

'vbwLine 173:            Case TYPE_BYTE, TYPE_INTEGER, TYPE_LONG
            Case IIf(vbwProfiler.vbwExecuteLine(173), VBWPROFILER_EMPTY, _
        TYPE_BYTE), TYPE_INTEGER, TYPE_LONG
vbwProfiler.vbwExecuteLine 174
                nodeName = XML_NODETYPE_INTEGER
vbwProfiler.vbwExecuteLine 175
                GoSub StoreSimpleTypeValue
'vbwLine 176:            Case TYPE_SINGLE, TYPE_DOUBLE, TYPE_CURRENCY, TYPE_DECIMAL
            Case IIf(vbwProfiler.vbwExecuteLine(176), VBWPROFILER_EMPTY, _
        TYPE_SINGLE), TYPE_DOUBLE, TYPE_CURRENCY, TYPE_DECIMAL
vbwProfiler.vbwExecuteLine 177
                value = Replace(value, ",", ".")
                    ' the float value has to be stored with a decimal point regardless
                    ' of the locale!
vbwProfiler.vbwExecuteLine 178
                nodeName = XML_NODETYPE_FLOAT
vbwProfiler.vbwExecuteLine 179
                GoSub StoreSimpleTypeValue
'vbwLine 180:            Case TYPE_BOOLEAN
            Case IIf(vbwProfiler.vbwExecuteLine(180), VBWPROFILER_EMPTY, _
        TYPE_BOOLEAN)
vbwProfiler.vbwExecuteLine 181
                value = Abs(value)
vbwProfiler.vbwExecuteLine 182
                nodeName = XML_NODETYPE_BOOL
vbwProfiler.vbwExecuteLine 183
                GoSub StoreSimpleTypeValue
'vbwLine 184:            Case TYPE_DATE
            Case IIf(vbwProfiler.vbwExecuteLine(184), VBWPROFILER_EMPTY, _
        TYPE_DATE)
vbwProfiler.vbwExecuteLine 185
                value = Format(value, "YYYY-MM-DDTHH:MM:SS")
vbwProfiler.vbwExecuteLine 186
                nodeName = XML_NODETYPE_DATETIME
vbwProfiler.vbwExecuteLine 187
                GoSub StoreSimpleTypeValue

            Case Else
vbwProfiler.vbwExecuteLine 188 'B
vbwProfiler.vbwExecuteLine 189
                If IsObject(value) Then
vbwProfiler.vbwExecuteLine 190
                    value = m_os.SerializeObject(value)
vbwProfiler.vbwExecuteLine 191
                    nodeName = XML_NODETYPE_OBJECTREF
vbwProfiler.vbwExecuteLine 192
                    GoSub StoreSimpleTypeValue
                Else
vbwProfiler.vbwExecuteLine 193 'B
vbwProfiler.vbwExecuteLine 194
                    If InStr(TypeName(value), "()") > 0 Then
                        Dim u%
vbwProfiler.vbwExecuteLine 195
                        On Error Resume Next
vbwProfiler.vbwExecuteLine 196
                        u = UBound(value, 2)
vbwProfiler.vbwExecuteLine 197
                        If err <> 0 Then
vbwProfiler.vbwExecuteLine 198
                            On Error GoTo 0
                            ' todo: need to handle various types of arrays here
vbwProfiler.vbwExecuteLine 199
                            Store_V_Array parentNode, propertyName, value
                        Else
vbwProfiler.vbwExecuteLine 200 'B
vbwProfiler.vbwExecuteLine 201
                            On Error GoTo 0
vbwProfiler.vbwExecuteLine 202
                            err.Raise errNoMultiDimArraysSupported, "StoreProperty", "Property '" & propertyName & "': No multidimensional arrays supported!"
                        End If
vbwProfiler.vbwExecuteLine 203 'B
                    Else
vbwProfiler.vbwExecuteLine 204 'B
vbwProfiler.vbwExecuteLine 205
                        err.Raise errUnknownType, "StoreProperty", "Property '" & propertyName & "' is of unknown type '" & TypeName(value) & "'!"
                    End If
vbwProfiler.vbwExecuteLine 206 'B
                End If
vbwProfiler.vbwExecuteLine 207 'B
        End Select
vbwProfiler.vbwExecuteLine 208 'B

    Else
vbwProfiler.vbwExecuteLine 209 'B
vbwProfiler.vbwExecuteLine 210
        StoreStructMember parentNode, propertyName, value
    End If
vbwProfiler.vbwExecuteLine 211 'B

vbwProfiler.vbwProcOut 14
vbwProfiler.vbwExecuteLine 212
    Exit Sub

StoreSimpleTypeValue:
vbwProfiler.vbwExecuteLine 213
    Set p = parentNode.ownerDocument.createNode(NODE_ELEMENT, nodeName, parentNode.namespaceURI)
vbwProfiler.vbwExecuteLine 214
        Set oAttribute = parentNode.ownerDocument.createAttribute(XML_ATTRIB_NAME)
vbwProfiler.vbwExecuteLine 215
        oAttribute.nodeValue = propertyName
vbwProfiler.vbwExecuteLine 216
        p.Attributes.setNamedItem oAttribute

vbwProfiler.vbwExecuteLine 217
    p.appendChild parentNode.ownerDocument.createTextNode(value)

vbwProfiler.vbwExecuteLine 218
    parentNode.appendChild p
vbwProfiler.vbwExecuteLine 219
    Return
vbwProfiler.vbwProcOut 14
vbwProfiler.vbwExecuteLine 220
End Sub


Private Function LoadProperty(ByRef parentNode As MSXML2.IXMLDOMNode, ByRef propertyName As String, ByRef propertyValue As Variant) As Boolean
vbwProfiler.vbwProcIn 15

vbwProfiler.vbwExecuteLine 221
    LoadProperty = True

vbwProfiler.vbwExecuteLine 222
    If InStr(propertyName, ".") = 0 Then
vbwProfiler.vbwExecuteLine 223
        Debug.Print "clsObjProperties:LoadProperty() -- Attempting to load property '" & propertyName & "'.";
        Dim pNode As MSXML2.IXMLDOMNode

vbwProfiler.vbwExecuteLine 224
        If Not parentNode Is Nothing Then
vbwProfiler.vbwExecuteLine 225
            Set pNode = parentNode.selectSingleNode("*[@name='" & propertyName & "']")

vbwProfiler.vbwExecuteLine 226
            If Not pNode Is Nothing Then
vbwProfiler.vbwExecuteLine 227
                Select Case pNode.nodeName
'vbwLine 228:                    Case XML_NODETYPE_DATETIME
                    Case IIf(vbwProfiler.vbwExecuteLine(228), VBWPROFILER_EMPTY, _
        XML_NODETYPE_DATETIME)
                        ' MSXML does not support the dateTime.iso8601 datatype so we have to
                        ' convert the data ourselves
                        Dim dtText As String
vbwProfiler.vbwExecuteLine 229
                        dtText = pNode.Text
vbwProfiler.vbwExecuteLine 230
                        propertyValue = DateSerial(CInt(Left$(dtText, 4)), CInt(Mid$(dtText, 6, 2)), CInt(Mid$(dtText, 9, 2))) + _
                               TimeSerial(CInt(Mid$(dtText, 12, 2)), CInt(Mid$(dtText, 15, 2)), CInt(Mid$(dtText, 18, 2)))

'vbwLine 231:                    Case XML_NODETYPE_OBJ_ARRAY
                    Case IIf(vbwProfiler.vbwExecuteLine(231), VBWPROFILER_EMPTY, _
        XML_NODETYPE_OBJ_ARRAY)
vbwProfiler.vbwExecuteLine 232
                            propertyValue = Load_V_Array(pNode) ' handles both variants and objects

                     'todo: to support these typs... need to create seperate loadarraay functions
                     '      so far only have S_ARRAY and V_ARRAY
                     'NOTE: The reason these various array loaders is necessary is because you cannot assign a variant array return value
                     '      to any other type except for Object arrays or other Variant arrays!  Pain in the ass really
                     '      since the code is nearly identcile.  I tried passing in integer() into the variant to see if that
                     '      would change the nature of the variant (sometimes in watch expression they show up as string/variant or integer/variant
                     '      so i thought that might make a difference, it doesnt.  We NEED a different array load depending on the array type
                     '      If only we had C++ style templates
'vbwLine 233:                    Case XML_NODETYPE_S_ARRAY
                    Case IIf(vbwProfiler.vbwExecuteLine(233), VBWPROFILER_EMPTY, _
        XML_NODETYPE_S_ARRAY)
vbwProfiler.vbwExecuteLine 234
                        propertyValue = Load_S_Array(pNode)
'
'                    Case XML_NODETYPE_F_ARRAY
'                        propertyValue = LoadArray(pNode)
'
'vbwLine 235:                    Case XML_NODETYPE_I_ARRAY
                    Case IIf(vbwProfiler.vbwExecuteLine(235), VBWPROFILER_EMPTY, _
        XML_NODETYPE_I_ARRAY)
vbwProfiler.vbwExecuteLine 236
                        propertyValue = Load_I_Array(pNode)
'                    Case XML_NODETYPE_L_ARRAY
'                        propertyValue = LoadArray(pNode)

'vbwLine 237:                    Case XML_NODETYPE_F_ARRAY_2D
                    Case IIf(vbwProfiler.vbwExecuteLine(237), VBWPROFILER_EMPTY, _
        XML_NODETYPE_F_ARRAY_2D)
vbwProfiler.vbwExecuteLine 238
                        propertyValue = load_f_2D_Array(pNode)

'vbwLine 239:                    Case XML_NODETYPE_F_TABLE
                    Case IIf(vbwProfiler.vbwExecuteLine(239), VBWPROFILER_EMPTY, _
        XML_NODETYPE_F_TABLE)
                        'load the 2d array of 'singles' as normal
                        Dim sng() As Single
vbwProfiler.vbwExecuteLine 240
                        sng = load_f_2D_Array(pNode)
                        ' then create our heap based table of appropriate size and copy this array data to it
                        Dim lPtr As Long
                        Dim tbl As TABLE
vbwProfiler.vbwExecuteLine 241
                        lPtr = createTable(UBound(sng, 1) + 1, UBound(sng, 2) + 1)
vbwProfiler.vbwExecuteLine 242
                        CopyMemory ByVal (lPtr + LenB(tbl)), sng(0, 0), (UBound(sng, 1) + 1) * (UBound(sng, 2) + 1) * 4

'                        #If DEBUG_MODE Then
'                            ' a few tests here to see if its working by returning the proper value
'                            ' note these test will be deleted very soon since they ONLY work on a specific test
'                            ' component xml
'                            Dim sngResult As Single
'
'                            getTableItemValue lPtr, 0, 0, sngResult
'                            Debug.Assert sngResult = 0.001
'                            Debug.Print "PersistenceManager:LoadProperty() -- assertion passed..."
'                            getTableItemValue lPtr, 0, 2, sngResult
'                            Debug.Assert sngResult = 0.2
'                            Debug.Print "PersistenceManager:LoadProperty() -- assertion passed..."
'                            getTableItemValue lPtr, 1, 1, sngResult
'                            Debug.Assert sngResult = 33.3
'                            Debug.Print "PersistenceManager:LoadProperty() -- assertion passed..."
'                            getTableItemValue lPtr, 1, 2, sngResult
'                            Debug.Assert sngResult = 44.4
'                            Debug.Print "PersistenceManager:LoadProperty() -- assertion passed..."
'                        #End If
                        ' return a pointer to the table
vbwProfiler.vbwExecuteLine 243
                        propertyValue = lPtr

'vbwLine 244:                    Case XML_NODETYPE_OBJECTREF
                    Case IIf(vbwProfiler.vbwExecuteLine(244), VBWPROFILER_EMPTY, _
        XML_NODETYPE_OBJECTREF)
                        ' Note that this doesnt call the top level .Deserialize... it goes straight to DeserializeObject
                        Dim sHandle As String
vbwProfiler.vbwExecuteLine 245
                        sHandle = pNode.nodeTypedValue
                        'todo: this next line is a horrible HACK which will probably not work all the time.
                        '      we need to effectively get back to the \ObjectStore\ level error free so we can selectsinglenode and
                        '      sure to find the object
vbwProfiler.vbwExecuteLine 246
                        Debug.Print ""
vbwProfiler.vbwExecuteLine 247
                        Debug.Print vbTab;
vbwProfiler.vbwExecuteLine 248
                        Set pNode = parentNode.parentNode.selectSingleNode("object[@handle='" & sHandle & "']")
vbwProfiler.vbwExecuteLine 249
                        Set propertyValue = m_os.DeserializeObject(sHandle, pNode)

'vbwLine 250:                    Case XML_NODETYPE_STRUCT
                    Case IIf(vbwProfiler.vbwExecuteLine(250), VBWPROFILER_EMPTY, _
        XML_NODETYPE_STRUCT)
vbwProfiler.vbwExecuteLine 251
                        propertyValue = LoadStructMember(pNode, propertyName, propertyValue)

                    Case Else
vbwProfiler.vbwExecuteLine 252 'B
                        'todo: i dont know if i like this... case else should be used as an error trap.
                        '      WE DO NEED TO ABORT LOADING FILES THAT ARE CORRUPT!
vbwProfiler.vbwExecuteLine 253
                        propertyValue = pNode.nodeTypedValue
                End Select
vbwProfiler.vbwExecuteLine 254 'B
vbwProfiler.vbwExecuteLine 255
                Debug.Print "   SUCCESS."
            Else
vbwProfiler.vbwExecuteLine 256 'B
vbwProfiler.vbwExecuteLine 257
                LoadProperty = False
vbwProfiler.vbwExecuteLine 258
                Debug.Print "   FAILED."
            End If
vbwProfiler.vbwExecuteLine 259 'B
        Else
vbwProfiler.vbwExecuteLine 260 'B
vbwProfiler.vbwExecuteLine 261
            LoadProperty = False
vbwProfiler.vbwExecuteLine 262
            Debug.Print "clsObjProperties:LoadProperty() -- Error: parentNode not set!"
        End If
vbwProfiler.vbwExecuteLine 263 'B
    Else
vbwProfiler.vbwExecuteLine 264 'B
vbwProfiler.vbwExecuteLine 265
        LoadProperty = LoadStructMember(parentNode, propertyName, propertyValue)
    End If
vbwProfiler.vbwExecuteLine 266 'B
vbwProfiler.vbwProcOut 15
vbwProfiler.vbwExecuteLine 267
End Function

Private Function load_f_2D_Array(ByRef parentNode As MSXML2.IXMLDOMNode) As Single()
vbwProfiler.vbwProcIn 16

'    Row (TL), Column(various)
'
'     0, 0 = Datatype
'     0, 1 = Weight1
'     0, 2 = Weight2
'     0, 3 = Cost
'     0, 4 = Volume
'
'     1, 0
'     1, 1
'     1, 2
'     1, 3
'     1, 4

    ' In XML these are named as follows:  "0,0" "0,1" "0,2" "0,3" "0,4" "1,0" "1,1" "1,2" ... etc

    Dim f() As Single
    Dim rl As Long, ru As Long, cl As Long, cu As Long
    Dim i As Long, j As Long
    Dim sNodeName As String

vbwProfiler.vbwExecuteLine 268
    rl = parentNode.Attributes.getNamedItem(XML_ATTRIB_ROWLOWERBOUND).nodeValue
vbwProfiler.vbwExecuteLine 269
    ru = parentNode.Attributes.getNamedItem(XML_ATTRIB_ROWUPPERBOUND).nodeValue
vbwProfiler.vbwExecuteLine 270
    cl = parentNode.Attributes.getNamedItem(XML_ATTRIB_COLUMNLOWERBOUND).nodeValue
vbwProfiler.vbwExecuteLine 271
    cu = parentNode.Attributes.getNamedItem(XML_ATTRIB_COLUMNUPPERBOUND).nodeValue

vbwProfiler.vbwExecuteLine 272
    ReDim f(rl To ru, cl To cu) As Single

vbwProfiler.vbwExecuteLine 273
    For i = rl To ru
vbwProfiler.vbwExecuteLine 274
        For j = cl To cu
vbwProfiler.vbwExecuteLine 275
            sNodeName = CStr(i) & "," & j
vbwProfiler.vbwExecuteLine 276
            LoadProperty parentNode, sNodeName, f(i, j)
vbwProfiler.vbwExecuteLine 277
        Next
vbwProfiler.vbwExecuteLine 278
    Next
vbwProfiler.vbwExecuteLine 279
    load_f_2D_Array = f
vbwProfiler.vbwProcOut 16
vbwProfiler.vbwExecuteLine 280
End Function


Private Function Load_S_Array(ByRef parentNode As MSXML2.IXMLDOMNode) As String()
vbwProfiler.vbwProcIn 17
    Dim lb As Long
    Dim ub As Long
    Dim i As Long
    Dim s() As String
vbwProfiler.vbwExecuteLine 281
    lb = parentNode.Attributes.getNamedItem(XML_ATTRIB_LOWERBOUND).nodeValue
vbwProfiler.vbwExecuteLine 282
    ub = parentNode.Attributes.getNamedItem(XML_ATTRIB_UPPERBOUND).nodeValue
vbwProfiler.vbwExecuteLine 283
    ReDim s(lb To ub)
vbwProfiler.vbwExecuteLine 284
    For i = lb To ub
vbwProfiler.vbwExecuteLine 285
        LoadProperty parentNode, CStr(i), s(i)
vbwProfiler.vbwExecuteLine 286
    Next
vbwProfiler.vbwExecuteLine 287
    Load_S_Array = s
vbwProfiler.vbwProcOut 17
vbwProfiler.vbwExecuteLine 288
End Function

Private Function Load_I_Array(ByRef parentNode As MSXML2.IXMLDOMNode) As Integer()
vbwProfiler.vbwProcIn 18
    Dim lb As Long
    Dim ub As Long
    Dim i As Long
    Dim l() As Integer
vbwProfiler.vbwExecuteLine 289
    lb = parentNode.Attributes.getNamedItem(XML_ATTRIB_LOWERBOUND).nodeValue
vbwProfiler.vbwExecuteLine 290
    ub = parentNode.Attributes.getNamedItem(XML_ATTRIB_UPPERBOUND).nodeValue
vbwProfiler.vbwExecuteLine 291
    ReDim l(lb To ub)
vbwProfiler.vbwExecuteLine 292
    For i = lb To ub
vbwProfiler.vbwExecuteLine 293
        LoadProperty parentNode, CStr(i), l(i)
vbwProfiler.vbwExecuteLine 294
    Next
vbwProfiler.vbwExecuteLine 295
    Load_I_Array = l
vbwProfiler.vbwProcOut 18
vbwProfiler.vbwExecuteLine 296
End Function

Private Function Load_V_Array(ByRef parentNode As MSXML2.IXMLDOMNode) As Variant
vbwProfiler.vbwProcIn 19
    Dim lb As Long
    Dim ub As Long
    Dim i As Long
    Dim v() As Variant
vbwProfiler.vbwExecuteLine 297
    lb = parentNode.Attributes.getNamedItem(XML_ATTRIB_LOWERBOUND).nodeValue
vbwProfiler.vbwExecuteLine 298
    ub = parentNode.Attributes.getNamedItem(XML_ATTRIB_UPPERBOUND).nodeValue
vbwProfiler.vbwExecuteLine 299
    ReDim v(lb To ub)
vbwProfiler.vbwExecuteLine 300
    For i = lb To ub
vbwProfiler.vbwExecuteLine 301
        LoadProperty parentNode, CStr(i), v(i)
vbwProfiler.vbwExecuteLine 302
    Next
vbwProfiler.vbwExecuteLine 303
    Load_V_Array = v
vbwProfiler.vbwProcOut 19
vbwProfiler.vbwExecuteLine 304
End Function

Private Function LoadPropertyList(ByRef parentNode As MSXML2.IXMLDOMNode) As Variant
vbwProfiler.vbwProcIn 20
    Dim lb As Long
    Dim ub As Long
vbwProfiler.vbwExecuteLine 305
    lb = parentNode.Attributes.getNamedItem(XML_ATTRIB_LOWERBOUND).nodeValue
vbwProfiler.vbwExecuteLine 306
    ub = parentNode.Attributes.getNamedItem(XML_ATTRIB_UPPERBOUND).nodeValue
vbwProfiler.vbwProcOut 20
vbwProfiler.vbwExecuteLine 307
End Function

Private Sub Store_V_Array(ByRef parentNode As MSXML2.IXMLDOMNode, ByRef propertyName As String, ByVal value As Variant)
vbwProfiler.vbwProcIn 21
    Dim i As Long
    Dim arrayNode As MSXML2.IXMLDOMNode
    Dim oAttribute As MSXML2.IXMLDOMAttribute

vbwProfiler.vbwExecuteLine 308
    Set arrayNode = parentNode.ownerDocument.createNode(NODE_ELEMENT, XML_NODETYPE_V_ARRAY, parentNode.namespaceURI)
vbwProfiler.vbwExecuteLine 309
    Set oAttribute = parentNode.ownerDocument.createAttribute(XML_ATTRIB_NAME)
vbwProfiler.vbwExecuteLine 310
        oAttribute.nodeValue = propertyName
vbwProfiler.vbwExecuteLine 311
        arrayNode.Attributes.setNamedItem oAttribute
vbwProfiler.vbwExecuteLine 312
    Set oAttribute = parentNode.ownerDocument.createAttribute(XML_ATTRIB_LOWERBOUND)
vbwProfiler.vbwExecuteLine 313
        oAttribute.nodeValue = LBound(value)
vbwProfiler.vbwExecuteLine 314
        arrayNode.Attributes.setNamedItem oAttribute
vbwProfiler.vbwExecuteLine 315
    Set oAttribute = parentNode.ownerDocument.createAttribute(XML_ATTRIB_UPPERBOUND)
vbwProfiler.vbwExecuteLine 316
        oAttribute.nodeValue = UBound(value)
vbwProfiler.vbwExecuteLine 317
        arrayNode.Attributes.setNamedItem oAttribute

vbwProfiler.vbwExecuteLine 318
    parentNode.appendChild arrayNode
vbwProfiler.vbwExecuteLine 319
    For i = LBound(value) To UBound(value)
vbwProfiler.vbwExecuteLine 320
        StoreProperty arrayNode, CStr(i), value(i)
vbwProfiler.vbwExecuteLine 321
    Next
vbwProfiler.vbwProcOut 21
vbwProfiler.vbwExecuteLine 322
End Sub

Private Sub StoreStructMember(ByRef parentNode As MSXML2.IXMLDOMNode, ByRef propertyName As String, ByVal structMemberValue As Variant)
vbwProfiler.vbwProcIn 22
    Dim i As Long
    Dim structName As String
    Dim structNode As MSXML2.IXMLDOMNode

vbwProfiler.vbwExecuteLine 323
    i = InStr(propertyName, ".")
vbwProfiler.vbwExecuteLine 324
    structName = Left$(propertyName, i - 1)
vbwProfiler.vbwExecuteLine 325
    propertyName = Mid$(propertyName, i + 1)

vbwProfiler.vbwExecuteLine 326
    Set structNode = parentNode.selectSingleNode("struct[@name='" & structName & "']")
vbwProfiler.vbwExecuteLine 327
    If structNode Is Nothing Then
        Dim oAttribute As MSXML2.IXMLDOMAttribute
vbwProfiler.vbwExecuteLine 328
        Set structNode = parentNode.ownerDocument.createNode(NODE_ELEMENT, XML_NODETYPE_STRUCT, parentNode.namespaceURI)
vbwProfiler.vbwExecuteLine 329
        Set oAttribute = parentNode.ownerDocument.createAttribute(XML_ATTRIB_NAME)
vbwProfiler.vbwExecuteLine 330
            oAttribute.nodeValue = structName
vbwProfiler.vbwExecuteLine 331
            structNode.Attributes.setNamedItem oAttribute
vbwProfiler.vbwExecuteLine 332
        parentNode.appendChild structNode
    End If
vbwProfiler.vbwExecuteLine 333 'B
vbwProfiler.vbwExecuteLine 334
    StoreProperty structNode, propertyName, structMemberValue
vbwProfiler.vbwProcOut 22
vbwProfiler.vbwExecuteLine 335
End Sub

Private Function LoadStructMember(ByRef parentNode As MSXML2.IXMLDOMNode, ByRef propertyName As String, ByRef structMemberValue As Variant) As Boolean
vbwProfiler.vbwProcIn 23
    Dim i As Long
    Dim structName As String
    Dim structNode As MSXML2.IXMLDOMNode

vbwProfiler.vbwExecuteLine 336
    i = InStr(propertyName, ".")
vbwProfiler.vbwExecuteLine 337
    structName = Left$(propertyName, i - 1)
vbwProfiler.vbwExecuteLine 338
    propertyName = Mid$(propertyName, i + 1)

vbwProfiler.vbwExecuteLine 339
    Set structNode = parentNode.selectSingleNode("struct[@name='" & structName & "']")
vbwProfiler.vbwExecuteLine 340
    LoadProperty structNode, propertyName, structMemberValue
    'todo: for now lets assume it always succeeds. We dont use structs (yet) in gvd so we shouldnt ever enter this funct
vbwProfiler.vbwExecuteLine 341
    LoadStructMember = True
vbwProfiler.vbwProcOut 23
vbwProfiler.vbwExecuteLine 342
End Function



