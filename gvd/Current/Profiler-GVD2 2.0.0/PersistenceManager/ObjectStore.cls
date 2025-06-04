VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectStore"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum xopErrors
    errMSXMLerror = vbObjectError + 1
    errNoPeristentData
    errUnknownType
    errNoMultiDimArraysSupported
End Enum
Public Event RequestObject(ByVal Classname As String, ByVal DefPath As String, ByVal Def_ID As String, ByRef newObject As cIPersist)
Private m_XMLDom40 As MSXML2.DOMDocument40
Private m_colCache As Collection
Const SELECTION_LANGUAGE = "XPath"

' This function gets called ONCE per user initiated file deserialization.
' In other words, all child objects are deserialized using calls to DeserializeObject() and not Deserialize()
' To repeat -- this only looks for the first <Object> in the file.  It gets a handle and makes call to
' .Deserialize().  Luckily, children within parents are automatically deserialized for us.
Public Function Deserialize(ByRef xopSource As String, ByVal lngParentHandle As Long, ByRef sXPath As String, Optional obj As Object) As Object
vbwProfiler.vbwProcIn 1

vbwProfiler.vbwExecuteLine 1
    Set m_XMLDom40 = New MSXML2.DOMDocument40

vbwProfiler.vbwExecuteLine 2
    m_XMLDom40.async = False ' <-- forces that processing halt til the file is fully loaded by the parser
vbwProfiler.vbwExecuteLine 3
    m_XMLDom40.validateOnParse = True
vbwProfiler.vbwExecuteLine 4
    m_XMLDom40.resolveExternals = False
vbwProfiler.vbwExecuteLine 5
    m_XMLDom40.setProperty "SelectionLanguage", SELECTION_LANGUAGE

vbwProfiler.vbwExecuteLine 6
    If Left$(xopSource, 1) = "<" Then
vbwProfiler.vbwExecuteLine 7
        m_XMLDom40.loadXML xopSource 'todo: should have better code to determine whether its a file path or actual XML string
                                     '      Espeically considering that once we start using PAK type files, we will need to read strings.
                                    '       Of course, given user created .DEF's or CMP's, they will not be contained in PAK's
    Else
vbwProfiler.vbwExecuteLine 8 'B
vbwProfiler.vbwExecuteLine 9
        m_XMLDom40.Load xopSource
    End If
vbwProfiler.vbwExecuteLine 10 'B

vbwProfiler.vbwExecuteLine 11
    If m_XMLDom40.parseError.errorCode = 0 Then
vbwProfiler.vbwExecuteLine 12
        Set m_colCache = New Collection

        Dim s As String
        Dim oXNode As MSXML2.IXMLDOMNode
        Dim l As Long

        ' select the object specified by the XPATH string
vbwProfiler.vbwExecuteLine 13
        Set oXNode = m_XMLDom40.selectSingleNode("/ObjectStore/" & sXPath)

vbwProfiler.vbwExecuteLine 14
        If Not oXNode Is Nothing Then
            's = oXNode.selectSingleNode("string[@name='" & XML_ATTRIB_HANDLE & "']").nodeTypedValue  'todo: more constants need to be switched
vbwProfiler.vbwExecuteLine 15
            s = oXNode.Attributes.getNamedItem(XML_ATTRIB_HANDLE).nodeValue
vbwProfiler.vbwExecuteLine 16
            Debug.Print "ObjectStore:Deserialize() -- XML Node for Object with handle = '" & s & "' found.  Attempting to Deserialize"
vbwProfiler.vbwExecuteLine 17
            Debug.Assert s <> ""
vbwProfiler.vbwExecuteLine 18
            Set Deserialize = DeserializeObject(s, oXNode, obj)
vbwProfiler.vbwExecuteLine 19
            Set m_colCache = Nothing
        End If
vbwProfiler.vbwExecuteLine 20 'B
    Else
vbwProfiler.vbwExecuteLine 21 'B
vbwProfiler.vbwExecuteLine 22
        Set Deserialize = Nothing
        #If DEBUG_MODE Then
            Dim errMsg As String
vbwProfiler.vbwExecuteLine 23
            errMsg = "Error loading XOP-Source! Reason: " & m_XMLDom40.parseError.reason & "XML-Source (line/char " & m_XMLDom40.parseError.Line & "/" & m_XMLDom40.parseError.linepos & "): '" & m_XMLDom40.parseError.srcText & "'"
vbwProfiler.vbwExecuteLine 24
            err.Raise errMSXMLerror, "Load", errMsg
        #End If
    End If
vbwProfiler.vbwExecuteLine 25 'B

vbwProfiler.vbwExecuteLine 26
    Set m_XMLDom40 = Nothing
vbwProfiler.vbwProcOut 1
vbwProfiler.vbwExecuteLine 27
End Function


Public Function Serialize(ByVal rootObject As cIPersist, Optional ByVal returnXMLDOM As Boolean = True) As Variant
vbwProfiler.vbwProcIn 2
vbwProfiler.vbwExecuteLine 28
    Set m_XMLDom40 = New DOMDocument40
vbwProfiler.vbwExecuteLine 29
    Set m_colCache = New Collection
    Dim s As String

    'todo: need to figure out why there is a problem with reading the schema in the current path
    's = "<ObjectStore version=""1.0"" xmlns=""x-schema:D:\visualb\samples\XMLObject Persistance\samples\xop-schema.xml""/>"
    's = "<ObjectStore version=""1.0""/>"
vbwProfiler.vbwExecuteLine 30
    s = "<ObjectStore version=""1.0"" encoding=""windows-1252""/>"  'todo: Use constant for this string
    'NOTE: If this fails to load, check the minoccurs for objects in the object store schema.  make sure its 0.
vbwProfiler.vbwExecuteLine 31
    If m_XMLDom40.loadXML(s) Then
vbwProfiler.vbwExecuteLine 32
        m_XMLDom40.preserveWhiteSpace = True

vbwProfiler.vbwExecuteLine 33
        SerializeObject rootObject
vbwProfiler.vbwExecuteLine 34
        Set m_colCache = Nothing

vbwProfiler.vbwExecuteLine 35
        If returnXMLDOM Then
vbwProfiler.vbwExecuteLine 36
            Set Serialize = m_XMLDom40
        Else
vbwProfiler.vbwExecuteLine 37 'B
vbwProfiler.vbwExecuteLine 38
            Serialize = m_XMLDom40.xml ' returns the raw XML which we would want to save to file.
        End If
vbwProfiler.vbwExecuteLine 39 'B
    #If DEBUG_MODE Then
    Else
vbwProfiler.vbwExecuteLine 40 'B
        Dim errMsg As String
vbwProfiler.vbwExecuteLine 41
        errMsg = "Error loading XOP-Source! Reason: " & m_XMLDom40.parseError.reason & "XML-Source (line/char " & m_XMLDom40.parseError.Line & "/" & m_XMLDom40.parseError.linepos & "): '" & m_XMLDom40.parseError.srcText & "'"
vbwProfiler.vbwExecuteLine 42
        err.Raise errMSXMLerror, "Load", errMsg
vbwProfiler.vbwExecuteLine 43
        Set m_XMLDom40 = Nothing
    #End If
    End If
vbwProfiler.vbwExecuteLine 44 'B
vbwProfiler.vbwExecuteLine 45
    Set m_XMLDom40 = Nothing
vbwProfiler.vbwProcOut 2
vbwProfiler.vbwExecuteLine 46
End Function


'Public Property Get XML() As String
'    If m_XMLDom40 Is Nothing Then Err.Raise errNoPeristentData, "XML", "No persistent data has been loaded!"
'    XML = m_XMLDom40.XML
'End Property


'.------------------------------------------------------------------------------
'.  Function  : Friend Function DeserializeObject(ByVal strHandle As String) As Object
'.
'.  Parameters: strHandle as string -- ObjPtr & "_"
'.
'.  Returns   : Object
'.
'.  Comments  : Returns object which was instanced
'.
'.  Author    :
'.  Changed   : - Mike Joseph
'.------------------------------------------------------------------------------
'Friend Function DeserializeObject(ByVal strHandle As String, sXPath As String, Optional oPersistentObject As cIPersist) As Object
Friend Function DeserializeObject(ByVal strHandle As String, oXNode As MSXML2.IXMLDOMNode, Optional oPersistentObject As cIPersist) As Object
    'Dim oPersistentObject As cIPersist
    'Dim oXNode As MSXML2.IXMLDOMNode
vbwProfiler.vbwProcIn 3
    Dim sDefPath As String
    Dim sClassName As String
    Dim sGUID As String
    Dim lngMode As GVD_XML_TYPE

vbwProfiler.vbwExecuteLine 47
    On Error GoTo err

vbwProfiler.vbwExecuteLine 48
    If ItemIsCached(strHandle) Then
vbwProfiler.vbwExecuteLine 49
        Debug.Assert oPersistentObject Is Nothing  ' If the item is cached, oPersistantObject CANNOT be passed in.  Should be impossible... check with assert here anyway
        'Set oPersistentObject = m_colCache(strHandle)
vbwProfiler.vbwExecuteLine 50
        Debug.Print "ObjectStore:DeserializeObject() -- Item with strHandle = '" & strHandle & "' is cached, using cached instance..."
vbwProfiler.vbwExecuteLine 51
        Set DeserializeObject = m_colCache(strHandle)
vbwProfiler.vbwProcOut 3
vbwProfiler.vbwExecuteLine 52
        Exit Function
    Else
vbwProfiler.vbwExecuteLine 53 'B
vbwProfiler.vbwExecuteLine 54
        Debug.Print "ObjectStore:DeserializeObject() -- Item with strHandle = '" & strHandle & "' NOT cached.  Attempting to load..."
    End If
vbwProfiler.vbwExecuteLine 55 'B

vbwProfiler.vbwExecuteLine 56
    If Not oXNode Is Nothing Then
vbwProfiler.vbwExecuteLine 57
        If oPersistentObject Is Nothing Then
vbwProfiler.vbwExecuteLine 58
            lngMode = cmp

            ' When the user drags a component onto the tree from the listview, a call to deserialze that
            ' component is made.  We arrive here in this instance of ObjectStore and we read enough of the .cmp
            ' info to make a call to CreateObjectFromGUID which eventually gets an object instance returned from our oFactory.
            ' Within oFactory (and before returning here), a new instance of ObjectStore is created where we pass in the created object
            ' and then load the properties stored in the .def

            ' after the .def xml properties are loaded, the instanced object returns to our first instance of ObjectStore
            ' to load the properties in the .cmp file

            ' Now lets say i dont want to allow users to see unencrypted .def files and I also
            ' dont want them to know the classnames by peeking at the .cmp file
            ' i can remove the classname, but i have to add code at top of
            ' cFactory.CreateComponent() to load the def and read in the classname before trying to instance the object
            ' then i can delete the classname loading that follows...    The problem with that of course
            ' is it results in slower load times.  Is there somethign we can do about that though?  I REALLY
            ' want to limit saving .cmp's to just the actual user modified values and NOT classname.  What about GUID too?
            ' do we really need that in the .cmp's too?  If so, guid, defpath, classname should be in cIPersist.

            ' On the other hand, lets say we have GUID and it doesnt match the GUID of the xml def file...
            ' we can still at least compare classnames which would at least ensure that the .cmp was compatible
            ' with the DEF even if the GUID was not the same.

            ' Perhaps the solution is to use a ClassID.  This way the user has no clue what it is, but we arent
            ' spelling out the classnames either.
vbwProfiler.vbwExecuteLine 59
            Debug.Print "ObjectStore:DeserializeObject() -- Attempting to read ClassName.";
vbwProfiler.vbwExecuteLine 60
            sClassName = oXNode.selectSingleNode("string[@name='" & XML_NODE_CLASSNAME & "']").nodeTypedValue
vbwProfiler.vbwExecuteLine 61
            Debug.Print "   SUCCESS.  Classname = '" & sClassName & "'."
vbwProfiler.vbwExecuteLine 62
            Debug.Print "ObjectStore:DeserializeObject() -- Attempting to read Defpath.";
vbwProfiler.vbwExecuteLine 63
            sDefPath = oXNode.selectSingleNode("string[@name='" & XML_NODE_DEFPATH & "']").nodeTypedValue
vbwProfiler.vbwExecuteLine 64
            Debug.Print "   SUCCESS"
vbwProfiler.vbwExecuteLine 65
            Debug.Print "ObjectStore:DeserializeObject() -- Attempting to read GUID.";
vbwProfiler.vbwExecuteLine 66
            sGUID = oXNode.selectSingleNode("string[@name='" & XML_NODE_GUID & "']").nodeTypedValue
vbwProfiler.vbwExecuteLine 67
            Debug.Print "   SUCCESS"

            '   in the PAK file, we may be able to search the pak headers for the correct GUID.  A classname search would work too, but
            '   again, a problem would arrise if the original class was replaced (by the user) with a different one by accident

            '   By the way, part of what the Vehicle Factory Website would do is to manage user created DEF's into coherent packages.

            ' if an instanced object NOT already passed in, we must instance one
vbwProfiler.vbwExecuteLine 68
            Set oPersistentObject = CreateObjectFromGUID(sClassName, sDefPath, sGUID)
        Else
vbwProfiler.vbwExecuteLine 69 'B
vbwProfiler.vbwExecuteLine 70
            lngMode = DEF
vbwProfiler.vbwExecuteLine 71
            Debug.Print "ObjectStore:DeserializeObject() -- '" & oPersistentObject.Classname & "' already instanced.  Loading properties..."
        End If
vbwProfiler.vbwExecuteLine 72 'B

        Dim oProperties As clsObjProperties
vbwProfiler.vbwExecuteLine 73
        Set oProperties = New clsObjProperties
vbwProfiler.vbwExecuteLine 74
        oProperties.Constructor Me, oXNode
        ' cache it.
vbwProfiler.vbwExecuteLine 75
        m_colCache.Add oPersistentObject, strHandle

vbwProfiler.vbwExecuteLine 76
        If lngMode = DEF Then
            'todo: if i ever change the call to .loadproperties into a function and return boolean
            '      here is where I would add code to decide to fail gracefully if the load fails?
            '      Hrm... even still that only works on a per object basis but seems to still carry on
            '      and go to next object?  Do we want a global fail to unload entire vehicle?
vbwProfiler.vbwExecuteLine 77
            oPersistentObject.LoadProperties oProperties, DEF
        Else
vbwProfiler.vbwExecuteLine 78 'B
vbwProfiler.vbwExecuteLine 79
            oPersistentObject.LoadProperties oProperties, cmp
        End If
vbwProfiler.vbwExecuteLine 80 'B
vbwProfiler.vbwExecuteLine 81
        Set oProperties = Nothing
vbwProfiler.vbwExecuteLine 82
        Set DeserializeObject = oPersistentObject
    #If DEBUG_MODE Then
    Else
vbwProfiler.vbwExecuteLine 83 'B
vbwProfiler.vbwExecuteLine 84
        Debug.Print "ObjectStore:DeserializeObject() -- ERROR: Cannot find root XML Node for object with objectID = '" & strHandle
    #End If
    End If
vbwProfiler.vbwExecuteLine 85 'B
vbwProfiler.vbwProcOut 3
vbwProfiler.vbwExecuteLine 86
    Exit Function
err:
vbwProfiler.vbwExecuteLine 87
    Debug.Print "ObjectStore:DeserializeObject() -- Error # " & err.Number & " " & err.Description
vbwProfiler.vbwExecuteLine 88
    Debug.Print "ObjectStore:DeserializeObject() -- Item '" & sClassName & "' could not be deserialized.  Exiting Function."
vbwProfiler.vbwExecuteLine 89
    Set oProperties = Nothing
vbwProfiler.vbwExecuteLine 90
    Set oPersistentObject = Nothing
vbwProfiler.vbwExecuteLine 91
    Set DeserializeObject = Nothing
vbwProfiler.vbwProcOut 3
vbwProfiler.vbwExecuteLine 92
End Function


Friend Function SerializeObject(ByVal rootObject As cIPersist) As String
vbwProfiler.vbwProcIn 4
    Dim oPersistantObject As cIPersist
    Dim strHandle As String

vbwProfiler.vbwExecuteLine 93
    On Error Resume Next
vbwProfiler.vbwExecuteLine 94
    strHandle = ObjPtr(rootObject) & "_"

vbwProfiler.vbwExecuteLine 95
    Set oPersistantObject = m_colCache(strHandle)

vbwProfiler.vbwExecuteLine 96
    If err <> 0 Then
vbwProfiler.vbwExecuteLine 97
        On Error GoTo 0
vbwProfiler.vbwExecuteLine 98
        m_colCache.Add rootObject, strHandle

        Dim oXNode As MSXML2.IXMLDOMNode
        Dim oAttribute As MSXML2.IXMLDOMAttribute

        ' todo: might need to pass in XML_NODE_OBJECT so that we arent locked into serializing only a specific node type
vbwProfiler.vbwExecuteLine 99
        Set oXNode = m_XMLDom40.createNode(NODE_ELEMENT, XML_NODE_OBJECT, m_XMLDom40.documentElement.namespaceURI)

vbwProfiler.vbwExecuteLine 100
        Set oAttribute = m_XMLDom40.createAttribute(XML_ATTRIB_HANDLE)
vbwProfiler.vbwExecuteLine 101
        oAttribute.nodeValue = strHandle
vbwProfiler.vbwExecuteLine 102
        oXNode.Attributes.setNamedItem oAttribute

vbwProfiler.vbwExecuteLine 103
        Set oAttribute = m_XMLDom40.createAttribute(XML_NODE_GUID) 'todo: this needs to be a node!!!!
vbwProfiler.vbwExecuteLine 104
        oAttribute.nodeValue = rootObject.GUID
vbwProfiler.vbwExecuteLine 105
        oXNode.Attributes.setNamedItem oAttribute

vbwProfiler.vbwExecuteLine 106
        m_XMLDom40.documentElement.appendChild oXNode

        Dim oProperties As clsObjProperties
vbwProfiler.vbwExecuteLine 107
        Set oProperties = New clsObjProperties

vbwProfiler.vbwExecuteLine 108
        oProperties.Constructor Me, oXNode
vbwProfiler.vbwExecuteLine 109
        rootObject.StoreProperties oProperties

vbwProfiler.vbwExecuteLine 110
        Set oProperties = Nothing
    End If
vbwProfiler.vbwExecuteLine 111 'B

vbwProfiler.vbwExecuteLine 112
    SerializeObject = strHandle
vbwProfiler.vbwProcOut 4
vbwProfiler.vbwExecuteLine 113
End Function

'-------------------------------------------


Private Function CreateObjectFromGUID(ByVal sClassName As String, ByVal sDefPath As String, ByVal sGUID As String) As cIPersist
vbwProfiler.vbwProcIn 5
    Dim obj As cIPersist

    ' ask the application for an instance of the object´s class
vbwProfiler.vbwExecuteLine 114
    RaiseEvent RequestObject(sClassName, sDefPath, sGUID, obj)

    ' if no instance was provided try to create one
vbwProfiler.vbwExecuteLine 115
    If obj Is Nothing Then
        '''Oct.29.2002 MPJ - Changed mind.  We dont want to circumvent the system.  This should work the first time or fail.
        ''Oct.28.2002 MPJ - this is a last ditched effort to create the object.  I either want to remove this, or extensively check for when its actually used
        'On Error Resume Next
vbwProfiler.vbwExecuteLine 116
        Debug.Print TypeName(Me) & ":CreateObjectFromGUID() -- " & sClassName & " Not Set. Attempting alternate method..."

vbwProfiler.vbwExecuteLine 117
        Set obj = CreateObject(sClassName) ' note: class must be set to Multiuse or Global Multiuse for this to work.  Private or Not Publice Createable will cause error 429
                                           ' Is this still true if the object is passed sClassName = Vehicls.ClassName or is the problem
                                           ' specific to jsut passing in Classname?
                                            ' todo: So if i want to preserve hidden objects in my vehicles.dll so that
                                            ' competition cant copy/steal my interface implementation, then i have to add
                                            ' the PersistantManager directly into this dll rather than keep it as an independant DLL?
        #If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 118
            If Not obj Is Nothing Then
                'NOTE: Objects instanced here must not be graphed into the tree since they do not implement cINode.
                '      Read the "NOTE:" in cFactory.CreateComponent() which describes this.
vbwProfiler.vbwExecuteLine 119
                 Debug.Print TypeName(Me) & ":CreateObjectFromGUID() -- Success: " & sClassName & " instanced."
            End If
vbwProfiler.vbwExecuteLine 120 'B
        #End If
    End If
vbwProfiler.vbwExecuteLine 121 'B

vbwProfiler.vbwExecuteLine 122
    Set CreateObjectFromGUID = obj
vbwProfiler.vbwProcOut 5
vbwProfiler.vbwExecuteLine 123
End Function

Private Function ItemIsCached(strHandle As String) As Boolean
vbwProfiler.vbwProcIn 6
vbwProfiler.vbwExecuteLine 124
    On Error GoTo error

    Dim o As Object

vbwProfiler.vbwExecuteLine 125
    Set o = m_colCache(strHandle)
vbwProfiler.vbwExecuteLine 126
    Set o = Nothing
vbwProfiler.vbwExecuteLine 127
    ItemIsCached = True
vbwProfiler.vbwProcOut 6
vbwProfiler.vbwExecuteLine 128
    Exit Function
error:
    #If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 129
        If err.Number <> 5 Then
            ' we are expecting err 5, so only non err 5's need to be reported
vbwProfiler.vbwExecuteLine 130
            TypeName (Me) & ":ItemIsCahced() -- Error #" & err.Number & " " & err.Description
        End If
vbwProfiler.vbwExecuteLine 131 'B
    #End If
vbwProfiler.vbwExecuteLine 132
    ItemIsCached = False

vbwProfiler.vbwProcOut 6
vbwProfiler.vbwExecuteLine 133
End Function


