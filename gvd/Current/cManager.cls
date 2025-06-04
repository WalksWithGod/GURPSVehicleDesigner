VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents m_ObjectStore As PersistenceManager.ObjectStore
Attribute m_ObjectStore.VB_VarHelpID = -1
Private m_colVehicles As Collection    '10/28/02 MPJ  -- We are now supporting multiple vehicles
Private m_oFactory As Vehicles.cFactory  ' only one factory is needed however...

Private m_oVehicles As Collection

'  note: we are not going to fuck around with internet shit right now.  Instead
'  we are going to move all our client accessible commands here.  The user may
'  be able to instance local versions of cINode and cPropertyItem, but thats about it.
'  They will not have access to cIComponent or other such interfaces as they are not
'  allowed to access components directly.

'  Eventually we will incorp localhost loopback, but that will be after gvd 2.0 is done.

Private Sub Class_Initialize()
    Set m_oFactory = New Vehicles.cFactory
    Set m_colVehicles = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_oFactory = Nothing
    ' Destroy the collection and all objects in it
    Dim oVehicle As cVehicle
    For Each oVehicle In m_colVehicles
        Set oVehicle = Nothing
    Next
    Set m_colVehicles = Nothing
End Sub

Public Function createVehicle(sFilePath As String) As Long
    Dim lptr As Long
    Dim sKey As String
    Dim oVeh As Vehicles.cINode
    Dim o2 As Object
    
    Set m_ObjectStore = New PersistenceManager.ObjectStore
    
    Set o2 = m_ObjectStore.Deserialize(sFilePath, 0&, XML_NODE_OBJECT & "[@handle='0_']")
    If Not o2 Is Nothing Then
        Set oVeh = o2
        lptr = oVeh.Handle
        
        sKey = CStr(lptr) & "_"  ' todo: make this function from gvd project available here too KeyFromLong(lptr)
        ' add the vehicle to our vehicle collection
        m_colVehicles.Add oVeh, sKey
        Set oVeh = Nothing
        Set o2 = Nothing
        createVehicle = lptr
    End If
    Set m_ObjectStore = Nothing
End Function
Public Function destroyVehicle(ByVal h As Long) As Boolean
    On Error GoTo err
    Set m_oVehicles.Item(h) = Nothing
    destroyVehicle = True
    Exit Function
err:
    Debug.Print TypeName(Me) & ":destroyVehicle() -- Error:  Invalid vehicle handle."
End Function

'''''''''''''''''''''''''''''''''
' Node manipulation functions
'''''''''''''''''''''''''''''''''
'NOTE: Most of these functions will return a boolean
'  Upon receiving success, the client can call getNode() to refresh the list of nodes
'  Remember that on switch to internet messages, the client will only receive notification messages
'  like ADD_SUCESS, or ADD_FAIL.
Public Function addNode(sFilePath As String, ByVal hDestNode As Long) As Long
    Dim oPersist As cIPersist
    Dim oChild As cINode
    
    ' all nodes are guaranteed to have a unique handle so i dont know if i really need to use
    ' the hVehicle handle.
    ' This function by definition adds a new node/branch from saved file
    ' Inter-tree "adding" is actually moveNode()
    Set m_ObjectStore = New PersistenceManager.ObjectStore
    Set oPersist = m_ObjectStore.Deserialize(sFilePath, hDestNode, XML_NODE_OBJECT & "[@handle='0_']")
    If Not oPersist Is Nothing Then
        Set oChild = oPersist
        
        Dim oNode As cINode
        CopyMemory oNode, hDestNode, 4
        
        ' todo: parent needs to do a location check here.
        If oNode.addChild(oChild) Then ' if this returns false and we try to reference this object via its handle later on, we crash so check for TRUE return value
            oChild.Parent = oNode.Handle
            Debug.Assert oChild.Parent <> 0
            addNode = oChild.Handle
        End If
        CopyMemory oNode, 0&, 4
        Set oChild = Nothing
    Else
        Debug.Print TypeName(Me) & ":addNode() -- Error.  Could not add node to vehicle."
    End If
    Set m_ObjectStore = Nothing
            
    ' todo: remember that cINode will need to distinguish between logical parents and actual parents for
    ' purposes of checking child add permissions / priveledges
    ' maybe check how scene graphs manage testing for allowed children?
    ' Actually checking between logical parents (oGroup) should be easy since oGroup's dont implement
    ' cIContainer.
End Function
Public Function DeleteNode(ByVal hNode As Long) As Boolean
    Dim lngAttributes As Long
    Dim oNode As cINode
    Dim oParent As cINode
    Dim hParent As Long
    Dim index As Long
    Dim f As Boolean
    
    CopyMemory oNode, hNode, 4
    If Not oNode Is Nothing Then
        lngAttributes = oNode.Attributes
    
        f = lngAttributes And NODE_REQUIRED ' verify the user isnt trying to delete a node which must not be
        If Not f Then
            hParent = oNode.Parent
            CopyMemory oParent, hParent, 4
            If Not oParent Is Nothing Then
                index = oParent.getChildIndexByHandle(hNode)
                oParent.removeChild index
                DeleteNode = True
            End If
            CopyMemory oParent, 0&, 4
        End If
    End If
    CopyMemory oNode, 0&, 4
End Function
Public Function renameNode(ByVal hNode As Long, ByRef sName As String) As Boolean
    Dim oNode As cINode
    Dim f As Boolean
    Dim lngAttributes As Long
    
    CopyMemory oNode, hNode, 4
    If Not oNode Is Nothing Then
        lngAttributes = oNode.Attributes
        f = lngAttributes And NODE_RENAMEABLE
        If f Then
            oNode.Name = sName
            renameNode = True
        End If
    End If
    CopyMemory oNode, 0&, 4
End Function
Public Function moveNode(ByVal hSrcNode As Long, ByVal hDestNode As Long) As Boolean
    Dim oSrc As cINode
    Dim oDest As cINode
    Dim hParent As Long
    Dim oParent As cINode
    Dim f As Boolean
    Dim lngAttributes As Long
    
    If hSrcNode <> hDestNode Then ' cant move a node onto itself!
        CopyMemory oSrc, hSrcNode, 4
        If Not oSrc Is Nothing Then
            lngAttributes = oSrc.Attributes
            f = lngAttributes And NODE_FIXED  ' check for move priveledges
            If Not f Then
                CopyMemory oDest, hDestNode, 4
                If Not oDest Is Nothing Then
                    hParent = oDest.Parent
                    ' check for circular errors (i.e. you can't move a parent node onto its children or subchildren)
                    Do While hParent <> 0
                        CopyMemory oParent, hParent, 4
                        hParent = oParent.Parent
                        CopyMemory oParent, 0&, 4
                        If hParent = hSrcNode Then
                            fcircular = True
                            Exit Do
                        End If
                    Loop
                    If Not fcircular Then
                        ' todo: location checking must be done next
                        oDest.addChild oSrc
                        moveNode = True
                    End If
                End If
                CopyMemory oDest, 0&, 4
            End If
        End If
        CopyMemory oSrc, 0&, 4
    End If
End Function
Public Function copyNode(ByVal hNode As Long, ByRef sFileName As String, ByVal bEntireBranch As Boolean) As Boolean
    ' this just calls saveNode() except that we check attributes for NODE_COPYABLE first
    Dim oNode As cINode
    Dim lngAttributes As Long
    Dim f As Boolean
    Dim s As String
    
    s = sFileName
    CopyMemory oNode, hNode, 4
    lngAttributes = oNode.Attributes
    CopyMemory oNode, 0&, 4
    f = lngAttributes And NODE_COPYABLE
    If f Then
        saveNode hNode, s, bEntireBranch
        copyNode = True
    End If
End Function
Public Function pasteNode(ByRef sFileName As String, ByVal hDest As Long) As Boolean
    If addNode(sFileName, hDest) Then
        pasteNode = True
    End If
End Function
Public Function saveNode(ByVal hNode As Long, ByRef sFileName As String, ByVal bEntireBranch As Boolean) As Boolean
    Dim oNode As cINode
    Dim f As Boolean
    Dim xmlDoc As New MSXML2.DOMDocument40
    Dim nodeRoot As MSXML2.IXMLDOMNode
    Dim sXML As String
    Dim hFile As Long
    Dim sStandAlone As String
    
    Dim lngAttributes As Long
    
    CopyMemory oNode, hNode, 4
    lngAttributes = oNode.Attributes
    hFile = FreeFile()
    sStandAlone = "yes"
    
    ' check for save priveledges
    f = lngAttributes And NODE_SAVEABLE
    f = True  'todo: just to force save while debugging
    If f Then
        'todo: how would i control whether we recurse if bEntireBranch = TRUE?
        Set m_ObjectStore = New PersistenceManager.ObjectStore
        Set xmlDoc = m_ObjectStore.Serialize(oNode, True)
        Set nodeRoot = xmlDoc.documentElement
        
        Open sFileName For Output As #hFile
        sXML = "<?xml version=" & """" & "1.0" & """" & " standalone=" & """" & sStandAlone & """" & "?>"
        Print #hFile, sXML
        
        sXML = FormatNodeAsString(nodeRoot)
        Print #hFile, sXML
        Close #hFile
        Shell "notepad " & sFileName, vbNormalFocus
        saveNode = True
    End If
    
    CopyMemory oNode, 0&, 4
End Function
Public Function getNode()
End Function

''''''''''''''''''''''''''''''''''''''
' This is an EVENT() generated by the clsObjectStore
' When it needs to create an object, it makes a call here...
' The only thing this event needs is an instanced object from vehicle.factory to be returned
Private Sub m_ObjectStore_RequestObject(ByVal Classname As String, ByVal sDefPath As String, ByVal sGUID As String, newObject As PersistenceManager.cIPersist)
    Dim oNode As cINode
    
    On Error GoTo err
    Set oNode = m_oFactory.CreateComponent(Classname, sDefPath, sGUID)
    Set newObject = oNode
    
    Set oNode = Nothing
    Exit Sub
err:
    Debug.Print "frmDesigner.m_ObjectStore_RequestObject() -- Error # " & err.Number & " " & err.Description
    If err.Number = 13 Then
        Debug.Print "frmDesigner.m_ObjectStore_RequestObject() -- The class '" & Classname & "' probably doesnt implement IPersist...."
    End If
    Set oNode = Nothing
End Sub

