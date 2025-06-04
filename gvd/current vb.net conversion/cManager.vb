Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cManager_NET.cManager")> Public Class cManager
	Private WithEvents m_ObjectStore As PersistenceManager.ObjectStore
	Private m_colVehicles As Collection '10/28/02 MPJ  -- We are now supporting multiple vehicles
	Private m_oFactory As Vehicles.cFactory ' only one factory is needed however...
	
	Private m_oVehicles As Collection
	
	'  note: we are not going to fuck around with internet shit right now.  Instead
	'  we are going to move all our client accessible commands here.  The user may
	'  be able to instance local versions of cINode and cPropertyItem, but thats about it.
	'  They will not have access to cIComponent or other such interfaces as they are not
	'  allowed to access components directly.
	
	'  Eventually we will incorp localhost loopback, but that will be after gvd 2.0 is done.
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Dim Vehicles As Object
		m_oFactory = New Vehicles.cFactory
		m_colVehicles = New Collection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object m_oFactory may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_oFactory = Nothing
		' Destroy the collection and all objects in it
		Dim oVehicle As cVehicle
		For	Each oVehicle In m_colVehicles
			'UPGRADE_NOTE: Object oVehicle may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oVehicle = Nothing
		Next oVehicle
		'UPGRADE_NOTE: Object m_colVehicles may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_colVehicles = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Function createVehicle(ByRef sFilePath As String) As Integer
		Dim PersistenceManager As Object
		Dim Vehicles As Object
		Dim lptr As Integer
		Dim sKey As String
		Dim oVeh As Vehicles.cINode
		Dim o2 As Object
		
		m_ObjectStore = New PersistenceManager.ObjectStore
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_ObjectStore.Deserialize. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		o2 = m_ObjectStore.Deserialize(sFilePath, 0, XML_NODE_OBJECT & "[@handle='0_']")
		If Not o2 Is Nothing Then
			oVeh = o2
			'UPGRADE_WARNING: Couldn't resolve default property of object oVeh.Handle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lptr = oVeh.Handle
			
			sKey = CStr(lptr) & "_" ' todo: make this function from gvd project available here too KeyFromLong(lptr)
			' add the vehicle to our vehicle collection
			m_colVehicles.Add(oVeh, sKey)
			'UPGRADE_NOTE: Object oVeh may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oVeh = Nothing
			'UPGRADE_NOTE: Object o2 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			o2 = Nothing
			createVehicle = lptr
		End If
		'UPGRADE_NOTE: Object m_ObjectStore may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_ObjectStore = Nothing
	End Function
	Public Function destroyVehicle(ByVal h As Integer) As Boolean
		On Error GoTo err_Renamed
		'UPGRADE_NOTE: Object m_oVehicles.Item() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_oVehicles.Item(h) = Nothing
		destroyVehicle = True
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print(TypeName(Me) & ":destroyVehicle() -- Error:  Invalid vehicle handle.")
	End Function
	
	'''''''''''''''''''''''''''''''''
	' Node manipulation functions
	'''''''''''''''''''''''''''''''''
	'NOTE: Most of these functions will return a boolean
	'  Upon receiving success, the client can call getNode() to refresh the list of nodes
	'  Remember that on switch to internet messages, the client will only receive notification messages
	'  like ADD_SUCESS, or ADD_FAIL.
	Public Function addNode(ByRef sFilePath As String, ByVal hDestNode As Integer) As Integer
		Dim PersistenceManager As Object
		'UPGRADE_ISSUE: cIPersist object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oPersist As _cIPersist
		Dim oChild As _cINode
		
		' all nodes are guaranteed to have a unique handle so i dont know if i really need to use
		' the hVehicle handle.
		' This function by definition adds a new node/branch from saved file
		' Inter-tree "adding" is actually moveNode()
		m_ObjectStore = New PersistenceManager.ObjectStore
		'UPGRADE_WARNING: Couldn't resolve default property of object m_ObjectStore.Deserialize. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oPersist = m_ObjectStore.Deserialize(sFilePath, hDestNode, XML_NODE_OBJECT & "[@handle='0_']")
		Dim oNode As _cINode
		If Not oPersist Is Nothing Then
			oChild = oPersist
			
			'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oNode, hDestNode, 4)
			
			' todo: parent needs to do a location check here.
			If oNode.addChild(oChild) Then ' if this returns false and we try to reference this object via its handle later on, we crash so check for TRUE return value
				oChild.Parent = oNode.Handle
				System.Diagnostics.Debug.Assert(oChild.Parent <> 0, "")
				addNode = oChild.Handle
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oNode, 0, 4)
			'UPGRADE_NOTE: Object oChild may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oChild = Nothing
		Else
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			Debug.Print(TypeName(Me) & ":addNode() -- Error.  Could not add node to vehicle.")
		End If
		'UPGRADE_NOTE: Object m_ObjectStore may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_ObjectStore = Nothing
		
		' todo: remember that cINode will need to distinguish between logical parents and actual parents for
		' purposes of checking child add permissions / priveledges
		' maybe check how scene graphs manage testing for allowed children?
		' Actually checking between logical parents (oGroup) should be easy since oGroup's dont implement
		' cIContainer.
	End Function
	Public Function DeleteNode(ByVal hNode As Integer) As Boolean
		Dim lngAttributes As Integer
		Dim oNode As _cINode
		Dim oParent As _cINode
		Dim hParent As Integer
		Dim index As Integer
		Dim f As Boolean
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, hNode, 4)
		If Not oNode Is Nothing Then
			lngAttributes = oNode.Attributes
			
			f = lngAttributes And NODE_REQUIRED ' verify the user isnt trying to delete a node which must not be
			If Not f Then
				hParent = oNode.Parent
				'UPGRADE_WARNING: Couldn't resolve default property of object oParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CopyMemory(oParent, hParent, 4)
				If Not oParent Is Nothing Then
					index = oParent.getChildIndexByHandle(hNode)
					oParent.removeChild(index)
					DeleteNode = True
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object oParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CopyMemory(oParent, 0, 4)
			End If
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, 0, 4)
	End Function
	Public Function renameNode(ByVal hNode As Integer, ByRef sName As String) As Boolean
		Dim oNode As _cINode
		Dim f As Boolean
		Dim lngAttributes As Integer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, hNode, 4)
		If Not oNode Is Nothing Then
			lngAttributes = oNode.Attributes
			f = lngAttributes And NODE_RENAMEABLE
			If f Then
				oNode.Name = sName
				renameNode = True
			End If
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, 0, 4)
	End Function
	Public Function moveNode(ByVal hSrcNode As Integer, ByVal hDestNode As Integer) As Boolean
		Dim fcircular As Object
		Dim oSrc As _cINode
		Dim oDest As _cINode
		Dim hParent As Integer
		Dim oParent As _cINode
		Dim f As Boolean
		Dim lngAttributes As Integer
		
		If hSrcNode <> hDestNode Then ' cant move a node onto itself!
			'UPGRADE_WARNING: Couldn't resolve default property of object oSrc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oSrc, hSrcNode, 4)
			If Not oSrc Is Nothing Then
				lngAttributes = oSrc.Attributes
				f = lngAttributes And NODE_FIXED ' check for move priveledges
				If Not f Then
					'UPGRADE_WARNING: Couldn't resolve default property of object oDest. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CopyMemory(oDest, hDestNode, 4)
					If Not oDest Is Nothing Then
						hParent = oDest.Parent
						' check for circular errors (i.e. you can't move a parent node onto its children or subchildren)
						Do While hParent <> 0
							'UPGRADE_WARNING: Couldn't resolve default property of object oParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							CopyMemory(oParent, hParent, 4)
							hParent = oParent.Parent
							'UPGRADE_WARNING: Couldn't resolve default property of object oParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							CopyMemory(oParent, 0, 4)
							If hParent = hSrcNode Then
								'UPGRADE_WARNING: Couldn't resolve default property of object fcircular. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								fcircular = True
								Exit Do
							End If
						Loop 
						If Not fcircular Then
							' todo: location checking must be done next
							oDest.addChild(oSrc)
							moveNode = True
						End If
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object oDest. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CopyMemory(oDest, 0, 4)
				End If
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object oSrc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oSrc, 0, 4)
		End If
	End Function
	Public Function copyNode(ByVal hNode As Integer, ByRef sFileName As String, ByVal bEntireBranch As Boolean) As Boolean
		' this just calls saveNode() except that we check attributes for NODE_COPYABLE first
		Dim oNode As _cINode
		Dim lngAttributes As Integer
		Dim f As Boolean
		Dim s As String
		
		s = sFileName
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, hNode, 4)
		lngAttributes = oNode.Attributes
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, 0, 4)
		f = lngAttributes And NODE_COPYABLE
		If f Then
			saveNode(hNode, s, bEntireBranch)
			copyNode = True
		End If
	End Function
	Public Function pasteNode(ByRef sFileName As String, ByVal hDest As Integer) As Boolean
		If addNode(sFileName, hDest) Then
			pasteNode = True
		End If
	End Function
	Public Function saveNode(ByVal hNode As Integer, ByRef sFileName As String, ByVal bEntireBranch As Boolean) As Boolean
		Dim FormatNodeAsString As Object
		Dim PersistenceManager As Object
		Dim MSXML2 As Object
		Dim oNode As _cINode
		Dim f As Boolean
		'UPGRADE_ISSUE: MSXML2.DOMDocument40 object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim xmlDoc As MSXML2.DOMDocument40 = New MSXML2.DOMDocument40
		'UPGRADE_ISSUE: MSXML2.IXMLDOMNode object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim nodeRoot As MSXML2.IXMLDOMNode
		Dim sXML As String
		Dim hFile As Integer
		Dim sStandAlone As String
		
		Dim lngAttributes As Integer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, hNode, 4)
		lngAttributes = oNode.Attributes
		hFile = FreeFile
		sStandAlone = "yes"
		
		' check for save priveledges
		f = lngAttributes And NODE_SAVEABLE
		f = True 'todo: just to force save while debugging
		If f Then
			'todo: how would i control whether we recurse if bEntireBranch = TRUE?
			m_ObjectStore = New PersistenceManager.ObjectStore
			'UPGRADE_WARNING: Couldn't resolve default property of object m_ObjectStore.Serialize. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xmlDoc = m_ObjectStore.Serialize(oNode, True)
			'UPGRADE_WARNING: Couldn't resolve default property of object xmlDoc.documentElement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nodeRoot = xmlDoc.documentElement
			
			FileOpen(hFile, sFileName, OpenMode.Output)
			sXML = "<?xml version=" & """" & "1.0" & """" & " standalone=" & """" & sStandAlone & """" & "?>"
			PrintLine(hFile, sXML)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object FormatNodeAsString(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sXML = FormatNodeAsString(nodeRoot)
			PrintLine(hFile, sXML)
			FileClose(hFile)
			Shell("notepad " & sFileName, AppWinStyle.NormalFocus)
			saveNode = True
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, 0, 4)
	End Function
	Public Function getNode() As Object
	End Function
	
	''''''''''''''''''''''''''''''''''''''
	' This is an EVENT() generated by the clsObjectStore
	' When it needs to create an object, it makes a call here...
	' The only thing this event needs is an instanced object from vehicle.factory to be returned
	Private Sub m_ObjectStore_RequestObject(ByVal Classname As String, ByVal sDefPath As String, ByVal sGUID As String, ByRef newObject As PersistenceManager.cIPersist)
		Dim PersistenceManager As Object
		Dim oNode As _cINode
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oFactory.CreateComponent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oNode = m_oFactory.CreateComponent(Classname, sDefPath, sGUID)
		newObject = oNode
		
		'UPGRADE_NOTE: Object oNode may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oNode = Nothing
		Exit Sub
err_Renamed: 
		Debug.Print("frmDesigner.m_ObjectStore_RequestObject() -- Error # " & Err.Number & " " & Err.Description)
		If Err.Number = 13 Then
			Debug.Print("frmDesigner.m_ObjectStore_RequestObject() -- The class '" & Classname & "' probably doesnt implement IPersist....")
		End If
		'UPGRADE_NOTE: Object oNode may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oNode = Nothing
	End Sub
End Class