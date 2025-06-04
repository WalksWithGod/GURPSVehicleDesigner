Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cFeature_NET.cFeature")> Public Class cFeature
	Implements _cIPersist
	Implements _cINode
	Implements _cIDisplay
	Implements _cIBuild
	
	'todo: shoudl make these features appear in the component drop down list and that they can only be added to a parent
	' node of cSurface
	
	
	' -------cIBuild related variables
	Private Structure GVD_DATA_TABLE
		Dim ID As Integer
		Dim ptrTable As Integer
	End Structure
	Private Structure GVD_OPTIONS
		Dim index As Integer
		Dim selectionCount As Integer
		Dim ptrTable As Integer
	End Structure
	Private Structure GVD_USER_INPUT
		Dim sngValue As Single
		Dim sngURange As Single
		Dim sngLRange As Single
	End Structure
	Private Structure GVD_FORMULA
		Dim lngStatID As Integer
		Dim lngFormulaID As Integer
	End Structure
	Private m_Tables() As GVD_DATA_TABLE
	Private m_Options() As GVD_OPTIONS
	Private m_UserInput() As GVD_USER_INPUT
	Private m_Formulas() As GVD_FORMULA
	Private m_lngTableCount As Integer
	Private m_lngOptionCount As Integer
	Private m_lngUserInputCount As Integer
	Private m_lngFormulaCount As Integer
	
	' -------cINode interface variables
	Private m_lngMaxChildren As Integer
	Private m_lngChildCount As Integer
	Private m_oChildren() As cINode
	Private m_lngAttributes As Integer
	Private m_hParent As Integer
	Private m_hMe As Integer
	Private m_sName As String
	Private m_sDescription As String
	Private m_sImage As String
	
	
	' -------cIDisplay interface variables
	Private m_lngPropCount As Integer
	Private m_lngCurrentPropItem As Integer
	'UPGRADE_ISSUE: cPropertyItem object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Private m_oProperties() As cPropertyItem
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'todo: should i hardcode limits on m_lngMaxChildren so that def file will never accidentally override?
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		Dim i As Integer
		For i = 0 To m_lngChildCount - 1
			'UPGRADE_NOTE: Object m_oChildren() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			m_oChildren(i) = Nothing
		Next 
		
		For i = 0 To m_lngPropCount - 1
			'UPGRADE_NOTE: Object m_oProperties() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			m_oProperties(i) = Nothing
		Next 
		For i = 0 To m_lngTableCount - 1
			destroyTable(m_Tables(i).ptrTable)
		Next 
		For i = 0 To m_lngOptionCount - 1
			destroyTable(m_Options(i).ptrTable)
		Next 
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	'///////////////////////////////////////////////////
	'//cIBuild Implemented Properties and Functions
	Private Function cIBuild_getOption(ByVal lngIndex As Integer) As Integer Implements _cIBuild.getOption
		On Error GoTo err_Renamed
		cIBuild_getOption = m_Options(lngIndex).index
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print(TypeName(Me) & ":cIBuild_getOption -- ERROR #" & Err.Number & " " & Err.Description)
	End Function
	Private Function cIBuild_setOption(ByVal lngIndex As Integer, ByVal lngSelection As Integer) As Boolean Implements _cIBuild.setOption
		On Error GoTo err_Renamed
		' before assigning the value, check that the selection is valid by determining if its in the range of 0 to (SelectCount -1)
		If (lngSelection <= m_Options(lngIndex).selectionCount - 1) Then
			m_Options(lngIndex).index = lngSelection
			cIBuild_setOption = True
		Else
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			modHelper.InfoPrint(1, TypeName(Me) & ":cIBuild_setOption() -- ERROR.  Selection invalid.  Are you a hacker?")
		End If
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print(TypeName(Me) & ":cIBuild_setOption -- ERROR #" & Err.Number & " " & Err.Description)
		cIBuild_setOption = False
	End Function
	Private Function cIBuild_getUserInput(ByVal lngIndex As Integer) As Single Implements _cIBuild.getUserInput
		On Error GoTo err_Renamed
		cIBuild_getUserInput = m_UserInput(lngIndex).sngValue
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print(TypeName(Me) & ":cIBuild_getUserInput -- ERROR #" & Err.Number & " " & Err.Description)
	End Function
	Private Function cIBuild_setUserInput(ByVal lngIndex As Integer, ByVal sngValue As Single) As Boolean Implements _cIBuild.setUserInput
		On Error GoTo err_Renamed
		If (sngValue >= m_UserInput(lngIndex).sngLRange) And (sngValue <= m_UserInput(lngIndex).sngURange) Then
			m_UserInput(lngIndex).sngValue = sngValue
			cIBuild_setUserInput = True
		Else
			modHelper.InfoPrint(1, "User input for this field limited to values between " & m_UserInput(lngIndex).sngLRange & " and " & m_UserInput(lngIndex).sngURange)
			cIBuild_setUserInput = False
		End If
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print(TypeName(Me) & ":cIBuild_setUserInput -- ERROR #" & Err.Number & " " & Err.Description)
		cIBuild_setUserInput = False
	End Function
	Private Function cIBuild_calcStats(ByRef oVisitor As cStats) As Boolean
		'UPGRADE_WARNING: Couldn't resolve default property of object Me. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oVisitor.calcStats(Me)
		' todo: actually before we can pass m_lngTL to the stat table, we need to know the bounds
		'm_dblCost = m_sngTable(0, 2)
		'm_dblWeight = m_sngTable(0, 0)
		'm_dblVolume = m_sngTable(0, 1) '+ AddedVolume
		'm_dblSurfaceArea = CalcSurfaceArea(m_dblVolume)
		'm_dblHitpoints = CalcComponentHitpoints(m_dblSurfaceArea)
	End Function
	
	
	'///////////////////////////////////////////////////
	'//cIDisplay Implemented Properties and Functions
	Private Function cIDisplay_getFirstPropertyItem() As cPropertyItem Implements _cIDisplay.getFirstPropertyItem
		On Error GoTo err_Renamed
		If Not m_oProperties(0) Is Nothing Then
			cIDisplay_getFirstPropertyItem = m_oProperties(0)
			m_lngCurrentPropItem = 0
		End If
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print(TypeName(Me) & ":cIDisplay_getFirstPropertyItem() -- no properties in m_oProperties() array for " & TypeName(Me))
	End Function
	
	Private Function cIDisplay_getNextPropertyItem() As cPropertyItem Implements _cIDisplay.getNextPropertyItem
		m_lngCurrentPropItem = m_lngCurrentPropItem + 1
		If m_lngCurrentPropItem <= m_lngPropCount - 1 Then
			If Not m_oProperties(m_lngCurrentPropItem) Is Nothing Then
				cIDisplay_getNextPropertyItem = m_oProperties(m_lngCurrentPropItem)
			End If
		Else
			m_lngCurrentPropItem = m_lngCurrentPropItem - 1
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			Debug.Print(TypeName(Me) & ":cIDisplay:getNextPropertyItem -- nextItem exceeds Property Count.")
		End If
	End Function
	Private Function cIDisplay_getPropertyItemByIndex(ByVal iIndex As Integer) As cPropertyItem Implements _cIDisplay.getPropertyItemByIndex
		On Error Resume Next
		cIDisplay_getPropertyItemByIndex = m_oProperties(iIndex)
	End Function
	'/////////////////////////////////////////////
	'//Implemented cINode Properties and Functions
	Private Function cINode_AddChild(ByRef oBase As _cINode) As Boolean Implements _cINode.AddChild
		'TODO: will leaf options ever accept children?  Probably not... just remember
		' to investigate and then delete this code if its not ncessary, same for the getchild stuff
		If m_lngMaxChildren = m_lngChildCount Then
			cINode_AddChild = False
		Else
			m_lngChildCount = m_lngChildCount + 1
			ReDim Preserve m_oChildren(m_lngChildCount - 1)
			m_oChildren(m_lngChildCount - 1) = oBase
			cINode_AddChild = True
		End If
	End Function
	Private Function cINode_getChildrenByClassName(ByRef Classname As String, ByRef hChilds() As Integer) As Boolean Implements _cINode.getChildrenByClassName
	End Function
	Private Function cINode_getChildIndexByHandle(ByVal h As Integer) As Integer Implements _cINode.getChildIndexByHandle
		Dim i As Integer
		Dim lRet As Integer
		lRet = -1
		For i = 0 To m_lngChildCount - 1
			If m_oChildren(i).Handle = h Then lRet = i : Exit For
		Next 
		cINode_getChildIndexByHandle = lRet
	End Function
	Private Function cINode_getChild(ByVal lngIndex As Integer) As _cINode Implements _cINode.getChild
		If (lngIndex >= 0) And (m_lngChildCount > 0) And (lngIndex <= m_lngChildCount - 1) Then
			cINode_getChild = m_oChildren(lngIndex)
		End If
	End Function
	Private Function cINode_removeChild(ByVal lngIndex As Integer) As Boolean Implements _cINode.removeChild
		Dim i As Integer
		If (lngIndex <= m_lngChildCount - 1) And (lngIndex >= 0) Then
			'UPGRADE_NOTE: Object m_oChildren() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			m_oChildren(lngIndex) = Nothing
			For i = lngIndex + 1 To m_lngChildCount - 1
				m_oChildren(i - 1) = m_oChildren(i)
			Next 
		End If
		m_lngChildCount = m_lngChildCount - 1
		If m_lngChildCount > 0 Then
			ReDim Preserve m_oChildren(m_lngChildCount - 1)
		Else
			Erase m_oChildren
		End If
	End Function
	Private ReadOnly Property cINode_childCount() As Integer Implements _cINode.childCount
		Get
			cINode_childCount = m_lngChildCount
		End Get
	End Property
	Private ReadOnly Property cINode_ClassName() As String Implements _cINode.ClassName
		Get
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			cINode_ClassName = TypeName(Me)
		End Get
	End Property
	Private ReadOnly Property cINode_Attributes() As Integer Implements _cINode.Attributes
		Get
			cINode_Attributes = m_lngAttributes
		End Get
	End Property
	Private Property cINode_Handle() As Integer Implements _cINode.Handle
		Get
			cINode_Handle = m_hMe
		End Get
		Set(ByVal Value As Integer)
			m_hMe = Value
		End Set
	End Property
	Private Property cINode_Parent() As Integer Implements _cINode.Parent
		Get
			cINode_Parent = m_hParent
		End Get
		Set(ByVal Value As Integer)
			m_hParent = Value
		End Set
	End Property
	Private Property cINode_Name() As String Implements _cINode.Name
		Get
			cINode_Name = m_sName
		End Get
		Set(ByVal Value As String)
			m_sName = Value
		End Set
	End Property
	Private Property cINode_Description() As String Implements _cINode.Description
		Get
			cINode_Description = m_sDescription
		End Get
		Set(ByVal Value As String)
			m_sDescription = Value
		End Set
	End Property
	Private Property cINode_Image() As String Implements _cINode.Image
		Get
			cINode_Image = m_sImage
		End Get
		Set(ByVal Value As String)
			m_sImage = Value
		End Set
	End Property
	
	'//cIPersist Interface
	Private ReadOnly Property cIPersist_Classname() As String
		Get
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			cIPersist_Classname = TypeName(Me)
		End Get
	End Property
	Private ReadOnly Property cIPersist_GUID() As String
		Get
		End Get
	End Property
	Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Integer)
		Dim i As Integer
		
		If iMode = modGVDXMLSchemaConstants.GVD_XML_TYPE.cmp Then
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sName = op.Load(XML_NODE_NAME)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sDescription = op.Load(XML_NODE_DESCRIPTION)
			
			
			' now load the saved option indexes and userinput values from the .cmp file
			If m_lngOptionCount > 0 Then
				ReDim m_Options(m_lngOptionCount - 1)
				For i = 0 To m_lngOptionCount - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_Options(i).index = op.Load("option_value" & i)
				Next 
			End If
			
			If m_lngUserInputCount > 0 Then
				ReDim m_UserInput(m_lngUserInputCount - 1)
				'load in the min/max allowed ranges
				For i = 0 To m_lngUserInputCount - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_UserInput(i).sngValue = op.Load("userinput_value" & i)
				Next 
			End If
			
			' todo: call to update stats go here?
			' If so, then no need to ever save/load the statistics since we recalc them.
			' We need to make a decision here to either set a bChanged flag so that we can update later
			' or we update the stats NOW, even though we are still loading  potentially, other vehicle nodes.
			' Since this is a feature, calc'ing it now could result in wrong surface area being used.  If we use
			' flags, we can specify bChanged and bPriority to indicate when its ok to calc stats for this node.
			' we can also flag if bIndependant if these stats have zero requirements of other stats. (e.g. changing
			' of body surface area still wont effect the weight or volume of a radio component)
			
			' come to think of it, we have to use flags.  Just think about when adding a Turret, we can't calc stats
			' right after loading cuz there are children that havent been added yet which will ultimately effect its
			' stats.
			
			' Recall that some components like wheels, you cant load its stats anyway since they are in fact
			' dependant on the body's stats.  So this is yet another reason why we should flag here and re-calc
			' rather than just load the stats and not make any calls to update or flag for update.
			
			' So assuming we are going to set a flag here, what is the process?  Perhaps call to setFlag() internally
			' makes call to notifyParent() which recurses up to the root vehicle.  Once the root vehicle is reached, the
			' recursion can begin starting with the leaf nodes.  But what about sibling nodes.. wont that have some funky
			' tendancy to recurse all the way to the vehicle node more than once?  I could end the recursion prematurely
			' if i see the flag of a parent node is already set, therefore assuming the parent is already notified all the
			' way to the top.
			
			' And note that performances need to be updated virtually every change even though no child node will ever
			' be able to notify it (same with features)  Thus these need to be done seperately.  Wheel too though
			' it is on the components branch, but its dependant on body's stat.
			
			' Dont forget about armor
			' Dont forget about airbags and stabilization gear who's stats depend on the component they compliment
			
		Else 'DEF
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngAttributes = op.Load("attributes")
			
#If DEBUG_MODE Then
			'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
			Debug.Assert m_lngMaxChildren <= 0
#End If
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sImage = op.Load(XML_NODE_IMAGE)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngTableCount = op.Load(XML_NODE_STATS_TABLECOUNT)
			If m_lngTableCount > 0 Then
				ReDim m_Tables(m_lngTableCount - 1)
				For i = 0 To m_lngTableCount - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_Tables(i).ptrTable = op.Load(XML_NODE_TABLE & i)
				Next 
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngOptionCount = op.Load(XML_NODE_OPTION_MODIFER_TABLE_COUNT)
			If m_lngOptionCount > 0 Then
				ReDim m_Options(m_lngOptionCount - 1)
				For i = 0 To m_lngOptionCount - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_Options(i).selectionCount = op.Load("option_selectioncount" & i)
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_Options(i).ptrTable = op.Load("option_table" & i)
				Next 
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngUserInputCount = op.Load("userinput_count")
			If m_lngUserInputCount > 0 Then
				ReDim m_UserInput(m_lngUserInputCount - 1)
				'load in the min/max allowed ranges
				For i = 0 To m_lngUserInputCount - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_UserInput(i).sngURange = op.Load("userinput_urange" & i)
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_UserInput(i).sngLRange = op.Load("userinput_lrange" & i)
				Next 
			End If
			
			' load properties last, these will reference variables initialized above
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngPropCount = op.Load(XML_NODE_PROPERTYCOUNT)
			If m_lngPropCount > 0 Then
				ReDim m_oProperties(m_lngPropCount - 1)
				For i = 0 To m_lngPropCount - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_oProperties(i) = op.Load(XML_NODE_PROPERTY & i)
				Next 
			End If
		End If
	End Sub
	Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Store. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		op.Store("name", m_sName)
		
	End Sub
End Class