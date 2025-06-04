Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("aTest_NET.aTest")> Public Class aTest
	Implements _cIPersist
	Implements _cINode
	Implements _cIDisplay
	Implements _cIComponent
	Implements _cIBuild
	' NOTE: This is our reference test class.  From it we will increase functionality until we've covered all the
	' basic interfaces and component types.  We will restrict most todo and ntoe comments to here for now
	' as we hash out the design
	
	
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
	' todo: since this wont implement cIContainer, it actually doesnt need child count or the child array
	' after i get this thing to run and load into the tree, i will delete these since aTest is my base
	' class for which to model all other cIComponents (non Container)
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
	
	' ------- cIComponent interface variables
	Private m_sngTL As Single '<-- this is the only "stat" the user can modify directly
	'      Note: that we cant actually "test" for an overflow without actually causing an overflow.
	'      That is to say, our "test" might overflow too.  This puts us back to the max user input values
	'      As well as max number of components in a vehicle.
	'      Quantity must be taken into account too.  Still havent finalized how tohandlle that.
	Private m_dblCost As Double
	Private m_dblWeight As Double
	Private m_dblVolume As Double
	Private m_dblSurfaceArea As Double
	Private m_dblHitpoints As Double
	
	' ------- cIArmor interface variables todo: note i dont think this will be a seperate interface but rather
	'         apart of cIcomponent.
	Private m_lngDR As Integer 'todo: this dr is probably obsolete since it will come directly from armor?  hrm, but
	'whta about default DR for components that dont actually set component armor?
	Private m_oArmor As cArmor
	
	
	
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
	'//cIDisplay Implemented Properties and Functions
	Private Function cIDisplay_getFirstPropertyItem() As cPropertyItem Implements _cIDisplay.getFirstPropertyItem
		If Not m_oProperties(0) Is Nothing Then
			cIDisplay_getFirstPropertyItem = m_oProperties(0)
			m_lngCurrentPropItem = 0
		End If
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
	
	'///////////////////////////////////////////////////
	'//cINode Implemented Properties and Functions
	Private Function cINode_AddChild(ByRef oBase As _cINode) As Boolean Implements _cINode.AddChild
		'todo: this is a leaf node, probably dont need any implementation just return false
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
		'todo: dont need implemenation, this is leaf right?
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
	
	'///////////////////////////////////////////////////
	'//cIComponent Implemented Properties and Functions
	Private Property cIComponent_LogicalParent() As Integer Implements _cIComponent.LogicalParent
		Get
		End Get
		Set(ByVal Value As Integer)
		End Set
	End Property
	Private Property cIComponent_SurfaceArea() As Double Implements _cIComponent.SurfaceArea
		Get
			cIComponent_SurfaceArea = m_dblSurfaceArea
		End Get
		Set(ByVal Value As Double)
			m_dblSurfaceArea = Value
		End Set
	End Property
	Private Property cIComponent_Cost() As Double Implements _cIComponent.Cost
		Get
		End Get
		Set(ByVal Value As Double)
			'todo: delete all the lets since these are calculated internally?
			' err... but not for custom components... perhaps a seperate interface for those right?
			' these should all be referenced as "Added_XXXXX" e.g. Added_Cost
			' seems like the best way to do it since we will be using function calls
			' to calcStats ONLY when a stat specific variable is altered.
			
			' Also, when modifying something like "Description" or "Notes" we dont want to call that
			' function, but just update the print string.
			
		End Set
	End Property
	Private Property cIComponent_Volume() As Double Implements _cIComponent.Volume
		Get
		End Get
		Set(ByVal Value As Double)
		End Set
	End Property
	Private Property cIComponent_Weight() As Double Implements _cIComponent.Weight
		Get
		End Get
		Set(ByVal Value As Double)
		End Set
	End Property
	Private Property cIComponent_HitPoints() As Double Implements _cIComponent.HitPoints
		Get
			cIComponent_HitPoints = m_dblHitpoints
		End Get
		Set(ByVal Value As Double)
			m_dblHitpoints = Value
		End Set
	End Property
	
	' TL is mostly just a keystone build option modifier.
	' So cIComponent is basically the key "build" interface. I think the only
	' reason i dont include it as m_options() local to each class is that
	' i need to set a default TL when a component is added to the tree.
	' NOTE: I probably should have TL in cIBuild_TL AND cIComponent.  They'll both
	' access the same internal variable, but this way we can access it from either interface
	Private Property cIComponent_TL() As Single Implements _cIComponent.TL
		Get
			cIComponent_TL = m_sngTL
		End Get
		Set(ByVal Value As Single)
			m_sngTL = Value
		End Set
	End Property
	
	
	'///////////////////////////////////////////////////
	'//cIPersist Implemented Properties and Functions
	Private ReadOnly Property cIPersist_Classname() As String
		Get
			'todo: whats this property for?
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			cIPersist_Classname = TypeName(Me)
		End Get
	End Property
	Private ReadOnly Property cIPersist_GUID() As String
		Get
			
		End Get
	End Property
	
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
	
	Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Integer)
		Dim i As Integer
		
		If iMode = modGVDXMLSchemaConstants.GVD_XML_TYPE.cmp Then
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sName = op.Load(XML_NODE_NAME)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sDescription = op.Load(XML_NODE_DESCRIPTION)
			'note: default values for options() and userinput() are always stored in the .cmp file and
			'      not the .def file
			'todo: load our user input saved values
			'todo: load our option saved values
			'todo: we need to gracefully resume if an option or userinput index
			'      is not represented... or do we require they all be in the .cmp even with 0 value?
			
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sImage = op.Load(XML_NODE_IMAGE)
			
			' todo: testing of loading formulas, stats, and multipliers
			'm_lngFormulaCount = op.Load("formula_count")
			'ReDim m_Formulas(m_lngFormulaCount - 1)
			'For i = 0 to m_lngFormulaCount -1
			'   m_Formulas(i) = op.Load(XML_NODE_FORMULA)
			'Next
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngTableCount = op.Load(XML_NODE_STATS_TABLECOUNT)
			ReDim m_Tables(m_lngTableCount - 1)
			For i = 0 To m_lngTableCount - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_Tables(i).ptrTable = op.Load(XML_NODE_TABLE & i)
			Next 
			
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
		'todo: we need to abort if any of these loads fail right?  Like imagine if the
		'      stats tables dont load, thats disaster pretty much so we have to abort... means the def
		'      was corrupt.
		
		'      I could change this to a function and return TRUE if we make it through with no errors.
		
	End Sub
	Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
	End Sub
End Class