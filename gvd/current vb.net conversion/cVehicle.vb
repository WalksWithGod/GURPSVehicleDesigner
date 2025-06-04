Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cVehicle_NET.cVehicle")> Public Class cVehicle
	Implements _cIDisplay
	Implements _cINode
	Implements _cIPersist
	
	
	' -------cINode interface variables
	Private m_lngMaxChildren As Integer
	Private m_lngChildCount As Integer
	Private m_oChildren() As cINode ' stats,crew, surface,hull and components classes need to go here.  all implement cInode and are thus autographed.  Only "components" implements cIContainer and can thus contain other 'components'.
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
	
	' ------- cIomponent interface properties
	Private m_Table() As Single
	Private m_lngTL As Integer
	Private m_dblHitpoints As Double
	Private m_dblSurfaceArea As Double
	Private m_dblCost As Double
	Private m_dblVolume As Double
	Private m_dblWeight As Double
	Public m_oArmor As cArmorLayer ' if it includes armor here,it should implement cIComponent since that includes armor as a public composite object
	'todo: make sure this is defined in the cIComponent interface
	
	' ------- cVehicle required variables
	
	' note that these are not declared as cINode since they do not have a seperte node in the tree.  instead they are simply
	' referenced through the cVehicle's own node and display interfaces
	Public Author As cAuthor
	'UPGRADE_ISSUE: cVersion object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Public Version As cVersion
	Public Description As cDescription
	
	
	' Since the following will display in the treex, they must all implement cINode and cIDisplay
	' Since they all must be persistable, they will implement cIPersist
	' However, none of them need to implement cICompoent or cIContainer.  Perhaps they should also
	' all be stored in the children() array.  this provides auto graphing.  We'll only need to do a quick
	' search of first level children in cVehicles to find any particular one.
	
	'Public Stats As cINode 'cStats
	'Public Crew As cINode 'cCrew
	'Public Surface As cINode 'cSurface
	'Public Hull As cINode ' cHull
	
	
	'Public PowerProfiles As cINode
	'Public FuelProfiles As cINode
	'Public WeaponProfiles As cINode
	'Public PerformanceProfiles As cINode
	'Public BatteryProfiles As cINode
	'Public Profiles As clsProfile
	
	'//
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		' Public Veh as cVehicle is declared in modHelper as global.  Thats definetly bad.
		Veh = Me ' public reference to this class 'todo: This is bad
		' We do need a way for objects within a vehicle to refer to the root
		' BUT, if we have arrays of vehicles, they will each overwrite Veh with a reference to themselves.
		' We need to manage all vehicle references thru cManager and use a collection.
		
		'-----------------all below obsolete?
		
		' holds collection of link objects
		'Set WeaponProfiles = New clsProfile
		'Set Profiles = New clsProfile
		'Set BatteryProfiles = New clsProfile
		
		' holds performance class objects
		'Set PerformanceProfiles = New clsProfile
		'------------------------------------------
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'-----------------obsolete?
		' i believe so.  All of these items implement cInode and
		'  are added via XML script just like every other node in the tree
		'Set WeaponProfiles = Nothing
		'Set Profiles = Nothing
		'Set BatteryProfiles = Nothing
		'Set PerformanceProfiles = Nothing
		
		'---------------------------------
		Dim I As Integer
		
		'UPGRADE_NOTE: Object Author may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Author = Nothing
		'UPGRADE_NOTE: Object Description may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Description = Nothing
		'UPGRADE_NOTE: Object Version may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Version = Nothing
		
		For I = 0 To m_lngChildCount - 1
			'UPGRADE_NOTE: Object m_oChildren() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			m_oChildren(I) = Nothing
		Next 
		For I = 0 To m_lngPropCount - 1
			'UPGRADE_NOTE: Object m_oProperties() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			m_oProperties(I) = Nothing
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
	'/////////////////////////////////////////////
	'//Implemented cINode Properties and Functions
	Private Function cINode_AddChild(ByRef oNode As _cINode) As Boolean Implements _cINode.AddChild
		Dim oBase As Object
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
		Dim I As Integer
		Dim lRet As Integer
		lRet = -1
		For I = 0 To m_lngChildCount - 1
			If m_oChildren(I).Handle = h Then lRet = I : Exit For
		Next 
		cINode_getChildIndexByHandle = lRet
	End Function
	Private Function cINode_getChild(ByVal lngIndex As Integer) As _cINode Implements _cINode.getChild
		If (lngIndex >= 0) And (m_lngChildCount > 0) And (lngIndex <= m_lngChildCount - 1) Then
			cINode_getChild = m_oChildren(lngIndex)
		End If
	End Function
	Private Function cINode_removeChild(ByVal lngIndex As Integer) As Boolean Implements _cINode.removeChild
		Dim I As Integer
		If (lngIndex <= m_lngChildCount - 1) And (lngIndex >= 0) Then
			'UPGRADE_NOTE: Object m_oChildren() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			m_oChildren(lngIndex) = Nothing
			For I = lngIndex + 1 To m_lngChildCount - 1
				m_oChildren(I - 1) = m_oChildren(I)
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
		End Get
	End Property
	Private ReadOnly Property cIPersist_GUID() As String
		Get
		End Get
	End Property
	
	Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Integer)
		Dim I As Integer
		
		If iMode = modGVDXMLSchemaConstants.GVD_XML_TYPE.cmp Then
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sName = op.Load(XML_NODE_NAME)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sDescription = op.Load(XML_NODE_DESCRIPTION)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngChildCount = op.Load(XML_NODE_CHILDCOUNT)
			
			If m_lngChildCount > 0 Then
				ReDim m_oChildren(m_lngChildCount - 1)
				For I = 0 To m_lngChildCount - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_oChildren(I) = op.Load(XML_NODE_CHILD & I)
					m_oChildren(I).Parent = m_hMe
					System.Diagnostics.Debug.Assert(m_oChildren(I).Parent <> 0, "")
				Next 
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Author = op.Load(XML_NODE_AUTHOR)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Version = op.Load("version")
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Description = op.Load("description")
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sImage = op.Load(XML_NODE_IMAGE)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngPropCount = op.Load(XML_NODE_PROPERTYCOUNT)
			
			If m_lngPropCount > 0 Then
				ReDim m_oProperties(m_lngPropCount - 1)
				For I = 0 To m_lngPropCount - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_oProperties(I) = op.Load(XML_NODE_PROPERTY & I)
				Next 
			End If
		End If
		'todo: check that child/member objects are "not" Is  Nothing after "load" attempt? or some other way?
	End Sub
	Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
	End Sub
End Class