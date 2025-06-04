Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsVehicle_NET.clsVehicle")> Public Class clsVehicle
	Implements _cIComponent
	Implements _cIDisplay
	Implements _cINode
	Implements _cIPersist
	
	
	' -------cINode interface variables
	Private m_lngMaxChildren As Integer
	Private m_lngChildCount As Integer
	Private m_lngCurrentChild As Integer
	Private m_oChildren() As cINode
	Private m_hParent As Integer
	Private m_hMe As Integer
	Private m_sDescription As String
	Private m_sImage As String
	
	' -------cIDisplay interface variables
	Private m_lngPropCount As Integer
	Private m_lngCurrentPropItem As Integer
	'UPGRADE_ISSUE: cPropertyItem object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Private m_oProperties() As cPropertyItem
	
	' ------- component interface properties
	Private m_Table() As Single
	Private m_lngTL As Integer
	Private m_dblHitpoints As Double
	Private m_dblSurfaceArea As Double
	Private m_dblCost As Double
	Private m_dblVolume As Double
	Private m_dblWeight As Double
	
	'Public Stats As clsStats   'todo: these must all implement cIPersist
	'Public Crew As clsCrew
	'Public Surface As clsSurface
	'Public Options As clsOptions
	'Public Description As clsDescription
	
	
	'Public PowerProfiles As Collection
	'Public FuelProfiles As Collection
	'Public WeaponProfiles As clsProfile
	'Public PerformanceProfiles As clsProfile
	'Public BatteryProfiles As clsProfile
	'Public Profiles As clsProfile
	
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
		If m_lngMaxChildren = m_lngChildCount Then
			cINode_AddChild = False
		Else
			m_lngChildCount = m_lngChildCount + 1
			ReDim Preserve m_oChildren(m_lngChildCount)
			m_oChildren(m_lngChildCount) = oNode
			cINode_AddChild = True
		End If
	End Function
	Private Function cINode_getChildFromHandle(ByVal h As Integer) As _cINode
	End Function
	Private Function cINode_getChildrenByClassName(ByRef Classname As String, ByRef hChilds() As Integer) As Boolean Implements _cINode.getChildrenByClassName
	End Function
	Private Function cINode_getFirstChild() As _cINode
		m_lngCurrentChild = 0
		If m_lngCurrentChild <= m_lngChildCount Then
			If Not m_oChildren(m_lngCurrentChild) Is Nothing Then
				cINode_getFirstChild = m_oChildren(m_lngCurrentChild) 'todo:
			End If
		End If
	End Function
	Private Function cINode_getNextChild() As _cINode
		m_lngCurrentChild = m_lngCurrentChild + 1
		
		If m_lngCurrentChild <= m_lngChildCount - 1 Then
			If Not m_oChildren(m_lngCurrentChild) Is Nothing Then
				cINode_getNextChild = m_oChildren(m_lngCurrentChild)
			End If
		End If
	End Function
	Private Function cINode_RemoveChild(ByRef oNode As _cINode) As Boolean Implements _cINode.RemoveChild
	End Function
	Private ReadOnly Property cINode_ClassName() As String Implements _cINode.ClassName
		Get
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			cINode_ClassName = TypeName(Me)
		End Get
	End Property
	Private ReadOnly Property cINode_GraphDown() As Boolean
		Get
			cINode_GraphDown = True
		End Get
	End Property
	Private Property cINode_Parent() As Integer Implements _cINode.Parent
		Get
		End Get
		Set(ByVal Value As Integer)
		End Set
	End Property
	Private ReadOnly Property cINode_ContainerAbbrev() As String
		Get
		End Get
	End Property
	Private ReadOnly Property cINode_AllowUserDelete() As Boolean
		Get
		End Get
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
	Private Property cINode_Note() As String
		Get
		End Get
		Set(ByVal Value As String)
		End Set
	End Property
	Private Property cINode_Handle() As Integer Implements _cINode.Handle
		Get
			cINode_Handle = m_hMe
		End Get
		Set(ByVal Value As Integer)
			m_hMe = Value
		End Set
	End Property
	
	
	
	
	'//Implemented cIComponent Properties and Functions
	Private Property cIComponent_LogicalParent() As Integer Implements _cIComponent.LogicalParent
		Get
		End Get
		Set(ByVal Value As Integer)
		End Set
	End Property
	Private Property cIComponent_TL() As Integer Implements _cIComponent.TL
		Get
			cIComponent_TL = m_lngTL
		End Get
		Set(ByVal Value As Integer)
			m_lngTL = Value
		End Set
	End Property
	Private Property cIComponent_HitPoints() As Double Implements _cIComponent.HitPoints
		Get
		End Get
		Set(ByVal Value As Double)
		End Set
	End Property
	Private Property cIComponent_SurfaceArea() As Double Implements _cIComponent.SurfaceArea
		Get
		End Get
		Set(ByVal Value As Double)
		End Set
	End Property
	'End Sub
	Private Property cIComponent_Cost() As Double Implements _cIComponent.Cost
		Get
		End Get
		Set(ByVal Value As Double)
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
	
	
	'//cIPersist Interface
	Private ReadOnly Property cIPersist_Classname() As String
		Get
		End Get
	End Property
	Private ReadOnly Property cIPersist_GUID() As String
		Get
		End Get
	End Property
	
	Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As modGVDXMLSchemaConstants.GVD_XML_TYPE)
		Dim i As Integer
		
		If iMode = modGVDXMLSchemaConstants.GVD_XML_TYPE.cmp Then
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sDescription = op.Load(XML_NODE_DESCRIPTION)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngChildCount = op.Load(XML_NODE_CHILDCOUNT)
			
			If m_lngChildCount > 0 Then
				ReDim m_oChildren(m_lngChildCount - 1)
				For i = 0 To m_lngChildCount - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_oChildren(i) = op.Load(XML_NODE_CHILD & i)
				Next 
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sImage = op.Load(XML_NODE_IMAGE)
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
		'todo: check that child/member objects are "not" Is  Nothing after "load" attempt? or some other way?
	End Sub
	Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
	End Sub
	'//
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		'LoadMatrices 'load all of our component matrix data 'todo: obsolete this
		Veh = Me ' public reference to this class
		'Set Components = New Collection    'todo: obsolete this
		
		'todo: i shouldnt be Setting any of these here!  Instead, I should only check after loading from file that all
		' reqt objects are set
		'Set Crew = New clsCrew
		'Set Surface = New clsSurface
		'Set Options = New clsOptions
		'Set Description = New clsDescription
		'Set Stats = New clsStats
		
		' holds collection of link objects
		'Set WeaponProfiles = New clsProfile
		'Set Profiles = New clsProfile
		'Set BatteryProfiles = New clsProfile
		
		' holds performance class objects
		'Set PerformanceProfiles = New clsProfile
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		'todo: loop Set m_oProperties(0) = Nothing
		
		
		'Set Body = Nothing
		
		'Set Crew = Nothing
		'Set Surface = Nothing
		'Set Options = Nothing
		'Set Description = Nothing
		'Set Stats = Nothing
		'--
		
		'--
		'Set WeaponProfiles = Nothing
		'Set Profiles = Nothing
		'Set BatteryProfiles = Nothing
		'Set PerformanceProfiles = Nothing
		Dim i As Integer
		For i = 0 To m_lngChildCount - 1
			'UPGRADE_NOTE: Object m_oChildren() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			m_oChildren(i) = Nothing
		Next 
		
		For i = 0 To m_lngPropCount - 1
			'UPGRADE_NOTE: Object m_oProperties() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			m_oProperties(i) = Nothing
		Next 
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class