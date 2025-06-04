Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cArmor_NET.cArmor")> Public Class cArmor
	Implements _cINode
	Implements _cIDisplay
	Implements _cIPersist
	
	
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
	
	
	'Private m_Table() As Single 'todo: needed?
	
	
	'local variable(s) to hold property value(s)
	Private m_dblWeight As Double
	Private m_dblCost As Double
	Private m_dblSurfaceArea As Double
	Private m_sngAverageDR As Single
	Private m_sngAveragePD As Single
	
	
	Private m_lngArmorType As Integer ' e.g. overall, location, component, basic, complex, wheel,skirt,shield
	
	
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
		Debug.Print(TypeName(Me) & ":cIDisplay:getFirstPropertyItem -- no properties in m_oProperties() array.")
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
	
	'//Implemented cIComponent Properties and Functions
	Private Function cINode_AddChild(ByRef oBase As _cINode) As Boolean Implements _cINode.AddChild
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
	
	
	'///////////////////////////////////////////
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
	
	
	
	Public ReadOnly Property Weight() As Double
		Get
			Weight = 20.2 'm_dblWeight
		End Get
	End Property
	Public ReadOnly Property Cost() As Double
		Get
			Cost = 1234.56 ' m_dblCost
		End Get
	End Property
	Public ReadOnly Property SurfaceArea() As Double
		Get
			SurfaceArea = 1.9 ' m_dblSurfaceArea
		End Get
	End Property
	Public ReadOnly Property AverageDR() As Single
		Get
			AverageDR = m_sngAverageDR
		End Get
	End Property
	Public ReadOnly Property AveragePD() As Single
		Get
			AveragePD = m_sngAveragePD 'todo: i guess this would mean, average across faces? or maybe its irrelevant?  i should delete this
		End Get
	End Property
	Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Integer)
		
		Dim i As Integer
		If iMode = modGVDXMLSchemaConstants.GVD_XML_TYPE.cmp Then
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sName = op.Load(XML_NODE_NAME)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sDescription = op.Load(XML_NODE_DESCRIPTION)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngChildCount = op.Load(XML_NODE_CHILDCOUNT)
			
			If m_lngChildCount > 0 Then
				ReDim m_oChildren(m_lngChildCount - 1)
				For i = 0 To m_lngChildCount - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_oChildren(i) = op.Load(XML_NODE_CHILD & i)
					m_oChildren(i).Parent = m_hMe
					System.Diagnostics.Debug.Assert(m_oChildren(i).Parent <> 0, "")
				Next 
			End If
		Else 'DEF
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sImage = op.Load(XML_NODE_IMAGE)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngPropCount = op.Load(XML_NODE_PROPERTYCOUNT)
			'm_Table = op.Load(XML_NODE_STATSTABLE)
			If m_lngPropCount > 0 Then
				ReDim m_oProperties(m_lngPropCount - 1)
				For i = 0 To m_lngPropCount - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_oProperties(i) = op.Load("property" & i)
				Next 
			End If
		End If
	End Sub
	Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
	End Sub
	'//
	
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
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Public Sub Init()
		Dim mvarTL As Object
		Dim mvarParent As Object
		Dim DR6 As Object
		Dim DR5 As Object
		Dim DR4 As Object
		Dim DR3 As Object
		Dim DR2 As Object
		Dim DR1 As Object
		Dim Quality6 As Object
		Dim Quality5 As Object
		Dim Quality4 As Object
		Dim Quality3 As Object
		Dim Quality2 As Object
		Dim Quality1 As Object
		Dim Material6 As Object
		Dim Material5 As Object
		Dim Material4 As Object
		Dim Material3 As Object
		Dim Material2 As Object
		Dim Material1 As Object
		Dim DR As Object
		Dim Quality As Object
		Dim Material As Object
		Dim mvarDatatype As Object
		
		
		Select Case mvarDatatype
			
			Case ArmorComplexFacing
				'UPGRADE_WARNING: Couldn't resolve default property of object Material. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object DR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR = 1
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Material1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material1 = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Material2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material2 = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Material3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material3 = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Material4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material4 = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Material5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material5 = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Material6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material6 = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality1 = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality2 = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality3 = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality4 = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality5 = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality6 = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object DR1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR1 = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object DR2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR2 = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object DR3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR3 = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object DR4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR4 = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object DR5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR5 = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarParent).Datatype = Body Then
					'UPGRADE_WARNING: Couldn't resolve default property of object DR6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					DR6 = 1
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object DR6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					DR6 = 0
				End If
			Case ArmorBasicFacing
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Material. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object DR1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR1 = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object DR2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR2 = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object DR3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR3 = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object DR4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR4 = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object DR5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR5 = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarParent).Datatype = Body Then
					'UPGRADE_WARNING: Couldn't resolve default property of object DR6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					DR6 = 1
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object DR6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					DR6 = 0
				End If
			Case ArmorOpenFrame
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Material. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object DR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR = 1
				
			Case ArmorGunShield
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Material. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object DR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR = 1
				
			Case ArmorLocation
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Material. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object DR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR = 1
				
			Case ArmorComponent
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Material. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object DR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR = 1
				
			Case ArmorOverall
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Material. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object DR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR = 1
				
			Case ArmorWheelGuard
				'UPGRADE_WARNING: Couldn't resolve default property of object Material. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Material = "wood"
				'UPGRADE_WARNING: Couldn't resolve default property of object Quality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Quality = "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object DR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DR = 1
				
		End Select
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarTL. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarTL = Veh.Components(mvarParent).TL
		
	End Sub
	
	
	
	Public Sub StatsUpdate()
		Dim mvarQuality As Object
		Dim mvarCost As Object
		Dim mvarWeight As Object
		Dim GetSideLetterFromNumber As Object
		Dim mvarQuality6 As Object
		Dim mvarQuality5 As Object
		Dim mvarQuality4 As Object
		Dim mvarQuality3 As Object
		Dim mvarQuality2 As Object
		Dim mvarQuality1 As Object
		Dim mvarCoating As Object
		Dim mvarRadiation As Object
		Dim mvarThermal As Object
		Dim mvarElectrified As Object
		Dim mvarRAP As Object
		Dim mvarDatatype As Object
		Dim mvarPD As Object
		Dim mvarMaterial As Object
		Dim mvarDR As Object
		Dim mvarEffectiveDR6 As Object
		Dim mvarEffectiveDR5 As Object
		Dim mvarEffectiveDR4 As Object
		Dim mvarEffectiveDR3 As Object
		Dim mvarEffectiveDR2 As Object
		Dim mvarEffectiveDR1 As Object
		Dim mvarPD6 As Object
		Dim mvarMaterial6 As Object
		Dim mvarDR6 As Object
		Dim mvarPD5 As Object
		Dim mvarMaterial5 As Object
		Dim mvarDR5 As Object
		Dim mvarPD4 As Object
		Dim mvarMaterial4 As Object
		Dim mvarDR4 As Object
		Dim mvarPD3 As Object
		Dim mvarMaterial3 As Object
		Dim mvarDR3 As Object
		Dim mvarPD2 As Object
		Dim mvarMaterial2 As Object
		Dim mvarDR2 As Object
		Dim mvarPD1 As Object
		Dim mvarMaterial1 As Object
		Dim mvarDR1 As Object
		Dim CalcPD As Object
		Dim mvarParent As Object
		Dim mvarLocation As Object
		Dim GetLocation As Object
		Dim mvarPrintOutput As Object
		Dim mvarZZInit As Object
		Dim component As Short
		Dim SlopeR As String
		Dim SlopeL As String
		Dim SlopeF As String
		Dim SlopeB As String
		'UPGRADE_WARNING: Lower bound of array sCompare was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim sCompare(6) As String
		Dim sSides() As String
		Dim i As Integer
		Dim j As Integer
		Dim count As Integer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarZZInit. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarZZInit = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarPrintOutput = "" ' reinit this var
		
		'UPGRADE_WARNING: Couldn't resolve default property of object GetLocation. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarLocation. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarLocation = GetLocation
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		component = Veh.Components(mvarParent).Datatype
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDatatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (mvarDatatype = ArmorBasicFacing) Or (mvarDatatype = ArmorComplexFacing) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SlopeR = Veh.Components(mvarParent).SlopeR
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SlopeL = Veh.Components(mvarParent).SlopeL
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SlopeF = Veh.Components(mvarParent).SlopeF
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SlopeB = Veh.Components(mvarParent).SlopeB
			
			
			CalcByFacingArmorWeightCost()
			'get the PD
			' Get the PD for each side (used for both Complex and Basic)
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPD1 = CalcPD(mvarDR1, SlopeR, mvarMaterial1)
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPD2 = CalcPD(mvarDR2, SlopeL, mvarMaterial2)
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPD3 = CalcPD(mvarDR3, SlopeF, mvarMaterial3)
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPD4 = CalcPD(mvarDR4, SlopeB, mvarMaterial4)
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPD5 = CalcPD(mvarDR5, "none", mvarMaterial5) 'the top and underside dont have slope
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPD6 = CalcPD(mvarDR6, "none", mvarMaterial6)
			
			'get the effective DR
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcEffectiveDR(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEffectiveDR1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEffectiveDR1 = CalcEffectiveDR(mvarDR1, SlopeR)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcEffectiveDR(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEffectiveDR2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEffectiveDR2 = CalcEffectiveDR(mvarDR2, SlopeL)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcEffectiveDR(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEffectiveDR3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEffectiveDR3 = CalcEffectiveDR(mvarDR3, SlopeF)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcEffectiveDR(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEffectiveDR4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEffectiveDR4 = CalcEffectiveDR(mvarDR4, SlopeB)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEffectiveDR5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEffectiveDR5 = mvarDR5
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEffectiveDR6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEffectiveDR6 = mvarDR6
			
			
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarDatatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarDatatype = ArmorLocation Then 
			
			CalcArmorWeightCost()
			If component = Body Or component = Superstructure Or component = Turret Or component = Popturret Then
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SlopeR = Veh.Components(mvarParent).SlopeR
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SlopeL = Veh.Components(mvarParent).SlopeL
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SlopeF = Veh.Components(mvarParent).SlopeF
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SlopeB = Veh.Components(mvarParent).SlopeB
				' Get the PD for each side (used for both Complex and Basic)
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarPD1 = CalcPD(mvarDR, SlopeR, mvarMaterial)
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarPD2 = CalcPD(mvarDR, SlopeL, mvarMaterial)
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarPD3 = CalcPD(mvarDR, SlopeF, mvarMaterial)
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarPD4 = CalcPD(mvarDR, SlopeB, mvarMaterial)
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarPD5 = CalcPD(mvarDR, "none", mvarMaterial)
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarPD6 = CalcPD(mvarDR, "none", mvarMaterial)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarPD = CalcPD(mvarDR, "none", mvarMaterial)
			End If
		Else
			CalcArmorWeightCost()
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcPD(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPD = CalcPD(mvarDR, "none", mvarMaterial)
		End If
		
		'print output
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRAP. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarRAP Then
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarElectrified. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarElectrified Then
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarThermal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarThermal Then
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRadiation. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarRadiation Then
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarCoating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarCoating <> "none" Then
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarPD Then
		End If
		
		Dim counter As Short
		Select Case mvarDatatype
			Case ArmorComplexFacing
				'can have different everything
				'get each side
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(1) = " PD " & VB6.Format(mvarPD1) & ", DR " & VB6.Format(mvarDR1) & " " + mvarQuality1 + " " + mvarMaterial1 + ". "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(2) = " PD " & VB6.Format(mvarPD2) & ", DR " & VB6.Format(mvarDR2) & " " + mvarQuality2 + " " + mvarMaterial2 + ". "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(3) = " PD " & VB6.Format(mvarPD3) & ", DR " & VB6.Format(mvarDR3) & " " + mvarQuality3 + " " + mvarMaterial3 + ". "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(4) = " PD " & VB6.Format(mvarPD4) & ", DR " & VB6.Format(mvarDR4) & " " + mvarQuality4 + " " + mvarMaterial4 + ". "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(5) = " PD " & VB6.Format(mvarPD5) & ", DR " & VB6.Format(mvarDR5) & " " + mvarQuality5 + " " + mvarMaterial5 + ". "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(6) = " PD " & VB6.Format(mvarPD6) & ", DR " & VB6.Format(mvarDR6) & " " + mvarQuality6 + " " + mvarMaterial6 + ". "
				'move the first side (the Right) into the sSides array
				ReDim sSides(2, 1)
				sSides(1, 1) = sCompare(1)
				sSides(2, 1) = "R"
				count = 1
				'compare each side to see if it can be grouped with another
				For j = 2 To 6
					counter = count
					For i = 1 To counter
						If sCompare(j) = sSides(1, i) Then
							sSides(1, i) = sCompare(j)
							'UPGRADE_WARNING: Couldn't resolve default property of object GetSideLetterFromNumber(j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSides(2, i) = sSides(2, i) & "," + GetSideLetterFromNumber(j)
						ElseIf i = count Then 
							ReDim Preserve sSides(2, i + 1)
							count = count + 1
							sSides(1, i + 1) = sCompare(j)
							'UPGRADE_WARNING: Couldn't resolve default property of object GetSideLetterFromNumber(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSides(2, i + 1) = GetSideLetterFromNumber(j)
						End If
					Next 
				Next 
				'get final string and include the surface options to the armor
				For i = 1 To count
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarPrintOutput = mvarPrintOutput + " " + sSides(2, i) + ": " + sSides(1, i)
				Next 
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarPrintOutput = mvarPrintOutput + " (" + VB6.Format(mvarWeight, p_sFormat) + " lbs., $" + VB6.Format(mvarCost, p_sFormat) + ")."
				
			Case ArmorBasicFacing
				'same material and quality but different DR's and PD's
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(1) = " PD " & VB6.Format(mvarPD1) & ", DR " & VB6.Format(mvarDR1) & " "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(2) = " PD " & VB6.Format(mvarPD2) & ", DR " & VB6.Format(mvarDR2) & " "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(3) = " PD " & VB6.Format(mvarPD3) & ", DR " & VB6.Format(mvarDR3) & " "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(4) = " PD " & VB6.Format(mvarPD4) & ", DR " & VB6.Format(mvarDR4) & " "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(5) = " PD " & VB6.Format(mvarPD5) & ", DR " & VB6.Format(mvarDR5) & " "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(6) = " PD " & VB6.Format(mvarPD6) & ", DR " & VB6.Format(mvarDR6) & " "
				'move the first side (the Right) into the sSides array
				ReDim sSides(2, 1)
				sSides(1, 1) = sCompare(1)
				sSides(2, 1) = "R"
				count = 1
				'compare each side to see if it can be grouped with another
				For j = 2 To 6
					counter = count
					For i = 1 To counter
						If sCompare(j) = sSides(1, i) Then
							sSides(1, i) = sCompare(j)
							'UPGRADE_WARNING: Couldn't resolve default property of object GetSideLetterFromNumber(j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSides(2, i) = sSides(2, i) & "," + GetSideLetterFromNumber(j)
						ElseIf i = count Then 
							ReDim Preserve sSides(2, i + 1)
							count = count + 1
							sSides(1, i + 1) = sCompare(j)
							'UPGRADE_WARNING: Couldn't resolve default property of object GetSideLetterFromNumber(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSides(2, i + 1) = GetSideLetterFromNumber(j)
						End If
					Next 
				Next 
				'get final string and include the surface options to the armor
				For i = 1 To count
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarPrintOutput = mvarPrintOutput + " " + sSides(2, i) + ": " + sSides(1, i)
				Next 
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarPrintOutput = mvarQuality + " " + mvarMaterial + mvarPrintOutput + " (" + VB6.Format(mvarWeight, p_sFormat) + " lbs., $" + VB6.Format(mvarCost, p_sFormat) + ")."
				
				
			Case ArmorLocation 'this can still have different PD's on Body and Turrets do to slope differences
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(1) = " PD " & VB6.Format(mvarPD1) & ", DR " & VB6.Format(mvarDR) & " "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(2) = " PD " & VB6.Format(mvarPD2) & ", DR " & VB6.Format(mvarDR) & " "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(3) = " PD " & VB6.Format(mvarPD3) & ", DR " & VB6.Format(mvarDR) & " "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(4) = " PD " & VB6.Format(mvarPD4) & ", DR " & VB6.Format(mvarDR) & " "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(5) = " PD " & VB6.Format(mvarPD5) & ", DR " & VB6.Format(mvarDR) & " "
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCompare(6) = " PD " & VB6.Format(mvarPD6) & ", DR " & VB6.Format(mvarDR) & " "
				'move the first side (the Right) into the sSides array
				ReDim sSides(2, 1)
				sSides(1, 1) = sCompare(1)
				sSides(2, 1) = "R"
				count = 1
				'compare each side to see if it can be grouped with another
				For j = 2 To 6
					counter = count
					For i = 1 To counter
						If sCompare(j) = sSides(1, i) Then
							sSides(1, i) = sCompare(j)
							'UPGRADE_WARNING: Couldn't resolve default property of object GetSideLetterFromNumber(j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSides(2, i) = sSides(2, i) & "," + GetSideLetterFromNumber(j)
						ElseIf i = count Then 
							ReDim Preserve sSides(2, i + 1)
							count = count + 1
							sSides(1, i + 1) = sCompare(j)
							'UPGRADE_WARNING: Couldn't resolve default property of object GetSideLetterFromNumber(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSides(2, i + 1) = GetSideLetterFromNumber(j)
						End If
					Next 
				Next 
				'get final string and include the surface options to the armor
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarPrintOutput = "DR " & VB6.Format(mvarDR) & " "
				For i = 1 To count
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarPrintOutput = mvarPrintOutput + " " + sSides(2, i) + ": " + sSides(1, i)
				Next 
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarPrintOutput = mvarPrintOutput + mvarQuality + " " + mvarMaterial + " (" + VB6.Format(mvarWeight, p_sFormat) + " lbs., $" + VB6.Format(mvarCost, p_sFormat) + ")."
				
				
			Case ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPD. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarPrintOutput = "PD " & VB6.Format(mvarPD) & ", DR " & VB6.Format(mvarDR) & " " + mvarQuality + " " + mvarMaterial + " (" + VB6.Format(mvarWeight, p_sFormat) + " lbs., $" + VB6.Format(mvarCost, p_sFormat) + ")."
				
		End Select
		
	End Sub
	
	
	
	Public Function FillMaterial() As String()
		Dim mvarTL As Object
		' populate the material combo
		Dim materialarray() As String
		ReDim materialarray(1)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarTL. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarTL <= 6 Then
			materialarray = mAddKey(materialarray, "wood")
			materialarray = mAddKey(materialarray, "metal")
			materialarray = mAddKey(materialarray, "nonrigid")
		Else ' if its greater than or equal to 7
			materialarray = mAddKey(materialarray, "wood")
			materialarray = mAddKey(materialarray, "metal")
			materialarray = mAddKey(materialarray, "ablative")
			materialarray = mAddKey(materialarray, "fireproof ablative")
			materialarray = mAddKey(materialarray, "nonrigid")
			materialarray = mAddKey(materialarray, "composite")
			materialarray = mAddKey(materialarray, "laminate")
		End If
		
		FillMaterial = VB6.CopyArray(materialarray)
	End Function
	
	Public Function FillQuality(ByRef sMaterial As String) As String()
		Dim mvarTL As Object
		Dim MaterialCombo As System.Windows.Forms.ComboBox
		Dim Selected As String ' holds the users selected Armor material
		Dim arrQuality() As Short 'holds the list of suitable quality
		Dim iSelected As Short 'holds converted Selected string
		Dim element As Object 'one element of the arrQuality array
		Dim i As Short 'counter
		Dim count As Short 'another counter
		Dim qualityarray() As String
		Dim TempTL As Short
		
		ReDim qualityarray(1)
		
		Const Cheap As Short = 1
		Const Standard As Short = 2
		Const Expensive As Short = 3
		Const Advanced As Short = 4
		
		'get the type of armor that the user selected
		Selected = sMaterial
		'convert the Selected into an integer
		Select Case Selected
			Case "wood"
				iSelected = 1
			Case "metal"
				iSelected = 2
			Case "ablative"
				iSelected = 3
			Case "fireproof ablative"
				iSelected = 4
			Case "nonrigid"
				iSelected = 5
			Case "composite"
				iSelected = 6
			Case "laminate"
				iSelected = 7
		End Select
		
		count = 1 ' init the counter
		'given the tech level, produce list of quality types
		'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempTL = modPerformance.Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
		'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempTL = modPerformance.Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
		For i = 1 To UBound(ArmorMatrix)
			If ArmorMatrix(i).TL = TempTL Then
				If ArmorMatrix(i).MaterialType = iSelected Then
					If ArmorMatrix(i).WeightMod <> 0 Then
						ReDim Preserve arrQuality(count)
						arrQuality(count) = ArmorMatrix(i).Quality
						count = count + 1
						If count > 5 Then Exit For Else 
					End If
				End If
			End If
		Next 
		
		'fill the Quality combo with the list of available items
		For	Each element In arrQuality
			Select Case element
				Case Cheap
					qualityarray = mAddKey(qualityarray, "cheap")
					'If TempText = "cheap" Then TextFlag = True
				Case Standard
					qualityarray = mAddKey(qualityarray, "standard")
					'If TempText = "standard" Then TextFlag = True
				Case Expensive
					qualityarray = mAddKey(qualityarray, "expensive")
					'If TempText = "expensive" Then TextFlag = True
				Case Advanced
					qualityarray = mAddKey(qualityarray, "advanced")
					'If TempText = "advanced" Then TextFlag = True
			End Select
		Next element
		
		FillQuality = VB6.CopyArray(qualityarray)
		
	End Function
	
	Sub CalcArmorWeightCost()
		Dim CalcSurfaceFeaturesCostandWeight As Object
		Dim mvarCost As Object
		Dim mvarWeight As Object
		Dim mvarDR As Object
		Dim mvarTL As Object
		Dim mvarQuality As Object
		Dim mvarMaterial As Object
		Dim mvarDatatype As Object
		Dim mvarParent As Object
		Dim TempTL As Short
		' This routine calculates the Cost and Weight of the armor.
		'These contstant values must match those in the module "modArmor" since
		'the armormatrix uses integers and not string names for the material and quality
		Const Cheap As Short = 1
		Const Standard As Short = 2
		Const Expensive As Short = 3
		Const Advanced As Short = 4
		
		Const Wood As Short = 1
		Const Metal As Short = 2
		Const Ablative As Short = 3
		Const FireproofAblative As Short = 4
		Const NonRigid As Short = 5
		Const Composite As Short = 6
		Const Laminate As Short = 7
		
		Dim Area As Single 'holds the surface area
		Dim CostModifier As Single
		Dim WeightModifier As Single
		Dim SelectedMaterial As Short
		Dim SelectedQuality As Short
		Dim i As Short ' counter
		
		' Get the surface area based on the armor datatype being used
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDatatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (mvarDatatype = ArmorWheelGuard) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Area = Veh.Components(mvarParent).SurfaceArea / 2
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarDatatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf (mvarDatatype = ArmorGunShield) Or (mvarDatatype = ArmorOpenFrame) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Area = Veh.Components(mvarParent).SurfaceArea / 5
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarDatatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf (mvarDatatype = ArmorLocation) Or (mvarDatatype = ArmorComponent) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Area = Veh.Components(mvarParent).SurfaceArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarDatatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarDatatype = ArmorOverall Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Area = Veh.Stats.StructuralSurfaceArea
		End If
		
		'Determine Material
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarMaterial = "wood" Then
			SelectedMaterial = Wood
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarMaterial = "metal" Then 
			SelectedMaterial = Metal
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarMaterial = "ablative" Then 
			SelectedMaterial = Ablative
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarMaterial = "fireproof ablative" Then 
			SelectedMaterial = FireproofAblative
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarMaterial = "nonrigid" Then 
			SelectedMaterial = NonRigid
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarMaterial = "composite" Then 
			SelectedMaterial = Composite
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarMaterial = "laminate" Then 
			SelectedMaterial = Laminate
		End If
		
		'Determine Quality
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarQuality = "cheap" Then
			SelectedQuality = Cheap
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarQuality = "standard" Then 
			SelectedQuality = Standard
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarQuality = "expensive" Then 
			SelectedQuality = Expensive
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarQuality = "advanced" Then 
			SelectedQuality = Advanced
		End If
		
		' Get the Cost and Weight Modifiers
		'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempTL = modPerformance.Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
		'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempTL = modPerformance.Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
		For i = 1 To UBound(ArmorMatrix)
			If ArmorMatrix(i).TL = TempTL Then
				If ArmorMatrix(i).MaterialType = SelectedMaterial Then
					If ArmorMatrix(i).Quality = SelectedQuality Then
						CostModifier = ArmorMatrix(i).Cost
						WeightModifier = ArmorMatrix(i).WeightMod
						Exit For
					End If
				End If
			End If
		Next 
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarWeight = mvarDR * Area * WeightModifier
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarCost = mvarWeight * CostModifier
		
		'get the final weight and cost by adding the cost/weight of the surface features
		CalcSurfaceFeaturesCostandWeight(Area)
	End Sub
	
	Sub CalcByFacingArmorWeightCost()
		Dim CalcSurfaceFeaturesCostandWeight As Object
		Dim mvarWeight As Object
		Dim mvarCost As Object
		Dim mvarQuality As Object
		Dim mvarMaterial As Object
		Dim mvarTL As Object
		Dim mvarDatatype As Object
		Dim mvarParent As Object
		Dim mvarDR6 As Object
		Dim mvarDR5 As Object
		Dim mvarDR4 As Object
		Dim mvarDR3 As Object
		Dim mvarDR2 As Object
		Dim mvarDR1 As Object
		Dim mvarQuality6 As Object
		Dim mvarQuality5 As Object
		Dim mvarQuality4 As Object
		Dim mvarQuality3 As Object
		Dim mvarQuality2 As Object
		Dim mvarQuality1 As Object
		Dim mvarMaterial6 As Object
		Dim mvarMaterial5 As Object
		Dim mvarMaterial4 As Object
		Dim mvarMaterial3 As Object
		Dim mvarMaterial2 As Object
		Dim mvarMaterial1 As Object
		' This routine calculates the Cost and Weight of the armor
		Dim Area As Single 'holds the surface area
		Dim CostModifier(6) As Single
		Dim WeightModifier(6) As Single
		Dim iWeight(6) As Single
		Dim iCost(6) As Single
		Dim SelectedMaterial(6) As String
		Dim SelectedQuality(6) As String
		Dim iSelectedQuality(6) As Short
		Dim iSelectedMaterial(6) As Short
		Dim count As Short
		Dim i As Short
		Dim arrMaterial(6) As String
		Dim arrQuality(6) As String
		Dim arrDR(6) As Integer
		Dim TempCost As Single
		Dim TempWeight As Single
		Dim TempTL As Short
		'fill the arrMaterial array and arrQuality
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrMaterial(0) = mvarMaterial1
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrMaterial(1) = mvarMaterial2
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrMaterial(2) = mvarMaterial3
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrMaterial(3) = mvarMaterial4
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrMaterial(4) = mvarMaterial5
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarMaterial6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrMaterial(5) = mvarMaterial6
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrQuality(0) = mvarQuality1
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrQuality(1) = mvarQuality2
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrQuality(2) = mvarQuality3
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrQuality(3) = mvarQuality4
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrQuality(4) = mvarQuality5
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuality6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrQuality(5) = mvarQuality6
		
		'fill the arrDR aray
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrDR(0) = mvarDR1
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrDR(1) = mvarDR2
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrDR(2) = mvarDR3
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrDR(3) = mvarDR4
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrDR(4) = mvarDR5
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrDR(5) = mvarDR6
		
		' re-init variables
		TempCost = 0
		TempWeight = 0
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Area = Veh.Components(mvarParent).SurfaceArea
		' There are just two paths, one for Complex and one for Basic
		Select Case mvarDatatype
			Case ArmorComplexFacing
				For count = 0 To 5
					' Get the quality and material of the armor
					SelectedMaterial(count) = arrMaterial(count)
					SelectedQuality(count) = arrQuality(count)
					'convert the Selected into an integer
					Select Case SelectedMaterial(count)
						Case "wood"
							iSelectedMaterial(count) = 1
						Case "metal"
							iSelectedMaterial(count) = 2
						Case "ablative"
							iSelectedMaterial(count) = 3
						Case "fireproof ablative"
							iSelectedMaterial(count) = 4
						Case "nonrigid"
							iSelectedMaterial(count) = 5
						Case "composite"
							iSelectedMaterial(count) = 6
						Case "laminate"
							iSelectedMaterial(count) = 7
					End Select
					Select Case SelectedQuality(count)
						Case "cheap"
							iSelectedQuality(count) = 1
						Case "standard"
							iSelectedQuality(count) = 2
						Case "expensive"
							iSelectedQuality(count) = 3
						Case "advanced"
							iSelectedQuality(count) = 4
					End Select
					' Get the Cost and Weight Modifiers
					'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempTL = modPerformance.Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
					'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempTL = modPerformance.Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
					For i = 1 To UBound(ArmorMatrix)
						If ArmorMatrix(i).TL = TempTL Then
							If ArmorMatrix(i).MaterialType = iSelectedMaterial(count) Then
								If ArmorMatrix(i).Quality = iSelectedQuality(count) Then
									CostModifier(count) = ArmorMatrix(i).Cost
									WeightModifier(count) = ArmorMatrix(i).WeightMod
									Exit For
								End If
							End If
						End If
					Next 
					'get the average DR
					CalcAverageDr()
					' Get the Cost and weight of each face
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarParent = "1_" Then
						iWeight(count) = Val(CStr(arrDR(count))) * (Area / 6) * WeightModifier(count)
					Else
						iWeight(count) = Val(CStr(arrDR(count))) * (Area / 5) * WeightModifier(count)
					End If
					iCost(count) = iWeight(count) * CostModifier(count)
					TempWeight = TempWeight + iWeight(count)
					TempCost = TempCost + iCost(count)
				Next 
				
			Case ArmorBasicFacing
				
				'convert the Selected into an integer
				Select Case mvarMaterial
					Case "wood"
						iSelectedMaterial(0) = 1
					Case "metal"
						iSelectedMaterial(0) = 2
					Case "ablative"
						iSelectedMaterial(0) = 3
					Case "fireproof ablative"
						iSelectedMaterial(0) = 4
					Case "nonrigid"
						iSelectedMaterial(0) = 5
					Case "composite"
						iSelectedMaterial(0) = 6
					Case "laminate"
						iSelectedMaterial(0) = 7
				End Select
				Select Case mvarQuality
					Case "cheap"
						iSelectedQuality(0) = 1
					Case "standard"
						iSelectedQuality(0) = 2
					Case "expensive"
						iSelectedQuality(0) = 3
					Case "advanced"
						iSelectedQuality(0) = 4
				End Select
				' Get the Cost and Weight Modifiers
				'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TempTL = modPerformance.Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TempTL = modPerformance.Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
				For i = 1 To UBound(ArmorMatrix)
					If ArmorMatrix(i).TL = TempTL Then
						If ArmorMatrix(i).MaterialType = iSelectedMaterial(0) Then
							If ArmorMatrix(i).Quality = iSelectedQuality(0) Then
								CostModifier(0) = ArmorMatrix(i).Cost
								WeightModifier(0) = ArmorMatrix(i).WeightMod
							End If
						End If
					End If
				Next 
				'call routine to calc averagedr
				CalcAverageDr()
				' Get the Final Cost and Final Weight
				TempWeight = AverageDR * Area * WeightModifier(0)
				TempCost = TempWeight * CostModifier(0)
		End Select
		
		
		'save these cost and weight results to the armor class
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarCost = TempCost
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarWeight = TempWeight
		
		'get the final cost which includes the cost of the surface features
		CalcSurfaceFeaturesCostandWeight(Area)
	End Sub
	
	
	Function CalcEffectiveDR(ByRef DR As Integer, ByRef Slope As String) As Object
		Dim Modifier As Single
		
		If Slope = "none" Then
			Modifier = 1
		ElseIf Slope = "30 degrees" Then 
			Modifier = 1.5
		ElseIf Slope = "60 degrees" Then 
			Modifier = 2
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcEffectiveDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalcEffectiveDR = DR * Modifier
	End Function
	
	
	Function GetLowestDR() As Integer
		Dim mvarDR As Object
		Dim mvarMaterial As Object
		Dim mvarDR6 As Object
		Dim mvarMaterial6 As Object
		Dim mvarDR5 As Object
		Dim mvarMaterial5 As Object
		Dim mvarDR4 As Object
		Dim mvarMaterial4 As Object
		Dim mvarDR3 As Object
		Dim mvarMaterial3 As Object
		Dim mvarDR2 As Object
		Dim mvarMaterial2 As Object
		Dim mvarDR1 As Object
		Dim mvarMaterial1 As Object
		Dim mvarDatatype As Object
		'//this function only gets called during Aerial performance calculations.
		'//its job is to return the DR of the armor. In the case of seperate DR's
		'//for each face, then  it will return the lowest one.
		'//only DR from metal, composite or laminate armor counts.. all other types
		'//return 0.
		On Error Resume Next
		Dim lngRetval As Integer
		
		Select Case mvarDatatype
			Case ArmorComplexFacing
				Select Case mvarMaterial1
					Case "metal", "composite", "laminate"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRetval = mvarDR1
						Select Case mvarMaterial2
							Case "metal", "composite", "laminate"
								'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								lngRetval = modPerformance.Minimum(mvarDR1, mvarDR2)
								Select Case mvarMaterial3
									Case "metal", "composite", "laminate"
										'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										lngRetval = modPerformance.Minimum(lngRetval, mvarDR3)
										Select Case mvarMaterial4
											Case "metal", "composite", "laminate"
												'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												lngRetval = modPerformance.Minimum(lngRetval, mvarDR4)
												Select Case mvarMaterial5
													Case "metal", "composite", "laminate"
														'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														lngRetval = modPerformance.Minimum(lngRetval, mvarDR5)
														Select Case mvarMaterial6
															Case "metal", "composite", "laminate"
																'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
																lngRetval = modPerformance.Minimum(lngRetval, mvarDR6)
															Case Else
																lngRetval = 0
														End Select
													Case Else
														lngRetval = 0
												End Select
											Case Else
												lngRetval = 0
										End Select
									Case Else
										lngRetval = 0
								End Select
							Case Else
								lngRetval = 0
						End Select
					Case Else
						lngRetval = 0
				End Select
				
			Case ArmorBasicFacing
				Select Case mvarMaterial
					Case "metal", "composite", "laminate"
						'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRetval = modPerformance.Minimum(mvarDR1, mvarDR2)
						'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRetval = modPerformance.Minimum(lngRetval, mvarDR3)
						'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRetval = modPerformance.Minimum(lngRetval, mvarDR4)
						'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRetval = modPerformance.Minimum(lngRetval, mvarDR5)
						'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRetval = modPerformance.Minimum(lngRetval, mvarDR6)
						
					Case Else
						lngRetval = 0
				End Select
				
			Case ArmorLocation, ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
				Select Case mvarMaterial
					Case "metal", "composite", "laminate"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRetval = mvarDR
					Case Else
						lngRetval = 0
				End Select
				
		End Select
		
		
		GetLowestDR = lngRetval
	End Function
	
	Function GetLowestCrushDepthDR() As Integer
		Dim mvarDR As Object
		Dim mvarDR6 As Object
		Dim mvarDR5 As Object
		Dim mvarDR4 As Object
		Dim mvarDR3 As Object
		Dim mvarDR2 As Object
		Dim mvarDR1 As Object
		Dim mvarDatatype As Object
		
		On Error Resume Next
		Dim lngRetval As Integer
		
		Select Case mvarDatatype
			Case ArmorComplexFacing
				
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = mvarDR1
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(mvarDR1, mvarDR2)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR3)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR4)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR5)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR6)
				
			Case ArmorBasicFacing
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(mvarDR1, mvarDR2)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR3)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR4)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR5)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR6)
				
				
			Case ArmorLocation, ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = mvarDR
				
		End Select
		
		
		GetLowestCrushDepthDR = lngRetval
	End Function
	
	Private Sub CalcAverageDr()
		Dim mvarAverageDR As Object
		Dim mvarParent As Object
		Dim mvarDR6 As Object
		Dim mvarDR5 As Object
		Dim mvarDR4 As Object
		Dim mvarDR3 As Object
		Dim mvarDR2 As Object
		Dim mvarDR1 As Object
		Dim i As Short
		Dim TempAverage As Integer
		Dim Divisor As Short
		Dim arrDR(5) As Integer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrDR(0) = Val(mvarDR1)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrDR(1) = Val(mvarDR2)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR3. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrDR(2) = Val(mvarDR3)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR4. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrDR(3) = Val(mvarDR4)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR5. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrDR(4) = Val(mvarDR5)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarDR6. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrDR(5) = Val(mvarDR6)
		
		TempAverage = 0
		
		For i = 0 To 5
			TempAverage = TempAverage + arrDR(i)
		Next 
		
		'find the Average DR (this is done only for Basic Armor by facing
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarParent = "1_" Then
			Divisor = 6
		Else
			Divisor = 5
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarAverageDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarAverageDR = System.Math.Round(TempAverage / Divisor, 2)
		
	End Sub
End Class