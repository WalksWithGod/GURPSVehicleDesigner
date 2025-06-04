Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cArmorFace_NET.cArmorFace")> Public Class cArmorFace
	Implements _cIPersist
	Implements _cIDisplay
	Implements _cINode
	
	
	
	Private m_lngPropCount As Integer
	Private m_lngCurrentPropItem As Integer
	'UPGRADE_ISSUE: cPropertyItem object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Private m_oProperties() As cPropertyItem
	
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
	
	
	
	Private m_lngTL As Integer
	
	
	' armor specific crap
	Private m_bRAP As Boolean
	Private m_bElectrified As Boolean
	Private m_bThermalCoating As Boolean
	Private m_bRadShielding As Boolean
	Private m_ReflectiveCoating As String
	Private m_lngPD As Integer
	Private m_lngDR As Integer '<--- todo: need more space? DR is cumlative in the "Face" since it adds all layer's DR
	Private m_dblSurfaceArea As Double
	Private m_dblWeight As Double
	Private m_dblCost As Double
	
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
	
	
	Public Property RAP() As Boolean
		Get
			RAP = m_bRAP
		End Get
		Set(ByVal Value As Boolean)
			m_bRAP = Value
		End Set
	End Property
	Public Property Electrified() As Boolean
		Get
			Electrified = m_bElectrified
		End Get
		Set(ByVal Value As Boolean)
			m_bElectrified = Value
		End Set
	End Property
	Public Property ThermalCoating() As Boolean
		Get
			ThermalCoating = m_bThermalCoating
		End Get
		Set(ByVal Value As Boolean)
			m_bThermalCoating = Value
		End Set
	End Property
	Public Property RadShielding() As Boolean
		Get
			RadShielding = m_bRadShielding
		End Get
		Set(ByVal Value As Boolean)
			m_bRadShielding = Value
		End Set
	End Property
	Public Property ReflectiveCoating() As String
		Get
			Dim m_sReflectiveCoating As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object m_sReflectiveCoating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ReflectiveCoating = m_sReflectiveCoating
		End Get
		Set(ByVal Value As String)
			Dim m_sReflectiveCoating As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object m_sReflectiveCoating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sReflectiveCoating = Value
		End Set
	End Property
	Public Property PD() As Integer
		Get
			PD = m_lngPD
		End Get
		Set(ByVal Value As Integer)
			m_lngPD = Value
		End Set
	End Property
	Public Property DR() As Integer
		Get
			DR = m_lngDR
		End Get
		Set(ByVal Value As Integer)
			m_lngDR = Value
		End Set
	End Property
	Public ReadOnly Property Weight(ByVal D As Double) As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Weight = m_dblWeight
		End Get
	End Property
	Public ReadOnly Property Cost(ByVal D As Double) As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object Cost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cost = m_dblCost
		End Get
	End Property
	Public ReadOnly Property SurfaceArea(ByVal D As Double) As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object SurfaceArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SurfaceArea = m_dblSurfaceArea
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
	
	
	Private Function GetSideLetterFromNumber(ByRef i As Integer) As String
		Select Case i
			Case 1
				GetSideLetterFromNumber = "R"
			Case 2
				GetSideLetterFromNumber = "L"
			Case 3
				GetSideLetterFromNumber = "F"
			Case 4
				GetSideLetterFromNumber = "B"
			Case 5
				GetSideLetterFromNumber = "T"
			Case 6
				GetSideLetterFromNumber = "U"
		End Select
	End Function
	
	Function CalcPD(ByRef DR As Integer, ByRef Slope As String, ByRef Material As String) As Integer
		Dim PD As Short
		
		If DR = 0 Then
			PD = 0
		ElseIf DR = 1 Then 
			PD = 1
		ElseIf DR <= 4 Then 
			PD = 2
		ElseIf DR <= 15 Then 
			PD = 3
		ElseIf DR >= 16 Then 
			PD = 4
		End If
		
		'check for max values for Wood and Nonrigid armor
		If Material = "nonrigid" Then
			If PD > 2 Then PD = 2
		ElseIf Material = "wood" Then 
			If PD > 3 Then PD = 3
		End If
		
		'add bonus for slope
		'Note: the bonus's are placed below the checks for  max values for wood and nonrigid.
		'if users request, it can be moved above it
		If Slope = "none" Then
		ElseIf Slope = "30 degrees" Then 
			PD = PD + 1
		ElseIf Slope = "60 degrees" Then 
			PD = PD + 2
		End If
		
		CalcPD = PD
	End Function
	
	Sub CalcSurfaceFeaturesCostandWeight(ByRef SurfaceArea As Single)
		Dim mvarWeight As Object
		Dim mvarCost As Object
		Dim mvarElectrified As Object
		Dim mvarRAP As Object
		Dim mvarThermal As Object
		Dim mvarCoating As Object
		Dim mvarRadiation As Object
		Const RadWeight As Short = 2
		Const RadCost As Short = 20
		Const ReflectCost As Short = 30
		Const RetroCost As Short = 150
		Const ThermCost As Short = 250
		Const ThermWeight As Double = 0.25
		Const RAPCost As Short = 20
		Const RAPWeight As Short = 8
		Const ElectCost As Short = 10
		Const ElectWeight As Double = 0.2
		Dim TempWeight As Single
		Dim TempCost As Single
		
		If SurfaceArea = 0 Then Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRadiation. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarRadiation Then
			TempWeight = RadWeight * SurfaceArea
			TempCost = RadCost * SurfaceArea
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarCoating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarCoating = "reflective" Then
			TempCost = TempCost + (ReflectCost * SurfaceArea)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarCoating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarCoating = "retro-reflective" Then 
			TempCost = TempCost + (RetroCost * SurfaceArea)
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarThermal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarThermal Then
			TempWeight = TempWeight + (ThermWeight * SurfaceArea)
			TempCost = TempCost + (ThermCost * SurfaceArea)
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRAP. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarRAP Then
			TempWeight = TempWeight + (RAPWeight * SurfaceArea)
			TempCost = TempCost + (RAPCost * SurfaceArea)
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarElectrified. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarElectrified Then
			TempWeight = TempWeight + (ElectWeight * SurfaceArea)
			TempCost = TempCost + (ElectCost * SurfaceArea)
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarCost = System.Math.Round(mvarCost + TempCost, 2)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarWeight = System.Math.Round(mvarWeight + TempWeight, 2)
		
	End Sub
End Class