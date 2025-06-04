Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cStats_NET.cStats")> Public Class cStats
	Implements _cIPersist
	Implements _cINode
	Implements _cIDisplay
	
	
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
		End Get
		Set(ByVal Value As String)
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
	Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Integer)
		Dim i As Integer
		If iMode = modGVDXMLSchemaConstants.GVD_XML_TYPE.cmp Then
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sName = op.Load(XML_NODE_NAME)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sDescription = op.Load(XML_NODE_DESCRIPTION)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngChildCount = op.Load(XML_NODE_CHILDCOUNT)
			
			System.Diagnostics.Debug.Assert(m_lngChildCount = 0, "")
		Else 'DEF
			
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
			System.Diagnostics.Debug.Assert(m_lngMaxChildren = 0, "")
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
	
	
	' just remember this, using objects for this crap will let me use shell interfaces for each component
	'yet handle all the bulk chores with one object with one set of code instead of typing the same
	' code handling in every single class since i cant inherit that code with VB.
	
	'
	
	Public Function calcStats(ByRef oComponent As _cIComponent) As Boolean
		
#If DEBUG_MODE Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		Dim oNode As cINode
		
		Set oNode = oComponent
		InfoPrint 1, "calculating stats for " & oNode.Description
		Set oNode = Nothing
#End If
	End Function
End Class