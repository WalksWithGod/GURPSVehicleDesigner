Option Strict Off
Option Explicit On
Friend Class cBase
	Implements _cIPersist
	Implements _cIDisplay
	
	
	Private m_sName As String
	Private m_sClassname As String
	
	Private m_lngModifierCount As Integer
	'UPGRADE_WARNING: Untranslated statement in (Declarations). Please check source code.
	Private m_lngCurrentPropItem As Integer
	Private m_lngPropCount As Integer
	
	Private m_oAuthor As vehicles.cAuthor
	'UPGRADE_ISSUE: vehicles.cVersion object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Private m_oVersion As vehicles.cVersion
	
	Private m_sIcon As String 'path
	Private m_lngMatrix() As Single
	
	
	Public ReadOnly Property propertyCount() As Integer
		Get
			Dim m_lngPropertyCount As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object m_lngPropertyCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			propertyCount = m_lngPropertyCount
		End Get
	End Property
	
	Public ReadOnly Property className() As String
		Get
			className = m_sClassname
		End Get
	End Property
	
	'-----------------
	
	Private ReadOnly Property cIPersist_Classname() As String
		Get
		End Get
	End Property
	
	Private ReadOnly Property cIPersist_GUID() As String
		Get
		End Get
	End Property
	
	'//cIDisplay Implemented Properties and Functions
	Private Function cIDisplay_getFirstPropertyItem() As cpropertyitem Implements _cIDisplay.getFirstPropertyItem
		If Not m_oProperties(0) Is Nothing Then
			cIDisplay_getFirstPropertyItem = m_oProperties(0)
			m_lngCurrentPropItem = 0
		End If
	End Function
	
	Private Function cIDisplay_getNextPropertyItem() As cpropertyitem Implements _cIDisplay.getNextPropertyItem
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
	Private Function cIDisplay_getPropertyItemByIndex(ByVal iIndex As Integer) As cpropertyitem Implements _cIDisplay.getPropertyItemByIndex
		On Error Resume Next
		cIDisplay_getPropertyItemByIndex = m_oProperties(iIndex)
	End Function
	
	Private Sub cIPersist_LoadProperties(ByVal op As PersistenceManager.clsObjProperties, ByVal iMode As PersistenceManager.GVD_XML_TYPE)
		Dim PersistenceManager As Object
		Dim i As Integer
		Dim o As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_oAuthor = op.Load("author")
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_oVersion = op.Load("version")
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sClassname = op.Load("classname")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_lngPropCount = op.Load("propertycount")
		
		If m_lngPropCount > 0 Then
			ReDim m_oProperties(m_lngPropCount - 1)
			For i = 0 To m_lngPropCount - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_oProperties(i) = op.Load("property" & i)
			Next 
		End If
		
		
	End Sub
	
	Private Sub cIPersist_StoreProperties(ByVal op As PersistenceManager.clsObjProperties)
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		Dim i As Integer
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