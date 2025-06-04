Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsSupplyConsumeGroup_NET.clsSupplyConsumeGroup")> Public Class clsSupplyConsumeGroup
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef hpvDest As Any, ByRef hpvSource As Any, ByVal cbCopy As Integer)
	Private m_sSupplier() As String 'holds keys of suppliers in this particular group
	Private m_sConsumer() As String ' holds keys of consumers in this particular group
	Private m_lngGroupIndex As Integer
	
	Private m_lngSupplierCount As Integer
	Private m_lngConsumerCount As Integer
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		m_lngConsumerCount = 0
		m_lngSupplierCount = 0
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		Erase m_sSupplier
		Erase m_sConsumer
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public WriteOnly Property GroupIndex() As Integer
		Set(ByVal Value As Integer)
			m_lngGroupIndex = Value
		End Set
	End Property
	
	Public Sub Show(ByRef O As Object)
		' check that there is at least one Supplier so we know if a Group node is needed
		Dim hGroup As Integer
		Dim hChildGroup As Integer
		Dim hSupplier As Integer
		Dim hConsumer As Integer
		'UPGRADE_ISSUE: TreeX object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oTree As TreeX
		Dim i As Integer
		oTree = O
		
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oTree.AddItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hGroup = oTree.AddItem(GROUP_NAME)
		'UPGRADE_WARNING: Couldn't resolve default property of object oTree.ItemData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oTree.ItemData(hGroup) = m_lngGroupIndex
		'UPGRADE_WARNING: Couldn't resolve default property of object oTree.AddItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hChildGroup = oTree.AddItem(CHILD_GROUP_NAME, hGroup, 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object oTree.ItemData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oTree.ItemData(hChildGroup) = m_lngGroupIndex
		
		If m_lngSupplierCount >= 1 Then
			For i = 1 To UBound(m_sSupplier)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oTree.AddItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				hSupplier = oTree.AddItem(Veh.Components(m_sSupplier(i)).CustomDescription, hGroup)
				'UPGRADE_WARNING: Couldn't resolve default property of object oTree.ItemData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oTree.ItemData(hSupplier) = i
			Next 
			'UPGRADE_WARNING: Couldn't resolve default property of object oTree.ExpandItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oTree.ExpandItem(hGroup) = True
		End If
		
		On Error GoTo err_Renamed
		
		If m_lngConsumerCount >= 1 Then
			For i = 1 To UBound(m_sConsumer)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oTree.AddItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				hConsumer = oTree.AddItem(Veh.Components(m_sConsumer(i)).CustomDescription, hChildGroup)
				'UPGRADE_WARNING: Couldn't resolve default property of object oTree.ItemData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oTree.ItemData(hConsumer) = i
			Next 
			'UPGRADE_WARNING: Couldn't resolve default property of object oTree.ExpandItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oTree.ExpandItem(hChildGroup) = True
		End If
err_Renamed: 
		'UPGRADE_NOTE: Object oTree may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oTree = Nothing
	End Sub
	
	Public Function Consumer(ByVal index As Integer) As String
		Consumer = m_sConsumer(index)
	End Function
	Public Function Supplier(ByVal index As Integer) As String
		Supplier = m_sSupplier(index)
	End Function
	
	Public Function ConsumerCount() As Integer
		On Error GoTo err_Renamed
		ConsumerCount = UBound(m_sConsumer)
		Exit Function
err_Renamed: 
		ConsumerCount = 0
	End Function
	Public Function SupplierCount() As Integer
		SupplierCount = UBound(m_sSupplier)
	End Function
	
	Public Function AddConsumer(ByVal s As String) As Integer
		' function accepts the Consumer's key value
		m_lngConsumerCount = m_lngConsumerCount + 1
		ReDim Preserve m_sConsumer(m_lngConsumerCount)
		
		m_sConsumer(m_lngConsumerCount) = s
		AddConsumer = True
		
	End Function
	
	Public Function RemoveConsumer(ByRef s As String) As Integer
		Dim index As Integer
		Dim bFound As Boolean
		
		For index = 1 To m_lngConsumerCount
			If m_sConsumer(index) = s Then
				bFound = True
				Exit For
			End If
		Next 
		
		If Not bFound Then Exit Function
		Call RemoveConsumerByIndex(index)
	End Function
	Public Function RemoveConsumerByIndex(ByVal index As Integer) As Integer
		Dim i As Integer
		
		If m_lngConsumerCount <> index Then
			For i = index To m_lngConsumerCount - 1
				m_sConsumer(i) = m_sConsumer(i + 1)
			Next 
		End If
		
		m_lngConsumerCount = m_lngConsumerCount - 1
		ReDim Preserve m_sConsumer(m_lngConsumerCount)
		System.Diagnostics.Debug.Assert(m_lngConsumerCount >= 0, "")
		
	End Function
	
	Public Function AddSupplier(ByRef s As String) As Integer
		m_lngSupplierCount = m_lngSupplierCount + 1
		ReDim Preserve m_sSupplier(m_lngSupplierCount)
		
		m_sSupplier(m_lngSupplierCount) = s
		AddSupplier = True
	End Function
	
	Public Function RemoveSupplier(ByRef s As String) As Integer
		Dim index As Integer
		Dim bFound As Boolean
		
		For index = 1 To m_lngSupplierCount
			If m_sSupplier(index) = s Then
				bFound = True
				Exit For
			End If
		Next 
		
		If Not bFound Then Exit Function
		Call RemoveSupplierByIndex(index)
		
	End Function
	
	Public Function RemoveSupplierByIndex(ByVal lngIndex As Integer) As Object
		Dim i As Integer
		
		If m_lngSupplierCount <> lngIndex Then
			For i = lngIndex To m_lngSupplierCount - 1
				m_sSupplier(i) = m_sSupplier(i + 1)
			Next 
			' copy memory cant work here as is unless I loop... im working with variable length string arrays and not a
			' continuous block of memory
			'CopyMemory m_sSupplier(lngIndex), m_sSupplier(lngIndex + 1), m_lngSupplierCount - lngIndex
		End If
		m_lngSupplierCount = m_lngSupplierCount - 1
		ReDim Preserve m_sSupplier(m_lngSupplierCount)
		System.Diagnostics.Debug.Assert(m_lngSupplierCount >= 0, "")
	End Function
End Class