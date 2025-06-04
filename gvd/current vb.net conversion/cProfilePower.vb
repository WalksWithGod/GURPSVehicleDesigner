Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cProfilePower_NET.cProfilePower")> Public Class cProfilePower
	
	' This is a generic Profile class for handling EITHER Power Suppliers + Power Consumer associations or Fuel Suppliers and Fuel Consuming associations
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef hpvDest As Any, ByRef hpvSource As Any, ByVal cbCopy As Integer)
	
	Private m_SC_Group() As clsSupplyConsumeGroup ' a single profile can and will usually have multiple SCGroups
	Private m_sUnAssignedConsumerList() As String
	Private sKey As String ' the name of this particular profile given by the user
	'UPGRADE_ISSUE: TreeX object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Private m_TreeX As TreeX
	Private m_lngGroupCount As Integer
	Private m_lngUnAssignedCount As Integer
	Private m_lngProfileType As Integer
	Private mvarDescription As String
	
	Public Property UnAssignedConsumerList() As String()
		Get
			UnAssignedConsumerList = VB6.CopyArray(m_sUnAssignedConsumerList)
		End Get
		Set(ByVal Value() As String)
			Dim sArr() As String ' this is so lame... Option Base 1 arrays when passed from a function turn dyanmic arrays into Fixed Arrays!
			Dim i As Integer
			
			m_lngUnAssignedCount = UBound(Value)
			ReDim m_sUnAssignedConsumerList(m_lngUnAssignedCount)
			
			' this  hack keeps our m_sUnaassignedConsumerList dynamic and redim-able
			For i = 1 To m_lngUnAssignedCount
				m_sUnAssignedConsumerList(i) = Value(i)
			Next 
			
			' cant use direct assignment cuz it turns our m_sUnassignedConsumerList into a fixed array.  This only happens when
			' passing Option Base 1 arrays (even if they're dynamic!) to Subs/Functions and then trying to use them for assignment
			'm_sUnAssignedConsumerList = s
			
		End Set
	End Property
	
	Public ReadOnly Property GroupCount() As Integer
		Get
			GroupCount = m_lngGroupCount
		End Get
	End Property
	
	Public WriteOnly Property Tree() As Object
		Set(ByVal Value As Object)
			
			m_TreeX = Value
		End Set
	End Property
	
	
	Public Property ProfileType() As Integer
		Get
			ProfileType = m_lngProfileType
		End Get
		Set(ByVal Value As Integer)
			m_lngProfileType = Value
		End Set
	End Property
	
	
	
	
	Public Property Key() As String
		Get
			Key = sKey
		End Get
		Set(ByVal Value As String)
			sKey = Value
		End Set
	End Property
	
	
	Public Property Description() As String
		Get
			Description = mvarDescription
		End Get
		Set(ByVal Value As String)
			mvarDescription = Value
		End Set
	End Property
	
	Public Function Group(ByVal iIndex As Integer) As clsSupplyConsumeGroup
		Group = m_SC_Group(iIndex)
	End Function
	
	Public Sub Show()
		Dim i As Integer
		If Not m_TreeX Is Nothing Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_TreeX.RemoveAllItems. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_TreeX.RemoveAllItems()
			For i = 1 To m_lngGroupCount
				m_SC_Group(i).Show(m_TreeX)
			Next 
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.ActiveProfile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Veh.ActiveProfile = sKey
		System.Diagnostics.Debug.Assert(m_lngProfileType > 0, "")
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.ActiveProfileType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Veh.ActiveProfileType = m_lngProfileType
	End Sub
	
	Friend Function SortAvailableSuppliers(ByRef vArr As Object) As Object
		' actually this is stupid.  Every time a supplier or consumer is added to a vehicle,
		' these should be added to every profile by calling
		' Veh.AddConsumer or Veh.AddSupplier  <- the keymanager shouldnt be doing this since it doesnt
		' likewise, when a supplier or consumer is removed, Veh.RemoveConsumer and Veh.RemoveSupplier should be called.
		' HOWEVER, if a new profile is made, we automatically create a NEW group for each supplier.  But there is
		' NO SORTING needed since each is initially put in its own group.
		' need to manage these for anything else...
		' However, we do need a Collection class to hold all profiles rather than using a standard collection
		' Hrmm... but wait.  I do need to keep list of all suppliers and consumers in the keymanager also dont i?  Since
		' when a new profile is created after alot of components have already been added, these all need to be
		' received from the keymanager to the new profile.  Hrmm... think about this some more... trying to avoid redundant tracking of these
		' stupid keys
		
		
		' 05/15/02 MPJ - as stated above ("this is stupid") this function is not needed.  The way this should be handled
		' is all suppliers and consumers are added to individual key arrays in the keymanager.  This is the only place these
		' keys will be stored.
		' this function receives all available suppliers in the vehicle and Creates new groups
		' for those not already in a group
		
		'    Dim i As Long
		'    Dim s As String
		'
		'    If vArr(i) = "" Then Exit Function
		'    For i = 1 To UBound(vArr)
		'        For j = 1 To UBound(m_SC_Group.suppliers)
		'    Next
		
	End Function
	
	Public Function AssignConsumer(ByRef s As String, ByVal lngGroupIndex As Integer) As Object
		Dim l As Integer
		
		If RemoveUnAssignedConsumer(s) Then
			m_SC_Group(lngGroupIndex).AddConsumer(s)
		Else
			Debug.Print("clsProfilePower:AssignConsumer -- ERROR:  RemoveUnAssignedConsumer Returned FALSE")
		End If
	End Function
	Public Function MoveConsumer(ByRef lngConsumer As Object, ByRef lngOldGroup As Object, ByRef lngNewGroup As Object) As Object
		' called to move a consumer from one SC_Group to another
		Dim sKey As String
		'UPGRADE_WARNING: Couldn't resolve default property of object lngConsumer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object lngOldGroup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sKey = m_SC_Group(lngOldGroup).Consumer(lngConsumer)
		'UPGRADE_WARNING: Couldn't resolve default property of object lngConsumer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object lngOldGroup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_SC_Group(lngOldGroup).RemoveConsumerByIndex(lngConsumer)
		'UPGRADE_WARNING: Couldn't resolve default property of object lngNewGroup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_SC_Group(lngNewGroup).AddConsumer(sKey)
		
	End Function
	
	Public Function AddConsumer(ByRef s As String) As Integer
		' called when a new consuemr is added to the vehicle
		m_lngUnAssignedCount = UBound(m_sUnAssignedConsumerList)
		If (m_lngUnAssignedCount = 1) Then
			If (m_sUnAssignedConsumerList(1) = "") Then
				m_sUnAssignedConsumerList(m_lngUnAssignedCount) = s
				Exit Function
			End If
		End If
		
		m_lngUnAssignedCount = m_lngUnAssignedCount + 1
		
		ReDim Preserve m_sUnAssignedConsumerList(m_lngUnAssignedCount)
		m_sUnAssignedConsumerList(m_lngUnAssignedCount) = s
		Debug.Print("Consumer Added " & s)
	End Function
	
	Public Function RemoveConsumer(ByRef s As String) As Integer
		' called when consumer is removed from entire vehicle
		Dim i As Integer
		If Not RemoveUnAssignedConsumer(s) Then
			For i = 1 To m_lngGroupCount
				m_SC_Group(i).RemoveConsumer(s)
			Next 
		End If
	End Function
	
	Public Function UnAssignConsumer(ByVal iGroupIndex As Integer, ByVal iConsumerIndex As Integer) As Object
		Dim s As String
		
		s = m_SC_Group(iGroupIndex).Consumer(iConsumerIndex)
		m_SC_Group(iGroupIndex).RemoveConsumer(s)
		
		m_lngUnAssignedCount = m_lngUnAssignedCount + 1
		ReDim Preserve m_sUnAssignedConsumerList(m_lngUnAssignedCount)
		m_sUnAssignedConsumerList(m_lngUnAssignedCount) = s
		
	End Function
	
	Private Function RemoveUnAssignedConsumer(ByRef s As String) As Integer
		' returns TRUE if the element was found and deleted
		Dim i As Integer
		Dim j As Integer
		
		For i = 1 To m_lngUnAssignedCount - 1
			If s = m_sUnAssignedConsumerList(i) Then
				' found it, now remove it
				For j = i To m_lngUnAssignedCount - 1
					m_sUnAssignedConsumerList(j) = m_sUnAssignedConsumerList(j + 1)
				Next 
				RemoveUnAssignedConsumer = True
				m_lngUnAssignedCount = m_lngUnAssignedCount - 1
				ReDim Preserve m_sUnAssignedConsumerList(m_lngUnAssignedCount)
				Exit Function
			End If
		Next 
		
		' didnt find it, so maybe its the very last one in the array
		If s = m_sUnAssignedConsumerList(m_lngUnAssignedCount) Then
			m_lngUnAssignedCount = m_lngUnAssignedCount - 1
			ReDim Preserve m_sUnAssignedConsumerList(m_lngUnAssignedCount)
			RemoveUnAssignedConsumer = True
			Exit Function
		End If
		RemoveUnAssignedConsumer = False
		
	End Function
	
	
	Public Function MovePowerSystem(ByRef lngSupplier As Object, ByRef lngOldGroup As Object, ByRef lngNewGroup As Object) As Object
		' called to move a power system from one SC_Group to another
		Dim sKey As String
		'UPGRADE_WARNING: Couldn't resolve default property of object lngSupplier. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object lngOldGroup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sKey = m_SC_Group(lngOldGroup).Supplier(lngSupplier)
		'UPGRADE_WARNING: Couldn't resolve default property of object lngSupplier. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object lngOldGroup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_SC_Group(lngOldGroup).RemoveSupplierByIndex(lngSupplier)
		'UPGRADE_WARNING: Couldn't resolve default property of object lngNewGroup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_SC_Group(lngNewGroup).AddSupplier(sKey)
		Call RemoveEmptySCGroup()
	End Function
	
	Public Function CreateNewSCGroup() As Integer
		Dim O As clsSupplyConsumeGroup
		
		m_lngGroupCount = m_lngGroupCount + 1
		ReDim Preserve m_SC_Group(m_lngGroupCount)
		O = New clsSupplyConsumeGroup
		m_SC_Group(m_lngGroupCount) = O
		m_SC_Group(m_lngGroupCount).GroupIndex = m_lngGroupCount
		'UPGRADE_NOTE: Object O may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		O = Nothing
		CreateNewSCGroup = m_lngGroupCount
	End Function
	Public Function AddPowerSystem(ByRef s As String) As Integer
		' called when a new power system has been added to the vehicle.  These power systems
		' get added to a new SC_Group to which they are the only member
		Dim iNewGroup As Integer
		
		iNewGroup = CreateNewSCGroup
		
		m_SC_Group(iNewGroup).AddSupplier(s)
		
	End Function
	
	Public Function RemovePowerSystem(ByRef s As String) As Integer
		' called when a power system is removed from entire vehicle
		Dim i As Integer
		For i = 1 To UBound(m_SC_Group)
			m_SC_Group(i).RemoveSupplier(s)
		Next 
		Call RemoveEmptySCGroup()
	End Function
	
	Private Sub RemoveEmptySCGroup()
		Dim i As Integer
		' Only groups with BOTH no suppliers AND no consumers are deleted automatically for the user
		' remove any empty sc groups IN REVERSE ORDER so that we dont have a situtation where i = 1 in first itteration then i = 2
		' is pointing to 3rd element since first itteration caused a shift in the entire array
		For i = UBound(m_SC_Group) To 1 Step -1
			If (m_SC_Group(i).SupplierCount = 0) And (m_SC_Group(i).ConsumerCount = 0) Then
				RemoveSCGroup(i)
			End If
		Next 
	End Sub
	Private Function RemoveSCGroup(ByVal index As Integer) As Integer
		System.Diagnostics.Debug.Assert(m_lngGroupCount > 0, "")
		Dim i As Integer
		' CAUTION: This function expects that elements are removed from tail to head
		' it then shifts all groups to the left by 1 starting at the spot where a group was deleted
		
		If index < m_lngGroupCount Then
			For i = index To m_lngGroupCount - 1
				m_SC_Group(i) = m_SC_Group(i + 1)
				m_SC_Group(i).GroupIndex = i
			Next 
		Else
			'UPGRADE_NOTE: Object m_SC_Group() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			m_SC_Group(m_lngGroupCount) = Nothing ' after the shift, the last item is now a double reference of 2nd to last item and must be deleted
		End If
		
		m_lngGroupCount = m_lngGroupCount - 1
		ReDim Preserve m_SC_Group(m_lngGroupCount)
		System.Diagnostics.Debug.Assert(m_lngGroupCount >= 0, "")
		
	End Function
	
	Public Function InitSuppliers(ByRef vArr As Object) As Integer
		' This function updates our SCGroup listings to make sure that the list of suppliers contained within each Group is valid.
		' If there is a supplier in the vehicle but not assigned to a group, then that supplier is placed into a new group
		' If there is a supplier NOT in the vehicle, but is in a group, that supplier is deleted AND if its the sole member of that group
		' then the entire group is deleted
		
		' Actually, this is somewhat insane as well (this is what happens when you code for a few hours, stop for a few days, then try to
		' remember what it was you were thinking several days ago)  Why check that in existing profile, components in the available list match those
		' in the SC_Groups.  Fact is, when a component is deleted or added to a vehicle, these components should be removed/added to the
		' SC Groups in all profiles anyway.  So lets just concentrate on making this function a InitSuppliers() call which is only
		' used on NEW profiles
		
		'''''''''''''''    Insane Code commented out (leaving all this crap here just so I dont forget why this is stupid
		'''''''''''''''    Dim i As Long
		'''''''''''''''    Dim lngNumTotalSuppliersInVehicle As Long
		'''''''''''''''    Dim lngNumGroups As Long
		'''''''''''''''    Dim lngNumSuppliersInGroup As Long
		'''''''''''''''
		'''''''''''''''    lngNumTotalSuppliersInVehicle = UBound(vArr)
		'''''''''''''''    lngNumGroups = UBound(m_SC_Group)
		'''''''''''''''
		'''''''''''''''    For i = 0 To lngNumGroups
		'''''''''''''''        lngNumSuppliersInGroup = m_SC_Group(i).SupplierCount
		'''''''''''''''        For j = 0 To lngNumSuppliersInGroup
		'''''''''''''''            If m_SC_Group(i).Supplier(j) = vArr(H) Then bFound = True
		'''''''''''''''        Next
		'''''''''''''''    Next
		
		
		Dim i As Integer
		Dim lngNumTotalSuppliersInVehicle As Integer
		Dim oGroup As clsSupplyConsumeGroup
		Dim s As String
		
		lngNumTotalSuppliersInVehicle = UBound(vArr)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object vArr(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (lngNumTotalSuppliersInVehicle = 1) And (vArr(1) = "") Then Exit Function
		
		ReDim m_SC_Group(lngNumTotalSuppliersInVehicle)
		For i = 1 To lngNumTotalSuppliersInVehicle
			oGroup = New clsSupplyConsumeGroup
			m_SC_Group(i) = oGroup
			m_SC_Group(i).GroupIndex = i
			m_lngGroupCount = i
			'UPGRADE_NOTE: Object oGroup may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oGroup = Nothing
			'UPGRADE_WARNING: Couldn't resolve default property of object vArr(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = vArr(i)
			m_SC_Group(i).AddSupplier(s)
		Next 
		
		InitSuppliers = True
		Exit Function
err_Renamed: 
		modHelper.InfoPrint(0, "clsProfilePower:InitSuppliers() -- Error #" & Err.Number & " -- " & Err.Description)
		InitSuppliers = False
	End Function
	
	Public Function InitConsumers(ByRef vArr As Object) As Object
		Dim i As Integer
		Dim lngNumTotalConsumersInVehicle As Integer
		' when a profile is just created, all consumers are thrown into the "available" list and none
		' belong to any SC_Groups at this point
		
		' with an existing profile, when a new component is added, it also goes to "available" list and is
		' not assigned to any SC_Group
		
		' with an existing profile, when a component is deleted, it must be deleted from any SC_Group or Available list
		
		' the single biggest conditions is, with an existing profile, sorting the available list from those already assigned to a group
		' in the profile.
		' to simplify this, we could save the local unAssignedConsumerList... then we rely on accurate node adding/removal tracking to rebuild these
		' profiles correctly from saved .veh files.  As long as the files in the local unassignedconsumerlist as well as those assigned to Groups
		' are always accurate, there will never be a need to sort the "Total Available Consumers" list to produce an accuracte UnAssignedConsumerList.
		
		' in new profile, entire array of consuming components goto UnAssignedConsumerList
		'    m_sUnAssignedConsumerList = vArr
		'    lngNumTotalConsumersInVehicle = UBound(m_sUnAssignedConsumerList)
		'
		'    InitConsumers = True
		Exit Function
err_Renamed: 
		modHelper.InfoPrint(0, "clsProfilePower:InitConsumers() -- Error #" & Err.Number & " -- " & Err.Description)
		'UPGRADE_WARNING: Couldn't resolve default property of object InitConsumers. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InitConsumers = False
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		ReDim m_SC_Group(0)
		
		m_SC_Group(0) = New clsSupplyConsumeGroup
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		Dim i As Integer
		For i = 0 To UBound(m_SC_Group)
			'UPGRADE_NOTE: Object m_SC_Group() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			m_SC_Group(i) = Nothing
		Next 
		
		'UPGRADE_NOTE: Object m_TreeX may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_TreeX = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class