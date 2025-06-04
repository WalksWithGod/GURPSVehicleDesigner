Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cProfile_NET.cProfile")> Public Class cProfile
	
	' weapon profiles only contain "suppliers" there are no consuming child devices to assign.
	'--------------------------------------------------------------
	' tracks active profile for mode switching (TODO: This crap shouldnt even be here... its a total GUI thing...)
	'Private m_sActiveWeaponProfile As String
	'Private m_sActiveBatteryProfile As String
	'Private m_sActiveFuelProfile As String
	'Private m_sActivePerformanceProfile As String
	Private m_lngActiveProfileType As Integer
	Private m_lngActiveCheckListType As Integer
	Private m_sActiveCheckList As String
	Private m_sActiveProfile As String
	Private m_sActiveComponent As String
	
	' link objects
	' What is a Profile object?  A profile is a user defined scenario which describes
	' how their vehicle manages and assigns power and fuel to devices which supply or use power and fuel.  In the
	' case of weapons, it defines which weapons will work together in a group.  In the case of performance, it defines
	' which thrust generating devices and which options shall be used to determine the performance of the vehicle.  A user
	' can create as many profiles for all of the above as they wish.
	
	
	Public Property ActiveProfile() As String
		Get
			ActiveProfile = m_sActiveProfile
		End Get
		Set(ByVal Value As String)
			m_sActiveProfile = Value
		End Set
	End Property
	
	
	Public Property ActiveProfileType() As Integer
		Get
			ActiveProfileType = m_lngActiveProfileType
		End Get
		Set(ByVal Value As Integer)
			m_lngActiveProfileType = Value
		End Set
	End Property
	
	
	Public Property ActiveCheckList() As String
		Get
			ActiveCheckList = m_sActiveCheckList
		End Get
		Set(ByVal Value As String)
			m_sActiveCheckList = Value
		End Set
	End Property
	
	
	Public Property ActiveCheckListType() As Integer
		Get
			ActiveCheckListType = m_lngActiveCheckListType
		End Get
		Set(ByVal Value As Integer)
			m_lngActiveCheckListType = Value
		End Set
	End Property
	'Public Property Get ActiveWeaponProfile() As String
	'    ActiveWeaponProfile = m_sActiveWeaponProfile
	'End Property
	'
	'Friend Property Let ActiveWeaponProfile(ByRef s As String)
	'    m_sActiveWeaponProfile = s
	'End Property
	'
	'Public Property Get ActiveBatteryProfile() As String
	'    ActiveBatteryProfile = m_sActiveBatteryProfile
	'End Property
	'
	'Friend Property Let ActiveBatteryProfile(ByRef s As String)
	'    m_sActiveBatteryProfile = s
	'End Property
	'
	'Public Property Get ActivePerformanceProfile() As String
	'    ActivePerformanceProfile = m_sActivePerformanceProfile
	'End Property
	'
	'Friend Property Let ActivePerformanceProfile(ByRef s As String)
	'    m_sActivePerformanceProfile = s
	'End Property
	'
	'Public Property Get ActiveFuelProfile() As String
	'    ActiveFuelProfile = m_sActiveFuelProfile
	'End Property
	'
	'Friend Property Let ActiveFuelProfile(ByRef s As String)
	'    m_sActiveFuelProfile = s
	'End Property
	
	
	'todo: must implement cIPersist and cINode and cIDisplay
	'-------------
	Public Function AddProfile(ByRef oTreeX As TreeX, ByRef sProfileName As String, ByVal lngType As Integer, ByRef sName As String) As Integer
		Dim KeyManager As Object
		Dim Profiles As Object
		'UPGRADE_ISSUE: clsProfilePower object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim O As clsProfilePower
		Dim suppliers() As String
		Dim consumers() As String
		
		On Error GoTo err_Renamed
		O = New clsProfilePower
		'UPGRADE_WARNING: Couldn't resolve default property of object O.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		O.Key = sProfileName
		'UPGRADE_WARNING: Couldn't resolve default property of object O.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		O.Description = sName
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Profiles.Add. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Profiles.Add(O, sProfileName)
		'UPGRADE_NOTE: Object O may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		O = Nothing
		
		If lngType = FUEL_PROFILE Then
			' retreive list  of all suppliers in the vehicle
			'UPGRADE_WARNING: Couldn't resolve default property of object KeyManager.GetCurrentFuelStorageKeys. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			suppliers = VB6.CopyArray(KeyManager.GetCurrentFuelStorageKeys)
			' retreive list of all consumers in the vehicle
			'UPGRADE_WARNING: Couldn't resolve default property of object KeyManager.GetCurrentFuelUsingSystemKeys. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			consumers = VB6.CopyArray(KeyManager.GetCurrentFuelUsingSystemKeys)
		Else
			System.Diagnostics.Debug.Assert(lngType = POWER_PROFILE, "")
			' retreive list  of all suppliers in the vehicle
			'UPGRADE_WARNING: Couldn't resolve default property of object KeyManager.GetCurrentPowerSystemKeys. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			suppliers = VB6.CopyArray(KeyManager.GetCurrentPowerSystemKeys)
			' retreive list of all consumers in the vehicle
			'UPGRADE_WARNING: Couldn't resolve default property of object KeyManager.GetCurrentPowerConsumptionKeys. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			consumers = VB6.CopyArray(KeyManager.GetCurrentPowerConsumptionKeys)
		End If
		
		' pass the list of suppliers to the Profile's .InitSuppliers (suppliers)
		'UPGRADE_WARNING: Couldn't resolve default property of object Profiles().InitSuppliers. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call Profiles(sProfileName).InitSuppliers(suppliers)
		
		' associate the treex to this profile
		'UPGRADE_WARNING: Couldn't resolve default property of object Profiles().Tree. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Profiles(sProfileName).Tree = oTreeX
		'UPGRADE_WARNING: Couldn't resolve default property of object Profiles().ProfileType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Profiles(sProfileName).ProfileType = lngType
		'UPGRADE_WARNING: Couldn't resolve default property of object Profiles().Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Profiles(sProfileName).Show()
		
		' pass the list of consumers to the new PowerProfile's .InitConsumers(consumers)
		'UPGRADE_WARNING: Couldn't resolve default property of object Profiles().UnAssignedConsumerList. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Profiles(sProfileName).UnAssignedConsumerList = VB6.CopyArray(consumers)
		
		AddProfile = True
		Exit Function
		
err_Renamed: 
		AddProfile = False
		Debug.Print("clsVehicle.AddProfile - ERROR# " & Err.Number & " " & Err.Description)
	End Function
	
	Public Function AddConsumerToAllProfiles(ByRef sKey As String, ByVal lngProfileType As Integer) As Object
		Dim Profiles As Object
		' When a FuelSystem component is added to the vehicle, this is called
		' to add this components key reference to all fuel profiles that exist
		'UPGRADE_ISSUE: clsProfilePower object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim O As clsProfilePower
		
		For	Each O In Profiles
			'UPGRADE_WARNING: Couldn't resolve default property of object O.ProfileType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object O.AddConsumer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If O.ProfileType = lngProfileType Then O.AddConsumer(sKey)
		Next O
	End Function
	
	Public Function RemoveConsumerFromAllProfiles(ByRef sKey As String, ByVal lngProfileType As Integer) As Object
		Dim Profiles As Object
		' When a component is deleted from the vehicle, this is called to
		' remove this components key reference from all fuel profiles that exist
		'UPGRADE_ISSUE: clsProfilePower object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim O As clsProfilePower
		
		For	Each O In Profiles
			'UPGRADE_WARNING: Couldn't resolve default property of object O.ProfileType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object O.RemoveConsumer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If O.ProfileType = lngProfileType Then O.RemoveConsumer(sKey)
		Next O
	End Function
	
	Public Function AddSupplierToAllProfiles(ByRef sKey As String, ByVal lngProfileType As Integer) As Object
		Dim Profiles As Object
		' When a Fuel Supplier (tank) component is added to the vehicle, this is called
		' to add this components key reference to all power profiles that exist
		'UPGRADE_ISSUE: clsProfilePower object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim O As clsProfilePower
		
		For	Each O In Profiles
			'UPGRADE_WARNING: Couldn't resolve default property of object O.ProfileType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object O.AddPowerSystem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If O.ProfileType = lngProfileType Then O.AddPowerSystem(sKey)
		Next O
	End Function
	
	Public Function RemoveSupplierFromAllProfiles(ByRef sKey As String, ByVal lngProfileType As Integer) As Object
		Dim Profiles As Object
		' When a tank is deleted from the vehicle, this is called to
		' remove this components key reference from all power profiles that exist
		'UPGRADE_ISSUE: clsProfilePower object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim O As clsProfilePower
		
		'NOTE: Checking for orphaned children is done in the RemoveEmptySCGroup function inside of clsPowerProfiles
		For	Each O In Profiles
			'UPGRADE_WARNING: Couldn't resolve default property of object O.ProfileType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object O.RemovePowerSystem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If O.ProfileType = lngProfileType Then O.RemovePowerSystem(sKey)
		Next O
	End Function
	
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		' todo: init any arrays
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		' todo: terminate all profiles in here
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class