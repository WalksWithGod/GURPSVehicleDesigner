Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsPerformanceSkid_NET.clsPerformanceSkid")> Public Class clsPerformanceSkid
	'local variable(s) to hold property value(s)
	Private mvartotalramjetthrust As Single
	Private mvarTotalTurboRamJetThrust As Single
	Private mvarTotalRamJetThrustAB As Single
	Private mvarTotalTurboramJetThrustAB As Single
	
	Private mvargSpeedFactor As Integer
	Private mvargAcceleration As Single
	Private mvargDeceleration As Single
	Private mvargManeuverability As Single
	Private mvargOffRoad As Single
	Private mvargPressure As Single
	Private mvargPressureDescription As String
	Private mvargStability As Single
	Private mvargTopSpeed As Single
	
	Private mvarKey As String
	Private mvarParent As String
	Private mvarDatatype As Integer
	Private mvarDescription As String
	Private mvarMotiveAssemblyKey As String
	
	Private mbResponsive As Boolean
	Private mvarTiltRotorForwardThrust As Single
	Private mvarTreatTiltRotorsAsPropellers As Boolean
	Private mvarAfterBurnersOn As Boolean
	Private mvarHardPointsOn As Boolean
	Private mvarWheelsSkidsExtended As Boolean
	Private mvarPopTurretsExtended As Boolean
	
	Private mvarPercentThrust As Single
	Private mvarPercentCrewWeight As Single
	Private mvarPercentFuelWeight As Single
	Private mvarPercentCargoWeight As Single
	Private mvarPercentHardpointWeight As Single
	Private mvarPercentProvisionWeight As Single
	Private mvarPercentAmmunitionWeight As Single
	Private mvarPercentAuxVehicleWeight As Single
	Private m_VWeight As Double
	Private m_VMass As Double
	Private g_sDC As String
	Private mvarZZInit As Byte
	 'make sure the Keychain array starts at 1 and not 0
	
	Private mvarKeyChain As Object
	'JAW 2000.06.12
	'mvarKeyChain is an array of PropulsionKeys representing engine components used in a performance profile.
	
	'JAW 2000.06.15
	'Short message to be displayed if something is wrong with mode of transportation.
	Private mvarAdvisory As String
	
	
	
	Public Property Parent() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Parent
			Parent = mvarParent
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Parent = 5
			mvarParent = Value
		End Set
	End Property
	
	
	
	Public Property Datatype() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.DataType
			Datatype = mvarDatatype
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.DataType = 5
			mvarDatatype = Value
		End Set
	End Property
	
	
	
	Public Property PercentThrust() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wTopSpeed
			PercentThrust = mvarPercentThrust * 100
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wTopSpeed = 5
			If Value > 100 Then Value = 100
			If Value < 0 Then Value = 0
			mvarPercentThrust = Value / 100
			
		End Set
	End Property
	
	
	
	Public Property PercentCrewWeight() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wTopSpeed
			PercentCrewWeight = mvarPercentCrewWeight * 100
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wTopSpeed = 5
			If Value > 100 Then Value = 100
			If Value < 0 Then Value = 0
			mvarPercentCrewWeight = Value / 100
		End Set
	End Property
	
	
	Public Property PercentFuelWeight() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wTopSpeed
			PercentFuelWeight = mvarPercentFuelWeight * 100
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wTopSpeed = 5
			If Value > 100 Then Value = 100
			If Value < 0 Then Value = 0
			mvarPercentFuelWeight = Value / 100
		End Set
	End Property
	
	
	Public Property PercentCargoWeight() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wTopSpeed
			PercentCargoWeight = mvarPercentCargoWeight * 100
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wTopSpeed = 5
			If Value > 100 Then Value = 100
			If Value < 0 Then Value = 0
			mvarPercentCargoWeight = Value / 100
		End Set
	End Property
	
	
	Public Property PercentHardpointWeight() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wTopSpeed
			PercentHardpointWeight = mvarPercentHardpointWeight * 100
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wTopSpeed = 5
			If Value > 100 Then Value = 100
			If Value < 0 Then Value = 0
			mvarPercentHardpointWeight = Value / 100
		End Set
	End Property
	
	
	Public Property PercentProvisionWeight() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wTopSpeed
			PercentProvisionWeight = mvarPercentProvisionWeight * 100
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wTopSpeed = 5
			If Value > 100 Then Value = 100
			If Value < 0 Then Value = 0
			mvarPercentProvisionWeight = Value / 100
		End Set
	End Property
	
	
	Public Property PercentAmmunitionWeight() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wTopSpeed
			PercentAmmunitionWeight = mvarPercentAmmunitionWeight * 100
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wTopSpeed = 5
			If Value > 100 Then Value = 100
			If Value < 0 Then Value = 0
			mvarPercentAmmunitionWeight = Value / 100
		End Set
	End Property
	
	
	
	
	Public Property PercentAuxVehicleWeight() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PercentAuxVehicleWeight
			PercentAuxVehicleWeight = mvarPercentAuxVehicleWeight * 100
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PercentAuxVehicleWeight = 5
			If Value > 100 Then Value = 100
			If Value < 0 Then Value = 0
			mvarPercentAuxVehicleWeight = Value / 100
		End Set
	End Property
	
	Public Property KeyChain() As Object
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.KeyChain
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			KeyChain = mvarKeyChain
		End Get
		Set(ByVal Value As Object)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.KeyChain = 5
			'UPGRADE_WARNING: Couldn't resolve default property of object vdata. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarKeyChain = Value
		End Set
	End Property
	
	
	
	Public Property Key() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Key
			Key = mvarKey
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Key = 5
			mvarKey = Value
		End Set
	End Property
	
	
	
	Public Property MotiveAssemblyKey() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.MotiveAssemblyKey
			MotiveAssemblyKey = mvarMotiveAssemblyKey
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.MotiveAssemblyKey = 5
			mvarMotiveAssemblyKey = Value
		End Set
	End Property
	
	
	
	Public Property Description() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Description
			Description = mvarDescription
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Description = 5
			mvarDescription = Value
		End Set
	End Property
	
	
	
	
	Public Property Advisory() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Advisory
			Advisory = mvarAdvisory
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Advisory = "5"
			mvarAdvisory = Value
		End Set
	End Property
	
	
	
	
	Public Property HardPointsOn() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.HardPointsOn
			HardPointsOn = mvarHardPointsOn
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.HardPointsOn = 5
			mvarHardPointsOn = Value
		End Set
	End Property
	
	
	
	
	
	Public Property WheelsSkidsExtended() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.WheelsSkidsExtended
			WheelsSkidsExtended = mvarWheelsSkidsExtended
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.WheelsSkidsExtended = 5
			mvarWheelsSkidsExtended = Value
		End Set
	End Property
	
	
	
	Public Property PopTurretsExtended() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.HardPointsOn
			PopTurretsExtended = mvarPopTurretsExtended
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PopTurretsExtended = 5
			mvarPopTurretsExtended = Value
		End Set
	End Property
	
	
	Public Property AfterburnersOn() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.AfterburnersOn
			AfterburnersOn = mvarAfterBurnersOn
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.AfterburnersOn = 5
			mvarAfterBurnersOn = Value
		End Set
	End Property
	
	
	
	Public Property TreatTiltRotorsAsPropellers() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.AfterburnersOn
			TreatTiltRotorsAsPropellers = mvarTreatTiltRotorsAsPropellers
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.AfterburnersOn = 5
			mvarTreatTiltRotorsAsPropellers = Value
		End Set
	End Property
	
	
	Public Property TotalRamJetThrust() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.TotalRamJetThrust
			TotalRamJetThrust = mvartotalramjetthrust
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.TotalRamJetThrust = 5
			mvartotalramjetthrust = Value
		End Set
	End Property
	
	
	Public Property TotalTurboRamJetThrust() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.TotalTurboRamJetThrust
			TotalTurboRamJetThrust = mvarTotalTurboRamJetThrust
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.TotalTurboRamJetThrust = 5
			mvarTotalTurboRamJetThrust = Value
		End Set
	End Property
	
	
	Public Property TotalRamJetThrustAB() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.TotalRamJetThrustAB
			TotalRamJetThrustAB = mvarTotalRamJetThrustAB
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.TotalRamJetThrustAB = 5
			mvarTotalRamJetThrustAB = Value
		End Set
	End Property
	
	
	Public Property TotalTurboRamJetThrustAB() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.TotalTurboRamJetThrustAB
			TotalTurboRamJetThrustAB = mvarTotalTurboramJetThrustAB
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.TotalTurboRamJetThrustAB = 5
			mvarTotalTurboramJetThrustAB = Value
		End Set
	End Property
	
	
	
	
	Public Property gSpeedFactor() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.gSpeedFactor
			gSpeedFactor = mvargSpeedFactor
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.gSpeedFactor = 5
			
			mvargSpeedFactor = Value
		End Set
	End Property
	
	
	
	
	
	Public Property gTopSpeed() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.gTopSpeed
			gTopSpeed = mvargTopSpeed
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.gTopSpeed = 5
			mvargTopSpeed = Value
		End Set
	End Property
	
	
	Public Property gStability() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.gStability
			gStability = mvargStability
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.gStability = 5
			mvargStability = Value
		End Set
	End Property
	
	
	
	Public Property gPressure() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.gPressure
			gPressure = mvargPressure
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.gPressure = 5
			mvargPressure = Value
		End Set
	End Property
	
	
	
	
	Public Property gPressureDescription() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.gPressureDescription
			gPressureDescription = mvargPressureDescription
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.gPressureDescription = 5
			mvargPressureDescription = Value
		End Set
	End Property
	
	
	
	Public Property gOffRoad() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.gOffRoad
			gOffRoad = mvargOffRoad
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.gOffRoad = 5
			mvargOffRoad = Value
		End Set
	End Property
	
	
	
	Public Property gManeuverability() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.gManeuverability
			gManeuverability = mvargManeuverability
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.gManeuverability = 5
			mvargManeuverability = Value
		End Set
	End Property
	
	
	
	
	Public Property gDeceleration() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.gDeceleration
			gDeceleration = mvargDeceleration
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.gDeceleration = 5
			mvargDeceleration = Value
		End Set
	End Property
	
	
	
	
	Public Property gAcceleration() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.gAcceleration
			gAcceleration = mvargAcceleration
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.gAcceleration = 5
			mvargAcceleration = Value
		End Set
	End Property
	
	
	
	
	Public Function GetCurrentKeys() As String()
		GetCurrentKeys = VariantArrayToStringArray(mvarKeyChain)
	End Function
	
	Public Sub AddKey(ByRef PropulsionKey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarKeyChain = mAddKey(mvarKeyChain, PropulsionKey)
	End Sub
	
	Public Sub RemoveKey(ByRef PropulsionKey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarKeyChain = mRemoveKey(mvarKeyChain, PropulsionKey)
	End Sub
	
	Public Sub CalcPerformance()
		mvarAdvisory = ""
		
		mvarMotiveAssemblyKey = GetMotiveAssemblyKey(mvarDatatype)
		
		If mvarMotiveAssemblyKey = "" Then
			modHelper.InfoPrint(0, "No skid subassembly detected on vehicle.  Cannot caclulate skid ground performance.")
			Exit Sub
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not TypeOf Veh.Components(mvarMotiveAssemblyKey) Is clsSkid Then
				modHelper.InfoPrint(0, "No skid subassembly detected on vehicle.  Cannot caclulate skid ground performance.")
				Exit Sub
			End If
		End If
		
		
		'determine if vehicle has responsive structure
		mbResponsive = VehicleHasResponsiveStruct
		
		Call GetVehicleWeight(PERFORMANCESKID, mvarPercentAuxVehicleWeight, mvarPercentCargoWeight, mvarPercentAmmunitionWeight, mvarPercentHardpointWeight, mvarPercentFuelWeight, mvarPercentProvisionWeight, m_VWeight, m_VMass)
		
		If mvarTreatTiltRotorsAsPropellers Then
			mvarTiltRotorForwardThrust = GetTiltRotorForwardThrust(mvarPercentThrust)
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcGroundSpeedFactor. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvargSpeedFactor = CalcGroundSpeedFactor 'this one cant be an mvar unless i move
		mvargTopSpeed = CalcGroundSpeed
		mvargAcceleration = CalcGroundAcceleration(mvargSpeedFactor, mvargTopSpeed)
		CalcGGroundDeceleration()
		CalcGSRandMR()
		CalcGPressureandOffRoadSpeed()
		
	End Sub
	
	Function CalcGroundSpeed() As Single
		Dim TempSpeed As Single
		Dim dType As String
		Dim TotalMotivePower As Single
		Dim i As Integer
		Dim sKey As String
		
		On Error GoTo errorhandler
		
		'if there are no propulsion systems on the Keychain, exit the function
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarKeyChain(1) = "" Then Exit Function
		
		
		For i = 1 To UBound(mvarKeyChain)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sKey = mvarKeyChain(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = Veh.Components(sKey).Datatype
			' add motive powers for all other types of thrust components
			Select Case dType
				Case CStr(RopeHarness), CStr(YokeandPoleHarness), CStr(ShaftandCollarHarness), CStr(WhiffletreeHarness)
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TotalMotivePower = TotalMotivePower + Veh.Components(sKey).MotivePower * mvarPercentThrust
				Case CStr(Turbojet), CStr(Turbofan), CStr(Hyperfan), CStr(FusionAirRam)
					'use Afterburner Thrust if this option is enabled by the user
					If mvarAfterBurnersOn Then
						'determine if this engine has afterburners or not
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Veh.Components(sKey).Afterburner Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).ABThrust / 4)
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotiveThrust / 4) * mvarPercentThrust
						End If
					Else 'use normal engine thrust without afterburner
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotiveThrust / 4) * mvarPercentThrust
					End If
				Case CStr(Ramjet)
					'store these values for use later in the TopSpeed calculations since
					'Ramjets only work if Topspeed is at least 375mph
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvartotalramjetthrust = mvartotalramjetthrust + (Veh.Components(sKey).MotiveThrust / 4 * mvarPercentThrust)
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarTotalRamJetThrustAB = mvarTotalRamJetThrustAB + (Veh.Components(sKey).ABThrust / 4)
					
				Case CStr(TurboRamjet)
					'store these values for use later since they add .2 x their thrust
					'if the speed is greater than 375mph
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarTotalTurboRamJetThrust = mvarTotalTurboRamJetThrust + (Veh.Components(sKey).MotiveThrust / 4 * mvarPercentThrust)
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarTotalTurboramJetThrustAB = mvarTotalTurboramJetThrustAB + (Veh.Components(sKey).ABThrust / 4)
					
					'but in the meantime, they just add their default thrust
					'use Afterburner Thrust if this option is enabled by the user
					If mvarAfterBurnersOn Then
						'determine if this engine has afterburners or not
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Veh.Components(sKey).Afterburner Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).ABThrust / 4)
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotiveThrust / 4 * mvarPercentThrust)
						End If
					Else 'use normal engine thrust without afterburner
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotiveThrust / 4 * mvarPercentThrust)
					End If
				Case Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotiveThrust / 4 * mvarPercentThrust)
			End Select
		Next 
		TempSpeed = System.Math.Sqrt(TotalMotivePower / m_VMass)
		TempSpeed = TempSpeed * gSpeedFactor
		
		
		'adjust speed for Wheels/Skids if using Ramjets or TurboRamJets
		
		If TempSpeed >= 375 Then
			If mvarAfterBurnersOn Then
				TotalMotivePower = TotalMotivePower + mvarTotalRamJetThrustAB
				TotalMotivePower = TotalMotivePower + (mvarTotalTurboramJetThrustAB * 0.2)
			Else
				TotalMotivePower = TotalMotivePower + mvartotalramjetthrust
				TotalMotivePower = TotalMotivePower + (mvarTotalTurboRamJetThrust * 0.2)
			End If
			TempSpeed = System.Math.Sqrt(TotalMotivePower / m_VMass)
			TempSpeed = TempSpeed * gSpeedFactor
		End If
		
		
		' add in the streamlining effects
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With Veh.surface
			If TempSpeed > 50 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If .StreamLining = "none" Then
					If TempSpeed > 600 Then TempSpeed = 600
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf .StreamLining = "fair" Then 
					TempSpeed = TempSpeed + (TempSpeed * 0.05)
					If TempSpeed > 600 Then TempSpeed = 600
				Else
					TempSpeed = TempSpeed + (TempSpeed * 0.1)
				End If
			End If
		End With
		
		' Check to make sure we dont exceed speed of slowest animal(if applicable)
		'UPGRADE_WARNING: Couldn't resolve default property of object MinimumNonZero(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempSpeed = MinimumNonZero(TempSpeed, GetSlowestAnimalSpeed(mvarKeyChain))
		
		' Do the final rounding
		If TempSpeed >= 20 Then
			' round to nearest 5mph
			CalcGroundSpeed = System.Math.Round(TempSpeed / 5, 0) * 5
		Else ' round to nearest whole number
			CalcGroundSpeed = System.Math.Round(TempSpeed, 0)
		End If
		
errorhandler: 
		If Err.Number = 9 Then 'subscript out of range check for the Keychain if it hasnt been intialized yet
			Exit Function
		End If
	End Function
	
	Function CalcGroundSpeedFactor() As Object
		Dim Bonus As Short
		Dim ImpvdSuspension As Boolean
		Dim element As Object
		
		
		'jaw 2000.06.18
		'added in to accomodate takeoff and landing calculations
		'If TempKey = "" Then
		'    For Each element In Veh
		'        If MotiveAssembly = element.Description Then
		'            TempKey = element.Key
		'            Exit For
		'        End If
		'    Next
		'End If
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Veh.Components(mvarMotiveAssemblyKey).ImprovedSuspension Then
			ImpvdSuspension = True
		End If
		
		
		'find the improved suspension bonus
		If ImpvdSuspension Then
			Bonus = 1
		End If
		
		' get the Final total speed factor
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcGroundSpeedFactor. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalcGroundSpeedFactor = 6 + Bonus
		
		
	End Function
	
	
	Private Function CalcGroundAcceleration(ByVal Speedfactor As Integer, ByVal TopSpeed As Single) As Single
		Dim TempAcceleration As Integer
		On Error GoTo errorhandler
		
		System.Diagnostics.Debug.Assert(Speedfactor <> 0, "")
		TempAcceleration = (TopSpeed / Speedfactor) * 0.8
		
		' Do the rounding
		If TempAcceleration > 5 Then
			mvargAcceleration = System.Math.Round(TempAcceleration / 5, 0) * 5 'to nearest 5mph
		Else
			mvargAcceleration = System.Math.Round(TempAcceleration, 0) 'to nearest 1mph
		End If
		
		CalcGroundAcceleration = TempAcceleration
		Exit Function
errorhandler: 
		Exit Function
	End Function
	
	Private Sub CalcGGroundDeceleration()
		
		'//////////////////////////////////////////////////////
		'now get the Ground Deceleration
		mvargDeceleration = 5
		
	End Sub
	
	Sub CalcGSRandMR()
		'///////////////////////////////////////////////////////////
		'now get the GroundStability and Manuever Ratings
		Dim MotiveSystem As Short
		Dim TempMR As Single
		Dim TempSR As Single
		Dim TempVolume As Single
		Dim VehicleWornAsHarness As Boolean
		Dim ImpvdSuspension As Boolean
		Dim element As Object 'used to finding exceptions for sails, harnessed animals, etc.
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempVolume = Veh.Components(BODY_KEY).Volume
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Veh.Components(mvarMotiveAssemblyKey).ImprovedSuspension Then
			ImpvdSuspension = True
		End If
		
		
		' Determine which Motive System to use
		'1 = skid
		'2 = Wheels 1
		'3 = Wheels 2
		'4 = Wheels 3
		'5 = wheels 4-7
		'6 = Wheels 8+
		'7 = tracks
		'8 = skitracks
		'9 = halftracks
		'10 = 2 legs
		'11 = 3 legs
		'12 = 4+ legs
		'13 = Flexibody
		
		MotiveSystem = 1
		
		' get the actual values from the table
		If TempVolume <= 30 Then
			TempMR = GroundStabMatrix(MotiveSystem).M1
			TempSR = GroundStabMatrix(MotiveSystem).S1
		ElseIf TempVolume <= 100 Then 
			TempMR = GroundStabMatrix(MotiveSystem).M2
			TempSR = GroundStabMatrix(MotiveSystem).S2
		ElseIf TempVolume <= 300 Then 
			TempMR = GroundStabMatrix(MotiveSystem).M3
			TempSR = GroundStabMatrix(MotiveSystem).S3
		ElseIf TempVolume <= 3000 Then 
			TempMR = GroundStabMatrix(MotiveSystem).M4
			TempSR = GroundStabMatrix(MotiveSystem).S4
		Else
			TempMR = GroundStabMatrix(MotiveSystem).M5
			TempSR = GroundStabMatrix(MotiveSystem).S5
		End If
		'//////////////////////////////////////////////////////
		' apply the gMR and gSR modifiers
		If ImpvdSuspension Then
			TempSR = TempSR + 1 'increase the gSR
			If TempMR = 0.125 Then
				TempMR = 0.25
			Else
				TempMR = TempMR + 0.25
			End If
		End If
		
		'add mods for Electronic or Computer controls
		If VehiclehasElectORCompcontrols Then
			If TempMR = 0.125 Then
				TempMR = 0.25
			Else
				TempMR = TempMR + 0.25
			End If
		End If
		
		'add modifier for responsive structure.
		If mbResponsive Then
			If TempMR = 0.125 Then
				TempMR = 0.25
			Else
				TempMR = TempMR + 0.25
			End If
		End If
		
		' conduct final exception checks for harnessed animals, sails or non-folded wings and Rotors
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsSail Then
				If TempMR > 0.5 Then TempMR = 0.5 'vehicle with sails cant exceed .5 MR
			ElseIf TypeOf element Is clsHarness Then 
				If TempMR > 0.5 Then TempMR = 0.5 'vehicle with harnessed animals cant exceed .5MR
			ElseIf TypeOf element Is clsWing Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Folding. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.Folding <> True Then
					If TempMR > 0.5 Then TempMR = 0.5 'vehicles with non folded wings limited to .5MR
				End If
			ElseIf TypeOf element Is clsRotor Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Folding. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.Folding <> True Then
					If TempMR > 0.5 Then TempMR = 0.5 'vehicles with non folded Rotors limited to .5MR
				End If
			ElseIf TypeOf element Is clsCrewStation Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.Datatype = HarnessCrewStation Then
					VehicleWornAsHarness = True
				End If
			End If
		Next element
		'vehicles worn as harness are SR - 1
		If VehicleWornAsHarness Then
			TempSR = TempSR - 1
		End If
		
		'Save the final gMR and gSR results
		gManeuverability = TempMR
		gStability = TempSR
		
	End Sub
	
	Sub CalcGPressureandOffRoadSpeed()
		'////////////////////////////////////////////////////////////
		'' get the contact area
		Dim element As Object
		Dim ContactArea As Single
		Dim TempContactArea As Single
		Dim i As Integer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarKeyChain(1) = "" Then 'exit sub if the keychain does not contain any propulsion systems
			mvargPressureDescription = ""
			mvargOffRoad = 0
			mvargPressure = 0
			Exit Sub
		End If
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ContactArea = Veh.Components(mvarMotiveAssemblyKey).SurfaceArea / 10
		
		
		' get the ground pressure
		Dim TempWeight As Single
		Dim MinWeight As Single ' contragrav cant reduce weight to less than 10% of original
		Dim sDescription As String
		Dim TempPressure As Single
		Dim First As Short
		'UPGRADE_NOTE: Second was upgraded to Second_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Second_Renamed As Short
		Dim arrPT(7, 28) As Object
		'fill the pressure table
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(1, 1) = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(1, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(1, 2) = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(1, 3). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(1, 3) = 4 / 5
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(1, 4). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(1, 4) = 2 / 3
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(2, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(2, 1) = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(2, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(2, 2) = 4 / 5
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(2, 3). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(2, 3) = 2 / 3
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(2, 4). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(2, 4) = 0.5
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(3, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(3, 1) = 4 / 5
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(3, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(3, 2) = 2 / 3
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(3, 3). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(3, 3) = 0.5
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(3, 4). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(3, 4) = 1 / 3
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(4, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(4, 1) = 2 / 3
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(4, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(4, 2) = 0.5
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(4, 3). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(4, 3) = 1 / 3
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(4, 4). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(4, 4) = 1 / 4
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(5, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(5, 1) = 0.5
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(5, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(5, 2) = 1 / 3
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(5, 3). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(5, 3) = 1 / 4
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(5, 4). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(5, 4) = 1 / 6
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(6, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(6, 1) = 1 / 3
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(6, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(6, 2) = 1 / 4
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(6, 3). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(6, 3) = 1 / 6
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(6, 4). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(6, 4) = 1 / 8
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(7, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(7, 1) = 1 / 4
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(7, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(7, 2) = 1 / 6
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(7, 3). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(7, 3) = 1 / 8
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(7, 4). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrPT(7, 4) = 0
		
		TempWeight = m_VWeight
		MinWeight = TempWeight * 0.1
		TempWeight = TempWeight - CalcTotalContragravLift
		If TempWeight < MinWeight Then TempWeight = MinWeight 'make sure contragrav reduction leaves at least 10% of loaded weight
		
		If ContactArea = 0 Then
			TempWeight = 0
		Else
			TempWeight = TempWeight / ContactArea 'this gives us our actual ground pressure
		End If
		mvargPressure = TempWeight
		
		If TempWeight <= 150 Then
			First = 1
			mvargPressureDescription = "extremely low"
		ElseIf TempWeight <= 900 Then 
			First = 2
			mvargPressureDescription = "very low"
		ElseIf TempWeight <= 1800 Then 
			First = 3
			mvargPressureDescription = "low"
		ElseIf TempWeight <= 2700 Then 
			First = 4
			mvargPressureDescription = "moderate"
		ElseIf TempWeight <= 7500 Then 
			First = 5
			mvargPressureDescription = "high"
		ElseIf TempWeight <= 15000 Then 
			First = 6
			mvargPressureDescription = "very high"
		Else
			First = 7
			mvargPressureDescription = "extremely high"
		End If
		
		
		Second_Renamed = 4
		' get off road speed
		'UPGRADE_WARNING: Couldn't resolve default property of object arrPT(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvargOffRoad = mvargTopSpeed * arrPT(First, Second_Renamed)
		
		
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		ReDim mvarKeyChain(1)
		mvarDatatype = PERFORMANCESKID
		
		mvarAfterBurnersOn = True
		'Private mvarHardPointsOn As Boolean
		'JAW 2000.06.20
		'value of 1 means 100%
		mvarPercentThrust = 1
		mvarPercentCrewWeight = 1
		mvarPercentFuelWeight = 1
		mvarPercentCargoWeight = 1
		mvarPercentHardpointWeight = 1
		mvarPercentProvisionWeight = 1
		mvarPercentAmmunitionWeight = 1
		mvarPercentAuxVehicleWeight = 1
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class