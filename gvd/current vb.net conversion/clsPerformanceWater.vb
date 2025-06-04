Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsPerformanceWater_NET.clsPerformanceWater")> Public Class clsPerformanceWater
	
	Private mvartotalramjetthrust As Single
	Private mvarTotalTurboRamJetThrust As Single
	Private mvarTotalRamJetThrustAB As Single
	Private mvarTotalTurboramJetThrustAB As Single
	
	Private mvarwAcceleration As Single
	Private mvarwDeceleration As Single
	Private mvarwIDeceleration As Single 'the increased deceleration
	Private mvarwDraft As Single
	Private mvarwHydroDrag As Single
	Private mvarwManeuverability As Single
	Private mvarwStability As Single
	Private mvarwTopSpeed As Single
	Private mvarwPlaningSpeed As Single 'note: planing speed is in addition to topspeed (see page 131 top right column)
	Private mvarwHydrofoilSpeed As Single 'note:hydrofoil speed is in addition to topspeed(see page 131 top right column
	Private mvarwTotalAquaticThrust As Single
	
	Private mvarKey As String
	Private mvarParent As String
	Private mvarDatatype As Integer
	Private mvarDescription As String
	
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
	
	
	
	
	Public Property wTopSpeed() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wTopSpeed
			wTopSpeed = mvarwTopSpeed
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wTopSpeed = 5
			mvarwTopSpeed = Value
		End Set
	End Property
	
	
	
	Public Property wHydrofoilSpeed() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wHydrofoilSpeed
			wHydrofoilSpeed = mvarwHydrofoilSpeed
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wHydrofoilSpeed = 5
			mvarwHydrofoilSpeed = Value
		End Set
	End Property
	
	
	
	Public Property wPlaningSpeed() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wPlaningSpeed
			wPlaningSpeed = mvarwPlaningSpeed
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wPlaningSpeed = 5
			mvarwPlaningSpeed = Value
		End Set
	End Property
	
	
	
	Public Property wStability() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wStability
			wStability = mvarwStability
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wStability = 5
			mvarwStability = Value
		End Set
	End Property
	
	
	
	Public Property wManeuverability() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wManeuverability
			wManeuverability = mvarwManeuverability
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wManeuverability = 5
			mvarwManeuverability = Value
		End Set
	End Property
	
	
	Public Property wHydroDrag() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wHydroDrag
			wHydroDrag = mvarwHydroDrag
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wHydroDrag = 5
			mvarwHydroDrag = Value
		End Set
	End Property
	
	
	Public Property wDraft() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wDraft
			wDraft = mvarwDraft
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wDraft = 5
			mvarwDraft = Value
		End Set
	End Property
	
	
	Public Property wDeceleration() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wDeceleration
			wDeceleration = mvarwDeceleration
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wDeceleration = 5
			mvarwDeceleration = Value
		End Set
	End Property
	
	Public Property wIDeceleration() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wIDeceleration
			wIDeceleration = mvarwIDeceleration
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wIDeceleration = 5
			mvarwIDeceleration = Value
		End Set
	End Property
	
	Public Property wAcceleration() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wAcceleration
			wAcceleration = mvarwAcceleration
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wAcceleration = 5
			mvarwAcceleration = Value
		End Set
	End Property
	
	
	Public Property wTotalAquaticThrust() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.wTotalAquaticThrust
			wTotalAquaticThrust = mvarwTotalAquaticThrust
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.wTotalAquaticThrust = 5
			mvarwTotalAquaticThrust = Value
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
		
		'determine if vehicle has responsive structure
		mbResponsive = VehicleHasResponsiveStruct
		
		Call GetVehicleWeight(PERFORMANCEWATER, mvarPercentAuxVehicleWeight, mvarPercentCargoWeight, mvarPercentAmmunitionWeight, mvarPercentHardpointWeight, mvarPercentFuelWeight, mvarPercentProvisionWeight, m_VWeight, m_VMass)
		
		If mvarTreatTiltRotorsAsPropellers Then
			mvarTiltRotorForwardThrust = GetTiltRotorForwardThrust(mvarPercentThrust)
		End If
		
		
		mvarwHydroDrag = CalcHydroDrag
		mvarwTopSpeed = CalcWaterSpeed
		mvarwAcceleration = CalcWaterAcceleration(mvarwTotalAquaticThrust, m_VWeight)
		Call CalcWaterMRandSR(mvarwStability, mvarwManeuverability, mvarKeyChain, mbResponsive)
		Call CalcWaterDeceleration(mvarwManeuverability, mvarwAcceleration, mvarwDeceleration, mvarwIDeceleration) 'this does both Deceleration and Increased Deceleration
		mvarwDraft = CalcDraft
		
	End Sub
	
	Function CalcHydroDrag() As Single
		Dim Trimaran As Object
		Dim Catamaran As Object
		Dim Hl As Short
		Dim Templift As Single
		Dim MinWeight As Single
		Dim TempWeight As Single
		Dim TempDrag As Single
		
		Templift = CalcTotalContragravLift
		Hl = GetHl
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With Veh.surface
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .CataTrimaran = Trimaran Then 'todo: need constant for this
				Hl = Hl + (Hl * 0.1)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf .CataTrimaran = Catamaran Then 
				Hl = Hl + (Hl * 0.2)
			End If
			TempWeight = m_VWeight
		End With
		
		MinWeight = TempWeight * 0.1
		TempWeight = TempWeight - Templift
		If TempWeight < MinWeight Then TempWeight = MinWeight
		TempDrag = ((TempWeight ^ (1 / 3)) ^ 2) / Hl
		CalcHydroDrag = System.Math.Round(TempDrag, 0)
		
	End Function
	
	
	Function CalcWaterSpeed() As Single
		Dim TempSpeed As Single
		Dim dType As String
		Dim sKey As String
		Dim i As Integer
		Dim TotalMotivePower As Single
		On Error Resume Next
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarKeyChain(1) = "" Then Exit Function 'if no propulsion systems, exit the function
		
		For i = 1 To UBound(mvarKeyChain)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sKey = mvarKeyChain(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = Veh.Components(sKey).Datatype
			' add motive powers for all other types of thrust components
			Select Case dType
				Case CStr(TrackedDrivetrain), CStr(LegDrivetrain), CStr(WheeledDrivetrain), CStr(AllWheelDriveWheeledDrivetrain)
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotivePower * 2 * mvarPercentThrust)
				Case CStr(FlexibodyDrivetrain)
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotivePower * 5 * mvarPercentThrust)
				Case "" ' debug Need to check if putting this hack in is ok
					'do nothing
				Case CStr(RowingPositions), CStr(PaddleWheel), CStr(ScrewPropeller), CStr(lightScrewPropeller), CStr(DuctedPropeller), CStr(Hydrojet), CStr(MHDTunnel), CStr(DuctedFan), CStr(AerialPropeller), CStr(RopeHarness), CStr(YokeandPoleHarness), CStr(ShaftandCollarHarness), CStr(WhiffletreeHarness), CStr(ForeandAftRig), CStr(SquareRig), CStr(FullRig), CStr(AerialSail), CStr(AerialSailForeAftRig), CStr(LiquidFuelRocket), CStr(MOXRocket), CStr(IonDrive), CStr(FissionRocket), CStr(FusionRocket), CStr(OptimizedFusion), CStr(AntimatterThermal), CStr(AntimatterPion), CStr(StandardThruster), CStr(SuperThruster), CStr(MegaThruster), CStr(SolidRocketEngine)
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TotalMotivePower = TotalMotivePower + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
				Case CStr(Turbojet), CStr(Turbofan), CStr(Hyperfan), CStr(FusionAirRam)
					'use Afterburner Thrust if this option is enabled by the user
					If mvarAfterBurnersOn Then
						'determine if this engine has afterburners or not
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Veh.Components(sKey).Afterburner Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TotalMotivePower = TotalMotivePower + Veh.Components(sKey).ABThrust
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TotalMotivePower = TotalMotivePower + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
						End If
					Else 'use normal engine thrust without afterburner
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						TotalMotivePower = TotalMotivePower + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
					End If
				Case CStr(Ramjet)
					'store these values for use later in the TopSpeed calculations since
					'Ramjets only work if Topspeed is at least 375mph
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvartotalramjetthrust = mvartotalramjetthrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarTotalRamJetThrustAB = mvarTotalRamJetThrustAB + Veh.Components(sKey).ABThrust
					
				Case CStr(TurboRamjet)
					'store these values for use later since they add .2 x their thrust
					'if the speed is greater than 375mph
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarTotalTurboRamJetThrust = mvarTotalTurboRamJetThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarTotalTurboramJetThrustAB = mvarTotalTurboramJetThrustAB + Veh.Components(sKey).ABThrust
					
					'but in the meantime, they just add their default thrust
					'use Afterburner Thrust if this option is enabled by the user
					If mvarAfterBurnersOn Then
						'determine if this engine has afterburners or not
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Veh.Components(sKey).Afterburner Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TotalMotivePower = TotalMotivePower + Veh.Components(sKey).ABThrust
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TotalMotivePower = TotalMotivePower + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
						End If
					Else 'use normal engine thrust without afterburner
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						TotalMotivePower = TotalMotivePower + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
					End If
			End Select
		Next 
		'//add in the title rotor forward thrust
		TotalMotivePower = TotalMotivePower + mvarTiltRotorForwardThrust
		
		wTotalAquaticThrust = TotalMotivePower 'save the totalaquaticthrust
		If mvarwHydroDrag = 0 Then
			TempSpeed = 0 '//can we have a speed of 0 here?
		Else
			TempSpeed = ((TotalMotivePower / mvarwHydroDrag) ^ (1 / 3)) * 6
		End If
		' add in the streamlining effects
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With Veh.surface
			If TempSpeed > 50 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If .StreamLining = "none" Then
					If TempSpeed > 150 Then TempSpeed = 150
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf .StreamLining = "fair" Then 
					TempSpeed = TempSpeed + (TempSpeed * 0.05)
					If TempSpeed > 150 Then TempSpeed = 150
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf .StreamLining = "good" Then 
					TempSpeed = TempSpeed + (TempSpeed * 0.1)
					If TempSpeed > 150 Then TempSpeed = 150
				Else
					TempSpeed = TempSpeed + (TempSpeed * 0.1)
				End If
			End If
		End With
		
		'adjust speed if using Ramjets or TurboRamJets
		If TempSpeed >= 375 Then
			If mvarAfterBurnersOn Then
				TotalMotivePower = TotalMotivePower + mvarTotalRamJetThrustAB
				TotalMotivePower = TotalMotivePower + (mvarTotalTurboramJetThrustAB * 0.2)
			Else
				TotalMotivePower = TotalMotivePower + mvartotalramjetthrust
				TotalMotivePower = TotalMotivePower + (mvarTotalTurboRamJetThrust * 0.2)
			End If
			TempSpeed = ((TotalMotivePower / wHydroDrag) ^ (1 / 3)) * 6 'Note: do i need to readd streamlining mods here?
		End If
		
		' Check to make sure we dont exceed speed of slowest animal(if applicable)
		'UPGRADE_WARNING: Couldn't resolve default property of object MinimumNonZero(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempSpeed = MinimumNonZero(TempSpeed, GetSlowestAnimalSpeed(mvarKeyChain))
		
		' Do the final rounding
		If TempSpeed >= 20 Then
			' round to nearest 5mph
			TempSpeed = System.Math.Round(TempSpeed / 5, 0) * 5
		Else ' round to nearest whole number
			TempSpeed = System.Math.Round(TempSpeed, 0)
		End If
		
		'TODO: check if hydrofoil and planing top speed calculations are done BEFORE or AFTER
		' Streamlining effects adjustments below (see page 131 top right column) If they are "before" then
		' i need to move Planing and Hydrofoil speed calcs above the Streamlinging effects section
		'TODO: check if the errata "hydrofoil modifier" it speaks of is indeed PlaningSpeed * 1.5
		
		'Get the Planing and Hydrofoil Top Speed
		Dim NeededHydrofoilSpeed As Single
		Dim NeededPlaningSpeed As Single
		Dim TempPlaningSpeed As Single
		Dim TempHydrofoilSpeed As Single
		Dim element As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NeededHydrofoilSpeed = 20 + (Veh.Stats / 100)
		NeededPlaningSpeed = ((GetHl * 5) + 5)
		If mvarwTotalAquaticThrust >= NeededPlaningSpeed / 100 * m_VWeight Then
			TempPlaningSpeed = TempSpeed * 2
		End If
		For	Each element In Veh
			If TypeOf element Is clsHydrofoil Then
				If TempPlaningSpeed > 0 Then
					TempHydrofoilSpeed = TempPlaningSpeed * 1.5
				Else
					TempHydrofoilSpeed = TempSpeed * 1.5
				End If
				Exit For 'exit this loop after we find a hydrofoil
			End If
		Next element
		wPlaningSpeed = TempPlaningSpeed 'save the calculations
		wHydrofoilSpeed = TempHydrofoilSpeed 'save the calculations
		'return the function's value
		CalcWaterSpeed = TempSpeed
	End Function
	
	
	
	
	
	Function CalcDraft() As Single
		
		Dim Hl As Single
		Dim TempDraft As Single
		Dim TempWeight As Single
		Dim sLines As String
		
		TempWeight = m_VWeight
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sLines = Veh.surface.HydrodynamicLines
		
		Select Case sLines
			Case "none"
				Hl = 1
			Case "submarine"
				Hl = 2
			Case "mediocre"
				Hl = 1.1
			Case "average"
				Hl = 1.2
			Case "fine"
				Hl = 1.3
			Case "very fine"
				Hl = 1.4
		End Select
		
		TempDraft = ((TempWeight ^ (1 / 3)) / 15) * Hl
		CalcDraft = System.Math.Round(TempDraft, 1) 'round to one decimal place
		
	End Function
	
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		ReDim mvarKeyChain(1)
		
		
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