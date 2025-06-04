Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsPerformanceMagLev_NET.clsPerformanceMagLev")> Public Class clsPerformanceMagLev
	'local variable(s) to hold property value(s)
	Private mvartotalramjetthrust As Single
	Private mvarTotalTurboRamJetThrust As Single
	Private mvarTotalRamJetThrustAB As Single
	Private mvarTotalTurboramJetThrustAB As Single
	
	Private mvarmlStallSpeed As Single
	Private mvarmlMotiveThrust As Single
	Private mvarmlDrag As Single
	Private mvarmlTopSpeed As Single
	Private mvarmlAcceleration As Single
	Private mvarmlManeuverability As Single
	Private mvarmlStability As Single
	Private mvarmlDeceleration As Single
	
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
	
	Public ReadOnly Property DesignCheckString() As String
		Get
			DesignCheckString = g_sDC
		End Get
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
	
	
	
	Public Property mlManeuverability() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.mlManeuverability
			mlManeuverability = mvarmlManeuverability
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.mlManeuverability = 5
			mvarmlManeuverability = Value
		End Set
	End Property
	
	
	
	
	
	Public Property mlDrag() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.mlDrag
			mlDrag = mvarmlDrag
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.mlDrag = 5
			mvarmlDrag = Value
		End Set
	End Property
	
	
	
	
	
	Public Property mlDeceleration() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.mlDeceleration
			mlDeceleration = mvarmlDeceleration
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.mlDeceleration = 5
			mvarmlDeceleration = Value
		End Set
	End Property
	
	
	
	
	
	Public Property mlAcceleration() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.mlAcceleration
			mlAcceleration = mvarmlAcceleration
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.mlAcceleration = 5
			mvarmlAcceleration = Value
		End Set
	End Property
	
	
	
	
	
	Public Property mlTopSpeed() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.mlTopSpeed
			mlTopSpeed = mvarmlTopSpeed
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.mlTopSpeed = 5
			mvarmlTopSpeed = Value
		End Set
	End Property
	
	
	
	
	
	Public Property mlStallSpeed() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.mlStallSpeed
			mlStallSpeed = mvarmlStallSpeed
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.mlStallSpeed = 5
			mvarmlStallSpeed = Value
		End Set
	End Property
	
	
	
	
	
	Public Property mlStability() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.mlStability
			mlStability = mvarmlStability
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.mlStability = 5
			mvarmlStability = Value
		End Set
	End Property
	
	
	
	
	
	Public Property mlMotiveThrust() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.mlMotiveThrust
			mlMotiveThrust = mvarmlMotiveThrust
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.mlMotiveThrust = 5
			mvarmlMotiveThrust = Value
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
		
		Dim sPType As String
		mvarAdvisory = ""
		sPType = CStr(mvarDatatype)
		
		'determine if vehicle has responsive structure
		mbResponsive = VehicleHasResponsiveStruct
		
		Call GetVehicleWeight(PERFORMANCEMAGLEV, mvarPercentAuxVehicleWeight, mvarPercentCargoWeight, mvarPercentAmmunitionWeight, mvarPercentHardpointWeight, mvarPercentFuelWeight, mvarPercentProvisionWeight, m_VWeight, m_VMass)
		
		If mvarTreatTiltRotorsAsPropellers Then
			mvarTiltRotorForwardThrust = GetTiltRotorForwardThrust(mvarPercentThrust)
		End If
		
		
		'mvarStaticLift = CalcTotalStaticLift 'todo: needed?
		mvarmlStallSpeed = 0
		mvarmlMotiveThrust = CalcMLMotiveThrust
		mvarmlDrag = CalcADrag(mvarPopTurretsExtended, mvarWheelsSkidsExtended, mbResponsive) 'uses same as air performance
		
		mvarmlTopSpeed = CalcATopSpeed(mvarmlDrag, mvarmlMotiveThrust, mvarAfterBurnersOn, mvartotalramjetthrust, mvarTotalTurboRamJetThrust, mvarTotalRamJetThrustAB, mvarTotalTurboramJetThrustAB) 'uses same as air
		
		'check for Max speed limits
		mvarmlTopSpeed = CalcAMaxSpeed(mvarmlTopSpeed, mvarKeyChain, mvarTreatTiltRotorsAsPropellers, g_sDC)
		mvarmlTopSpeed = System.Math.Round(mvarmlTopSpeed / 5, 0) * 5 'round to nearest 5mph
		
		mvarmlAcceleration = CalcAAcceleration(mvarmlMotiveThrust, m_VWeight)
		mvarmlManeuverability = CalcMLMR
		mvarmlStability = 6 'simple fixed rule
		mvarmlDeceleration = 40 * mvarmlManeuverability 'simple rule
		
		
	End Sub
	Function CalcMLMotiveThrust() As Single
		Dim i As Short
		Dim TempThrust As Single
		Dim sKey As String
		Dim dType As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarKeyChain(1) = "" Then Exit Function 'exit if there are no propulsion systems in the keychain
		
		For i = 1 To UBound(mvarKeyChain)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sKey = mvarKeyChain(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = Veh.Components(sKey).Datatype
			
			' add motive powers for all other types of thrust components
			Select Case dType
				Case MagLevLifter
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempThrust = TempThrust + Veh.Components(sKey).Lift * 0.2
				Case AerialPropeller, DuctedFan, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
					
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempThrust = TempThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
					
				Case Turbojet, Turbofan, Hyperfan, FusionAirRam
					'use Afterburner Thrust if this option is enabled by the user
					If mvarAfterBurnersOn Then
						'determine if this engine has afterburners or not
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Veh.Components(sKey).Afterburner Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TempThrust = TempThrust + Veh.Components(sKey).ABThrust 'PPP This should only include AB if its checked
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TempThrust = TempThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
						End If
					Else 'use normal engine thrust without afterburner
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						TempThrust = TempThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
					End If
				Case Ramjet
					'store these values for use later in the TopSpeed calculations since
					'Ramjets only work if Topspeed is at least 375mph
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvartotalramjetthrust = mvartotalramjetthrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarTotalRamJetThrustAB = mvarTotalRamJetThrustAB + Veh.Components(sKey).ABThrust
					
				Case TurboRamjet
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
							TempThrust = TempThrust + Veh.Components(sKey).ABThrust
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TempThrust = TempThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
						End If
					Else 'use normal engine thrust without afterburner
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						TempThrust = TempThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
					End If
			End Select
		Next 
		
		'//add in any forward thrust from tilt rotors
		TempThrust = TempThrust + mvarTiltRotorForwardThrust
		
		
		CalcMLMotiveThrust = TempThrust
	End Function
	
	Function CalcMLMR() As Single
		Dim TempMR As Single
		
		If gVehicleTL <= 7 Then
			CalcMLMR = 0.125
		Else : CalcMLMR = 0.25
		End If
	End Function
	
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		ReDim mvarKeyChain(1)
		mvarDatatype = PERFORMANCEMAGLEV
		
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