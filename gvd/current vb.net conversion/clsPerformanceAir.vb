Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsPerformanceAir_NET.clsPerformanceAir")> Public Class clsPerformanceAir
	'local variable(s) to hold property value(s)
	Private mvartotalramjetthrust As Single
	Private mvarTotalTurboRamJetThrust As Single
	Private mvarTotalRamJetThrustAB As Single
	Private mvarTotalTurboramJetThrustAB As Single
	
	
	Private mvarStaticLift As Single
	Private mvaraCanFly As Boolean
	Private mvaraTakeOffRun As Single
	Private mvaraLandingRun As Single
	Private mvaraAcceleration As Single
	Private mvaraDeceleration As Single
	Private mvaraDrag As Single
	Private mvaraManeuverability As Single
	Private mvaraMotiveThrust As Single
	Private mvaraStability As Single
	Private mvaraStallSpeed As Single
	Private mvaraTopSpeed As Single
	Private mvaraReservedVectoredThrust As Single
	Private mvarCeiling As Single
	
	Private mvarhAcceleration As Single
	Private mvarhDeceleration As Single
	Private mvarhSEVSidewalls As Boolean
	Private mvarhGEVSkirt As Boolean
	Private mvarhHoverAltitude As Single
	Private mvarhMotiveThrust As Single
	Private mvarhDrag As Single
	Private mvarhTopSpeed As Single
	Private mvarhManeuverability As Single
	Private mvarhStability As Single
	Private mvarhReservedHoverThrust As Single '//this doesnt actually need a Get/Let since it doesnt need to be saved in .VEH
	
	
	Private mvarKey As String
	Private mvarParent As String
	Private mvarDescription As String
	Private mvarDatatype As Integer
	
	
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
	
	
	
	
	
	Public Property aCanFly() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.aCanFly
			aCanFly = mvaraCanFly
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.aCanFly = 5
			mvaraCanFly = Value
		End Set
	End Property
	
	
	
	Public Property StaticLift() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.StaticLift
			StaticLift = mvarStaticLift
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.StaticLift = 5
			mvarStaticLift = Value
		End Set
	End Property
	
	
	
	Public Property aLandingRun() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.aLandingRun
			aLandingRun = mvaraLandingRun
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.aLandingRun = 5
			mvaraLandingRun = Value
		End Set
	End Property
	
	
	
	Public Property aTakeOffRun() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.aTakeOffRun
			aTakeOffRun = mvaraTakeOffRun
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.aTakeOffRun = 5
			mvaraTakeOffRun = Value
		End Set
	End Property
	
	
	
	
	
	
	Public Property aTopSpeed() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.aTopSpeed
			aTopSpeed = mvaraTopSpeed
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.aTopSpeed = 5
			mvaraTopSpeed = Value
		End Set
	End Property
	
	
	
	
	
	Public Property aStallSpeed() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.aStallSpeed
			aStallSpeed = mvaraStallSpeed
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.aStallSpeed = 5
			mvaraStallSpeed = Value
		End Set
	End Property
	
	
	
	
	
	Public Property aStability() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.aStability
			aStability = mvaraStability
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.aStability = 5
			mvaraStability = Value
		End Set
	End Property
	
	
	
	
	
	Public Property aMotiveThrust() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.aMotiveThrust
			aMotiveThrust = mvaraMotiveThrust
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.aMotiveThrust = 5
			mvaraMotiveThrust = Value
		End Set
	End Property
	
	
	
	
	
	Public Property aManeuverability() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.aManeuverability
			aManeuverability = mvaraManeuverability
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.aManeuverability = 5
			mvaraManeuverability = Value
		End Set
	End Property
	
	
	
	
	
	Public Property aDrag() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.aDrag
			aDrag = mvaraDrag
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.aDrag = 5
			mvaraDrag = Value
		End Set
	End Property
	
	
	
	
	
	Public Property aDeceleration() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.aDeceleration
			aDeceleration = mvaraDeceleration
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.aDeceleration = 5
			mvaraDeceleration = Value
		End Set
	End Property
	
	
	
	
	
	Public Property aAcceleration() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.aAcceleration
			aAcceleration = mvaraAcceleration
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.aAcceleration = 5
			mvaraAcceleration = Value
		End Set
	End Property
	
	
	
	Public Property hstability() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.hstability
			hstability = mvarhStability
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.hstability  = 5
			mvarhStability = Value
		End Set
	End Property
	
	
	
	
	Public Property hmaneuverability() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.hmaneuverability
			hmaneuverability = mvarhManeuverability
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.hmaneuverability  = 5
			mvarhManeuverability = Value
		End Set
	End Property
	
	
	
	Public Property hTopSpeed() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.hTopSpeed
			hTopSpeed = mvarhTopSpeed
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.hTopSpeed  = 5
			mvarhTopSpeed = Value
		End Set
	End Property
	
	
	
	
	
	Public Property hDrag() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.hDrag
			hDrag = mvarhDrag
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.hDrag  = 5
			mvarhDrag = Value
		End Set
	End Property
	
	
	
	
	Public Property hMotiveThrust() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.hMotiveThrust
			hMotiveThrust = mvarhMotiveThrust
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.hMotiveThrust  = 5
			mvarhMotiveThrust = Value
		End Set
	End Property
	
	
	
	
	Public Property hHoverAltitude() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.hHoverAltitude
			hHoverAltitude = mvarhHoverAltitude
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.hHoverAltitude = 5
			mvarhHoverAltitude = Value
		End Set
	End Property
	
	
	
	
	Public Property hSEVSidewalls() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.hSEVSidewalls
			hSEVSidewalls = mvarhSEVSidewalls
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.hSEVSidewalls = 5
			mvarhSEVSidewalls = Value
		End Set
	End Property
	
	
	
	
	Public Property hGEVSkirt() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.hGEVSkirt
			hGEVSkirt = mvarhGEVSkirt
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.hGEVSkirt = 5
			mvarhGEVSkirt = Value
		End Set
	End Property
	
	
	
	Public Property hDeceleration() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.hDeceleration
			hDeceleration = mvarhDeceleration
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.hDeceleration = 5
			mvarhDeceleration = Value
		End Set
	End Property
	
	
	Public Property hAcceleration() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.hAcceleration
			hAcceleration = mvarhAcceleration
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.hAcceleration = 5
			mvarhAcceleration = Value
		End Set
	End Property
	
	
	
	
	Public ReadOnly Property DesignCheckString() As String
		Get
			DesignCheckString = g_sDC
		End Get
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
		Dim sHovType As String
		
		mvarAdvisory = ""
		
		'determine if vehicle has responsive structure
		mbResponsive = VehicleHasResponsiveStruct
		
		Call GetVehicleWeight(PERFORMANCEAIR, mvarPercentAuxVehicleWeight, mvarPercentCargoWeight, mvarPercentAmmunitionWeight, mvarPercentHardpointWeight, mvarPercentFuelWeight, mvarPercentProvisionWeight, m_VWeight, m_VMass)
		
		If mvarTreatTiltRotorsAsPropellers Then
			mvarTiltRotorForwardThrust = GetTiltRotorForwardThrust(mvarPercentThrust)
		End If
		
		
		
		'aerial performance
		mvarStaticLift = CalcTotalStaticLift(mvarKeyChain, mvarTreatTiltRotorsAsPropellers, mvarhSEVSidewalls, mvarhGEVSkirt, PERFORMANCEAIR, mvarPercentThrust, m_VWeight)
		
		
		
		'this must be done first
		mvaraStallSpeed = CalcAStallSpeed
		'    mvargSpeedFactor = CalcGroundSpeedFactor 'this one cant be an mvar unless i move
		'    mvargTopSpeed = CalcGroundSpeed
		'    mvargAcceleration = CalcGroundAcceleration(mvargSpeedFactor, mvargTopSpeed)
		'    mvaraTakeOffRun = CalcTakeOffRun
		'    CalcGGroundDeceleration
		'mvaraLandingRun = CalcLandingRun
		' mvaraCanFly = CalcACanFly
		mvaraMotiveThrust = CalcAMotiveThrust(mvarKeyChain, mvaraReservedVectoredThrust, mvarTreatTiltRotorsAsPropellers, mvarTiltRotorForwardThrust, mvarPercentThrust, mvarAfterBurnersOn, mvarDatatype, mvartotalramjetthrust, mvarTotalRamJetThrustAB, mvarTotalTurboRamJetThrust, mvarTotalTurboramJetThrustAB)
		
		mvaraDrag = CalcADrag(mvarPopTurretsExtended, mvarWheelsSkidsExtended, mbResponsive)
		
		mvaraTopSpeed = CalcATopSpeed(mvaraDrag, mvaraMotiveThrust, mvarAfterBurnersOn, mvartotalramjetthrust, mvarTotalTurboRamJetThrust, mvarTotalRamJetThrustAB, mvarTotalTurboramJetThrustAB)
		
		'check for Max speed limits
		mvaraTopSpeed = CalcAMaxSpeed(mvaraTopSpeed, mvarKeyChain, mvarTreatTiltRotorsAsPropellers, g_sDC)
		
		' round to nearest 5
		mvaraTopSpeed = System.Math.Round(mvaraTopSpeed / 5, 0) * 5 'round to nearest 5mph
		
		mvaraAcceleration = CalcAAcceleration(mvaraMotiveThrust, m_VWeight)
		mvaraManeuverability = CalcAManeuverability(mvaraStallSpeed, mbResponsive, m_VWeight)
		mvaraStability = CalcAStability
		mvaraDeceleration = CalcADeceleration(mvaraManeuverability)
		
		
	End Sub
	
	
	
	'Function CalcTakeOffRun() As Single
	'JAW 2000.06.16
	
	'On Error Resume Next
	'    If mvargAcceleration < 0.1 Then
	'        CalcTakeOffRun = (mvaraStallSpeed ^ 2) / (4 * 0.1)
	'    Else
	'        CalcTakeOffRun = (mvaraStallSpeed ^ 2) / (4 * mvargAcceleration)
	'    End If
	'    If CalcTakeOffRun > Int(CalcTakeOffRun) Then CalcTakeOffRun = Int(CalcTakeOffRun) + 1
	'
	'End Function
	
	'Function CalcLandingRun() As Single
	''JAW 2000.06.16
	
	'On Error Resume Next
	'    'guard against div by 0
	'    If mvargDeceleration < 0.1 Then
	'        CalcLandingRun = (mvaraStallSpeed ^ 2) / (4 * 0.1)
	'    Else
	'        CalcLandingRun = (mvaraStallSpeed ^ 2) / (4 * mvargDeceleration)
	'    End If
	'    If CalcLandingRun > Int(CalcLandingRun) Then CalcLandingRun = Int(CalcLandingRun) + 1
	'End Function
	
	'Function CalcACanFly() As Boolean
	'JAW 2000.06.16
	'On Error Resume Next
	'    If mvaraStallSpeed <= mvargTopSpeed Then
	'        CalcACanFly = True
	'    Else
	'        CalcACanFly = False
	'        Advisory = Advisory & "CANT TAKE OFF."
	'    End If
	'End Function
	
	
	Function CalcAStallSpeed() As Single
		Dim Sl As Single
		Dim RS As Single
		Dim TempSpeed As Single
		Dim LiftArea As Single
		Dim element As Object
		Dim sStreamlining As String
		Dim sngVThrustNeeded As Single
		
		'JAW 2000.05.25
		'replaced the section that calculates weight in need of support by other means than
		'static lift.
		
		sngVThrustNeeded = m_VWeight - mvarStaticLift
		If sngVThrustNeeded <= 0 Then
			CalcAStallSpeed = 0
			Exit Function
		End If
		
		'see if total static lift exceeds vehicles loadedweight
		sngVThrustNeeded = m_VWeight - mvarStaticLift
		
		If sngVThrustNeeded > 0 Then
			mvaraReservedVectoredThrust = GetUseableVectoredThrust(sngVThrustNeeded, mvarKeyChain, mvarPercentThrust)
			'//if we already know we have enough static lift to cancel out our loaded
			'//weight then we know our stall speed is 0 and we can exit
			If sngVThrustNeeded <= mvaraReservedVectoredThrust Then
				CalcAStallSpeed = 0
				'also, note we will not reset the vectored thrust!
				Exit Function
			End If
		End If
		
		'determine total lift area
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsWing Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.SubType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.SubType = "STOL" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SurfaceArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					LiftArea = LiftArea + (1.5 * element.SurfaceArea)
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SubType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf element.SubType = "flarecraft" Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SurfaceArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					LiftArea = LiftArea + (3 * element.SurfaceArea)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SurfaceArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					LiftArea = LiftArea + element.SurfaceArea
				End If
			ElseIf TypeOf element Is clsRotor Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.SurfaceArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				LiftArea = LiftArea + (3 * element.SurfaceArea)
			End If
		Next element
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Veh.Components(BODY_KEY).LiftingBody Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LiftArea = LiftArea + (0.3 * Veh.Components(BODY_KEY).SurfaceArea)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LiftArea = LiftArea + (0.1 * Veh.Components(BODY_KEY).SurfaceArea)
		End If
		
		'determine the StreamLining modifier
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With Veh.surface
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sStreamlining = .StreamLining
			If sStreamlining = "none" Then
				Sl = 1
			ElseIf sStreamlining = "fair" Then 
				Sl = 1
			ElseIf sStreamlining = "good" Then 
				Sl = 1.05
			ElseIf sStreamlining = "very good" Then 
				Sl = 1.1
			ElseIf sStreamlining = "superior" Then 
				Sl = 1.15
			ElseIf sStreamlining = "excellent" Then 
				Sl = 1.2
			ElseIf sStreamlining = "radical" Then 
				Sl = 1.3
			End If
		End With
		
		'determine Responsive structure modififer
		If mbResponsive Then
			RS = 1.5
		Else : RS = 2
		End If
		
		
		'check for possible divide by zero
		If LiftArea = 0 Then
			CalcAStallSpeed = 0
			Advisory = Advisory & "No Lift Area. "
			Exit Function
		End If
		
		'do final calculation
		TempSpeed = ((m_VWeight - mvarStaticLift - mvaraReservedVectoredThrust) / LiftArea) * Sl * RS
		
		'do final rounding
		If TempSpeed < 0 Then
			CalcAStallSpeed = 0
		Else
			CalcAStallSpeed = System.Math.Round(TempSpeed / 5, 0) * 5 'round to nearest 5mph
		End If
	End Function
	
	
	Function CalcADeceleration(ByRef MR As Single) As Single
		CalcADeceleration = MR * 4
	End Function
	
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		ReDim mvarKeyChain(1)
		mvarDatatype = PERFORMANCEAIR
		
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