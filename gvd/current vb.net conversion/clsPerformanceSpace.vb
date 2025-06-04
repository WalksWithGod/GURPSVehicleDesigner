Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsPerformanceSpace_NET.clsPerformanceSpace")> Public Class clsPerformanceSpace
	
	Private mvarsManeuverability As Single
	Private mvarsTurnAroundTime As Single
	Private mvarsMotiveThrust As Single
	Private mvarsAccelerationG As Single
	Private mvarsAccelerationMPH As Single
	Private mvarsHyperSpeed As Single
	Private mvarsWarpSpeed As Single
	Private mvarsJumpDriveable As Boolean
	Private mvarsTeleportationDriveable As Boolean
	
	Private mvarKey As String
	Private mvarParent As String
	Private mvarDatatype As Integer
	Private mvarDescription As String
	
	Private mbResponsive As Boolean
	
	Private mvarHardPointsOn As Boolean
	
	
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
	
	
	
	Public Property sMotiveThrust() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.sMotiveThrust
			sMotiveThrust = mvarsMotiveThrust
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.sMotiveThrust = 5
			mvarsMotiveThrust = Value
		End Set
	End Property
	
	
	
	
	Public Property sAccelerationG() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.sAccelerationG
			sAccelerationG = mvarsAccelerationG
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.sAccelerationG = 5
			mvarsAccelerationG = Value
		End Set
	End Property
	
	
	
	Public Property sAccelerationMPH() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.sAccelerationMPH
			sAccelerationMPH = mvarsAccelerationMPH
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.sAccelerationMPH = 5
			mvarsAccelerationMPH = Value
		End Set
	End Property
	
	
	
	Public Property sTurnAroundTime() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.sTurnAroundTime
			sTurnAroundTime = mvarsTurnAroundTime
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.sTurnAroundTime = 5
			mvarsTurnAroundTime = Value
		End Set
	End Property
	
	
	
	Public Property sHyperSpeed() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.sHyperSpeed
			sHyperSpeed = mvarsHyperSpeed
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.sHyperSpeed = 5
			mvarsHyperSpeed = Value
		End Set
	End Property
	
	
	Public Property sWarpSpeed() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.sWarpSpeed
			sWarpSpeed = mvarsWarpSpeed
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.sWarpSpeed = 5
			mvarsWarpSpeed = Value
		End Set
	End Property
	
	
	Public Property sManeuverability() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.sManeuverability
			sManeuverability = mvarsManeuverability
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.sManeuverability = 5
			mvarsManeuverability = Value
		End Set
	End Property
	
	
	Public Property sTeleportationDriveable() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.sTeleportationDriveable
			sTeleportationDriveable = mvarsTeleportationDriveable
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.sTeleportationDriveable = 5
			mvarsTeleportationDriveable = Value
		End Set
	End Property
	
	
	Public Property sJumpDriveable() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.sJumpDriveable
			sJumpDriveable = mvarsJumpDriveable
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.sJumpDriveable = 5
			mvarsJumpDriveable = Value
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
		
		Call GetVehicleWeight(PERFORMANCESPACE, mvarPercentAuxVehicleWeight, mvarPercentCargoWeight, mvarPercentAmmunitionWeight, mvarPercentHardpointWeight, mvarPercentFuelWeight, mvarPercentProvisionWeight, m_VWeight, m_VMass)
		
		
		
		
		mvarsMotiveThrust = System.Math.Round(CalcSMotiveThrust, 2)
		mvarsAccelerationG = System.Math.Round(CalcSAccelerationG, 4)
		mvarsAccelerationMPH = System.Math.Round(CalcSAccelerationMPH, 4)
		mvarsManeuverability = System.Math.Round(mvarsAccelerationG, 2) 'this equals acceleration
		mvarsTurnAroundTime = System.Math.Round(CalcSTurnAroundTime, 2)
		mvarsHyperSpeed = System.Math.Round(CalcSHyperspeed, 2)
		mvarsWarpSpeed = System.Math.Round(CalcSWarpSpeed, 2)
		mvarsJumpDriveable = CalcJumpDriveable
		mvarsTeleportationDriveable = CalcTeleportationDriveable
		
	End Sub
	
	Function CalcSTurnAroundTime() As Single
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalcSTurnAroundTime = modPerformance.Maximum(Veh.Stats.SizeModifier * 10, 1)
	End Function
	
	Function CalcSMotiveThrust() As Object
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
				
				Case LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine
					
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempThrust = TempThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
					
				Case lightSail
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempThrust = TempThrust + Veh.Components(sKey).Thrust * mvarPercentThrust
			End Select
		Next 
		
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcSMotiveThrust. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalcSMotiveThrust = TempThrust
		
		
	End Function
	
	Function CalcSAccelerationG() As Single
		If m_VWeight <= 0 Then
			CalcSAccelerationG = 0
		Else
			CalcSAccelerationG = mvarsMotiveThrust / m_VWeight
		End If
	End Function
	
	Function CalcSAccelerationMPH() As Single
		Const GravitiesToMPH As Double = 21.9
		
		CalcSAccelerationMPH = mvarsAccelerationG * GravitiesToMPH
	End Function
	
	Function CalcSHyperspeed() As Single
		'Note: this only uses Hyperspeed drives that are added to the keychain since
		'other hyperdrives can be carried as just cargo
		
		'Note: the errate allows speeds higher than the flat .2 per day.  The book incorrectly states
		' a max of .2 parsecs a day
		
		Const HyperSpeedConstant As Double = 0.2 '.2 parsecs
		Dim Lmass As Single
		Dim i As Integer
		Dim HP As Single
		Dim sKey As String
		Dim dType As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarKeyChain(1) = "" Then Exit Function 'exit if there are no propulsion systems in the keychain
		
		Lmass = m_VMass
		
		For i = 1 To UBound(mvarKeyChain)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sKey = mvarKeyChain(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = Veh.Components(sKey).Datatype
			If dType = Hyperdrive Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				HP = HP + Veh.Components(sKey).capacity
			End If
		Next 
		
		CalcSHyperspeed = HP / Lmass * HyperSpeedConstant
	End Function
	
	Function CalcSWarpSpeed() As Single
		'Note: this only uses Warp drives that are added to the keychain since
		'other Warp drives can be carried as just cargo
		Const WarpSpeedConstant As Short = 1 '1 parsec per day
		Dim Lmass As Single
		Dim i As Integer
		Dim sKey As String
		Dim dType As Short
		Dim WF As Single
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarKeyChain(1) = "" Then Exit Function 'exit if there are no propulsion systems in the keychain
		
		Lmass = m_VMass
		
		For i = 1 To UBound(mvarKeyChain)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sKey = mvarKeyChain(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = Veh.Components(sKey).Datatype
			If dType = WarpDrive Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				WF = WF + Veh.Components(sKey).capacity
			End If
		Next 
		
		CalcSWarpSpeed = WF / Lmass * WarpSpeedConstant
		
	End Function
	
	Function CalcJumpDriveable() As Boolean
		'determine if the ship is capable of Jump travel via jump drives
		'Note: this only uses Jump drives that are added to the keychain since
		'other Jump can be carried as just cargo
		Dim TempRating As Single
		Dim i As Integer
		Dim sKey As String
		Dim dType As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarKeyChain(1) = "" Then Exit Function 'exit if there are no propulsion systems in the keychain
		
		For i = 1 To UBound(mvarKeyChain)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sKey = mvarKeyChain(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = Veh.Components(sKey).Datatype
			If dType = JumpDrive Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TempRating = TempRating + Veh.Components(sKey).capacity
			End If
		Next 
		
		If TempRating >= m_VMass Then
			CalcJumpDriveable = True
		Else
			CalcJumpDriveable = False
		End If
		
	End Function
	
	Function CalcTeleportationDriveable() As Boolean
		'determine if the ship is capable of Teleportation travel
		'Note: this only uses Teleportation drives that are added to the keychain since
		'other Teleportation drives can be carried as just cargo
		Dim i As Integer
		Dim sKey As String
		Dim dType As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarKeyChain(1) = "" Then Exit Function 'exit if there are no propulsion systems in the keychain
		
		For i = 1 To UBound(mvarKeyChain)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarKeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sKey = mvarKeyChain(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = Veh.Components(sKey).Datatype
			If dType = TeleportationDrive Then
				CalcTeleportationDriveable = True
				Exit Function
			End If
		Next 
		
		CalcTeleportationDriveable = False
		
	End Function
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		ReDim mvarKeyChain(1)
		mvarDatatype = PERFORMANCESPACE
		
		'mvarAfterBurnersOn = True ' not used in space
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