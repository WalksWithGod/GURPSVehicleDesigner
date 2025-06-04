Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cCrew_NET.cCrew")> Public Class cCrew
	
	'todo:  We need to contemplate this hiearchy since we want to handle crew, crew station assignments,
	'       station device assignments, computer/terminal device assignments, and crew shift profiles
	
	
	' Implements cInode
	' Implements cIDisplay
	' Implements cIPersist
	
	' Private m_Crew() as PERSON  ' eventually if we are going to want to store 'crew' as real individuals we can reference
	'                             ' we'll need to keep a list of "crew" and "passenger" objects.
	' Private m_Passengers() as PERSON
	
	' Type SKILL
	'   id as long
	'   level as long
	' End Type
	' Type PERSON  '<-- needs to be an object
	'   jobID as long ' this ID can be a position such as captain, mechanic, etc or a passenger
	'   skills() as SKILL
	' End Type
	
	' ENUM JOBS
	'   JOB_CAPTAIN = 1
	' End ENUM
	Private mvarOccupancy As String
	
	Private mvarNumShifts As Integer
	Private mvarNumCaptains As Integer
	Private mvarNumOfficers As Integer
	Private mvarNumCrewStationOperators As Integer
	Private mvarNumWeaponLoaders As Integer
	Private mvarNumRowers As Integer
	Private mvarNumSailors As Integer
	Private mvarNumRiggers As Integer
	Private mvarNumFuelStokers As Integer
	Private mvarNumMechanics As Integer
	Private mvarNumServiceCrewmen As Integer
	Private mvarNumMedics As Integer
	Private mvarNumScientists As Integer
	Private mvarNumAuxiliaryVehicleCrew As Integer
	Private mvarNumStewards As Integer
	Private mvarNumLuxury As Integer
	Private mvarNumFirstClass As Integer
	Private mvarNumSecondClass As Integer
	Private mvarNumSteerage As Integer
	Private mvarTotalNumberCrewPassengers As Integer
	Private mvarUseRecommendedCrew As Boolean
	Private mvarMilitaryVehicle As Boolean
	
	
	Public Property MilitaryVehicle() As Boolean
		Get
			MilitaryVehicle = mvarMilitaryVehicle
		End Get
		Set(ByVal Value As Boolean)
			mvarMilitaryVehicle = Value
		End Set
	End Property
	Public Property UseRecommendedCrew() As Boolean
		Get
			UseRecommendedCrew = mvarUseRecommendedCrew
		End Get
		Set(ByVal Value As Boolean)
			mvarUseRecommendedCrew = Value
		End Set
	End Property
	Public Property NumOfficers() As Integer
		Get
			NumOfficers = mvarNumOfficers
		End Get
		Set(ByVal Value As Integer)
			mvarNumOfficers = Value
		End Set
	End Property
	Public Property NumShifts() As Integer
		Get
			NumShifts = mvarNumShifts
		End Get
		Set(ByVal Value As Integer)
			mvarNumShifts = Value
		End Set
	End Property
	Public Property NumCaptains() As Integer
		Get
			NumCaptains = mvarNumCaptains
		End Get
		Set(ByVal Value As Integer)
			mvarNumCaptains = Value
		End Set
	End Property
	Public Property NumCrewStationOperators() As Integer
		Get
			NumCrewStationOperators = mvarNumCrewStationOperators
		End Get
		Set(ByVal Value As Integer)
			mvarNumCrewStationOperators = Value
		End Set
	End Property
	Public Property NumWeaponLoaders() As Integer
		Get
			NumWeaponLoaders = mvarNumWeaponLoaders
		End Get
		Set(ByVal Value As Integer)
			mvarNumWeaponLoaders = Value
		End Set
	End Property
	Public Property NumRowers() As Integer
		Get
			NumRowers = mvarNumRowers
		End Get
		Set(ByVal Value As Integer)
			mvarNumRowers = Value
		End Set
	End Property
	Public Property NumSailors() As Integer
		Get
			NumSailors = mvarNumSailors
		End Get
		Set(ByVal Value As Integer)
			mvarNumSailors = Value
		End Set
	End Property
	Public Property NumRiggers() As Integer
		Get
			NumRiggers = mvarNumRiggers
		End Get
		Set(ByVal Value As Integer)
			mvarNumRiggers = Value
		End Set
	End Property
	Public Property NumFuelStokers() As Integer
		Get
			NumFuelStokers = mvarNumFuelStokers
		End Get
		Set(ByVal Value As Integer)
			mvarNumFuelStokers = Value
		End Set
	End Property
	Public Property NumMechanics() As Integer
		Get
			NumMechanics = mvarNumMechanics
		End Get
		Set(ByVal Value As Integer)
			mvarNumMechanics = Value
		End Set
	End Property
	Public Property NumServiceCrewmen() As Integer
		Get
			NumServiceCrewmen = mvarNumServiceCrewmen
		End Get
		Set(ByVal Value As Integer)
			mvarNumServiceCrewmen = Value
		End Set
	End Property
	Public Property NumMedics() As Integer
		Get
			NumMedics = mvarNumMedics
		End Get
		Set(ByVal Value As Integer)
			mvarNumMedics = Value
		End Set
	End Property
	Public Property NumScientists() As Integer
		Get
			NumScientists = mvarNumScientists
		End Get
		Set(ByVal Value As Integer)
			mvarNumScientists = Value
		End Set
	End Property
	Public Property NumAuxiliaryVehicleCrew() As Integer
		Get
			NumAuxiliaryVehicleCrew = mvarNumAuxiliaryVehicleCrew
		End Get
		Set(ByVal Value As Integer)
			mvarNumAuxiliaryVehicleCrew = Value
		End Set
	End Property
	Public Property NumStewards() As Integer
		Get
			NumStewards = mvarNumStewards
		End Get
		Set(ByVal Value As Integer)
			mvarNumStewards = Value
		End Set
	End Property
	Public Property NumLuxury() As Integer
		Get
			NumLuxury = mvarNumLuxury
		End Get
		Set(ByVal Value As Integer)
			mvarNumLuxury = Value
		End Set
	End Property
	Public Property NumFirstClass() As Integer
		Get
			NumFirstClass = mvarNumFirstClass
		End Get
		Set(ByVal Value As Integer)
			mvarNumFirstClass = Value
		End Set
	End Property
	Public Property NumSecondClass() As Integer
		Get
			NumSecondClass = mvarNumSecondClass
		End Get
		Set(ByVal Value As Integer)
			mvarNumSecondClass = Value
		End Set
	End Property
	Public Property NumSteerage() As Integer
		Get
			NumSteerage = mvarNumSteerage
		End Get
		Set(ByVal Value As Integer)
			mvarNumSteerage = Value
		End Set
	End Property
	Public Property Occupancy() As String
		Get
			Occupancy = mvarOccupancy
		End Get
		Set(ByVal Value As String)
			mvarOccupancy = Value
		End Set
	End Property
	Public Property TotalNumberCrewPassengers() As Integer
		Get
			TotalNumberCrewPassengers = mvarTotalNumberCrewPassengers
		End Get
		Set(ByVal Value As Integer)
			mvarTotalNumberCrewPassengers = Value
		End Set
	End Property
	
	
	
	Sub StatsUpdate()
		Dim NumCaptains As Integer
		Dim NumOfficers As Integer
		Dim NumCrewStationOperators As Integer
		Dim NumLoaders As Integer
		Dim NumRowers As Integer
		Dim NumSailors As Integer
		Dim NumGasRiggers As Integer
		Dim NumFuelStokers As Integer
		Dim NumMechanics As Integer
		Dim NumServiceCrewmen As Integer
		Dim NumMedics As Integer
		Dim NumScientists As Integer
		Dim NumAuxiliaryVehicleCrewmen As Integer
		Dim NumStewards As Single 'depends on number of each class of passenger '6.02.2000 changed from Long to Single so it can be rounded up after NumShifts
		Dim NumLuxury As Integer 'user defined
		Dim NumFirstClass As Integer 'user defined
		Dim NumSecondClass As Integer 'user defined
		Dim NumSteerage As Integer 'user defined
		Dim TotalPassengers As Integer
		Dim TotalCrewSize As Integer
		Dim element As Object
		Dim SailDivisor As Short
		Dim GasVolume As Double
		Dim NumShifts As Short
		Dim LongTerm As Boolean
		Dim MilitaryVehicle As Boolean
		Dim i As Integer
		Dim dType As Short
		Dim nManeuverControlCount As Integer
		Dim element2 As Object
		Dim bMechanical As Boolean
		Dim arrSubs() As String
		Dim j As Integer
		Dim TotalMastHeight As Integer
		Dim lngDiv As Integer
		Dim dblTempPower As Double
		Dim bRequiresMedics As Boolean
		Dim lngTempMedics As Integer
		
		
		'init totalcrewsize to 0
		TotalCrewSize = 0
		
		'determine if long or short term occupancy
		If mvarOccupancy = "long" Then
			LongTerm = True
		End If
		
		'determine if military vehicle
		If mvarMilitaryVehicle Then
			MilitaryVehicle = True
		End If
		
		'get the number of work shifts
		NumShifts = mvarNumShifts
		
		'get the total mast height
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrSubs = VB6.CopyArray(Veh.KeyManager.GetCurrentSubAssembliesKeys)
		If arrSubs(1) <> "" Then
			For j = 1 To UBound(arrSubs)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(arrSubs(j)).Datatype = Mast Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TotalMastHeight = TotalMastHeight + (Veh.Components(arrSubs(j)).Quantity * Veh.Components(arrSubs(j)).Height)
				End If
			Next 
		End If
		'determine maneuver control modifiers which will be used for # of Sailors needed
		If gVehicleTL >= 7 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For	Each element2 In Veh.Components
				'UPGRADE_WARNING: Couldn't resolve default property of object element2.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dType = element2.Datatype
				If (dType = MechanicalManeuverControl) Or (dType = MechanicalDivingControl) Then
					bMechanical = True
					Exit For
				ElseIf TypeOf element2 Is clsManeuverControl Then 
					nManeuverControlCount = 1
					Exit For
				End If
			Next element2
		End If
		'//now we can search the vehicle and determine how many crew we need based on
		'what we find
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = element.Datatype
			If (dType = RoomyCrewStation) Or (dType = NormalCrewStation) Or (dType = CrampedCrewStation) Or (dType = CycleCrewStation) Or (dType = HarnessCrewStation) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Quantity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				NumCrewStationOperators = NumCrewStationOperators + element.Quantity
			ElseIf (dType = BattlesuitSystem) Or (dType = FormFittingBattleSuitSystem) Then 
				NumCrewStationOperators = NumCrewStationOperators + 1
			ElseIf (TypeOf element Is clsWeaponStoneBoltThrower) Or (TypeOf element Is clsWeaponGun) Or (TypeOf element Is clsWeaponLauncher) Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Loaders. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				NumLoaders = NumLoaders + element.Loaders
			ElseIf dType = RowingPositions Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Quantity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				NumRowers = NumRowers + element.Quantity
			ElseIf (dType = HotAir) Or (dType = Hydrogen) Or (dType = Helium) Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Volume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GasVolume = GasVolume + element.Volume
			ElseIf TypeOf element Is clsLabandWorkshop Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Quantity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				NumScientists = NumScientists + element.Quantity
			ElseIf (dType = SquareRig) Or (dType = ForeandAftRig) Or (dType = FullRig) Or (dType = AerialSail) Or (dType = AerialSailForeAftRig) Then 
				If (gVehicleTL >= 7) And (bMechanical = False) And (nManeuverControlCount <> 0) Then
					NumSailors = 0
				Else
					If dType = ForeandAftRig Then SailDivisor = 5 Else SailDivisor = 4
					NumSailors = System.Math.Sqrt(TotalMastHeight) / SailDivisor
				End If
			End If
		Next element
		
		'get the number of mechanics
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element2 In Veh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element2.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element2.Datatype
				
				Case TrackedDrivetrain, LegDrivetrain, FlexibodyDrivetrain, WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel, DuctedFan, AerialPropeller, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, OrnithopterDrivetrain, MagLevLifter, TeleportationDrive, JumpDrive, WarpDrive, ContraGravGenerator, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element2.PowerReqt. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					dblTempPower = dblTempPower + element2.PowerReqt
					
					'Case Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam
					'//jet engines dont have a power reqt so i guess they dont apply since the
					'//rules on page 75 for mechanics only mention Power Reqt systems
					
				Case Hyperdrive
					'UPGRADE_WARNING: Couldn't resolve default property of object element2.SustainedPower. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					dblTempPower = dblTempPower + element2.SustainedPower
			End Select
		Next element2
		
		'//get the divisor
		Select Case gVehicleTL
			Case Is <= 5
				lngDiv = 250
			Case 6
				lngDiv = 500
			Case 7
				lngDiv = 1000
			Case 8
				lngDiv = 5000
			Case 9
				lngDiv = 50000
			Case 10
				lngDiv = 250000
			Case 11
				lngDiv = 500000
			Case 12, 13
				lngDiv = 1000000
			Case Is >= 14
				lngDiv = 5000000
		End Select
		NumMechanics = Int(System.Math.Sqrt(dblTempPower / lngDiv))
		
		'get number of stewards
		NumLuxury = mvarNumLuxury
		NumFirstClass = mvarNumFirstClass
		NumSecondClass = mvarNumSecondClass
		NumSteerage = mvarNumSteerage
		TotalPassengers = NumSteerage + NumSecondClass + NumFirstClass + NumLuxury ' MPJ 06.02.2000 Medics will now take into account passenger count too
		
		NumStewards = NumSteerage / 100
		NumStewards = NumStewards + NumSecondClass / 50
		NumStewards = NumStewards + NumFirstClass / 20
		NumStewards = NumStewards + NumLuxury / 4
		
		
		'get number of gasriggers required (Linked to shifts)
		NumGasRiggers = (GasVolume / 200000)
		If GasVolume < 200000 Then NumGasRiggers = 0 'MPJ 05.19.2000
		
		'get number of medics needed (not linked to Shifts since it dependant on totalcrewsize)
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components 'MPJ 05.18.2000 - Updated this entire medics crew search
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = element.Datatype
			If (dType = StretcherPallet) Or (dType = DiagnosisTable) Then
				bRequiresMedics = True
			ElseIf dType = OperatingRoom Then 
				NumMedics = NumMedics + 1
			End If
		Next element
		' if we require at least one medic but have no operating tables, set NumMedics to 1
		If NumMedics < 1 And (bRequiresMedics) Then NumMedics = 1
		' make sure we have at least as many doctors needed for total crew size
		If LongTerm Then
			If gVehicleTL <= 4 Then
				lngTempMedics = TotalCrewSize + TotalPassengers / 50
			ElseIf gVehicleTL <= 7 Then 
				lngTempMedics = TotalCrewSize + TotalPassengers / 100
			Else
				lngTempMedics = TotalCrewSize + TotalPassengers / 200
			End If
		End If
		If lngTempMedics > NumMedics Then NumMedics = lngTempMedics
		
		
		'For all of the above ones, get final number based on number of Work Shifts
		NumMedics = NumMedics * NumShifts
		NumScientists = NumScientists * NumShifts
		NumGasRiggers = NumGasRiggers * NumShifts
		NumCrewStationOperators = NumCrewStationOperators * NumShifts
		NumRowers = NumRowers * NumShifts
		NumLoaders = NumLoaders * NumShifts
		NumSailors = NumSailors * NumShifts
		NumMechanics = NumMechanics * NumShifts
		NumFuelStokers = NumFuelStokers * NumShifts
		NumStewards = System.Math.Round(NumStewards * NumShifts, 0)
		
		
		'NumAuxiliaryVehicleCrewmen  'TODO: Right now it doesnt calc Auxiliary Crew.  This should
		'be a user defined choice anyway since its beyond the scope of GVD to calc crew of
		' other vehicles loaded onto this vehicle.
		NumAuxiliaryVehicleCrewmen = mvarNumAuxiliaryVehicleCrew
		
		'add them all up so far.  MPJ 06.02.2000 Total Passengers Now added Here instead of further below. (Note previous version appeared to add them here too, but the values were never set for them)
		TotalCrewSize = TotalPassengers + NumCaptains + NumOfficers + NumCrewStationOperators + NumLoaders + NumRowers + NumSailors + NumGasRiggers + NumFuelStokers + NumMechanics + NumServiceCrewmen + NumScientists + NumAuxiliaryVehicleCrewmen + NumStewards
		
		'get number of service crewmen needed (not linked to shifts)
		If LongTerm Then
			NumServiceCrewmen = TotalCrewSize / 20
		End If
		TotalCrewSize = TotalCrewSize + NumServiceCrewmen
		
		
		'update the total crew size
		TotalCrewSize = TotalCrewSize + NumMedics
		
		'get number of captains/commanders (Linked to Shifts)
		If (MilitaryVehicle) And (TotalCrewSize > 1) Then
			NumCaptains = 1 * NumShifts
		ElseIf TotalCrewSize > 12 Then 
			NumCaptains = 1 * NumShifts
		End If
		
		'get number of officers needed (not linked to shifts since its dependant on TotalCrewSize)
		NumOfficers = TotalCrewSize / 8
		
		
		TotalCrewSize = TotalCrewSize + NumOfficers
		
		'save it all
		mvarNumCaptains = NumCaptains
		mvarNumOfficers = NumOfficers
		mvarNumCrewStationOperators = NumCrewStationOperators
		mvarNumWeaponLoaders = NumLoaders
		mvarNumRowers = NumRowers
		mvarNumSailors = NumSailors
		mvarNumRiggers = NumGasRiggers
		mvarNumMechanics = NumMechanics
		mvarNumFuelStokers = NumFuelStokers
		mvarNumServiceCrewmen = NumServiceCrewmen
		mvarNumMedics = NumMedics
		mvarNumScientists = NumScientists
		mvarNumAuxiliaryVehicleCrew = NumAuxiliaryVehicleCrewmen
		mvarNumStewards = NumStewards
		
		'note that the passengers arent saved because the program
		'cant calculate how many passengers there should be.  User must enter.
		mvarTotalNumberCrewPassengers = TotalCrewSize
		
	End Sub
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		mvarNumShifts = 1
		mvarUseRecommendedCrew = True
		mvarOccupancy = "long"
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class