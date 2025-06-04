Option Strict Off
Option Explicit On
Module modPerformance
	
	'//////////////////////////////////////////////////////////////////////
	' modPerformance.base  - Michael P. Joseph
	' Created - 11/18/98
	' Helper Functions used by the clsPerformance.cls
	'//////////////////////////////////////////////////////////////////////
	
	Public Function GetMotiveAssemblyKey(ByVal PerformanceType As Integer) As String
		' code originally butchered from frmNewProfile (that form is obsolete)
		Dim KeyChain() As String
		Dim element As Object
		Dim i As Integer
		Dim Datatype As Integer
		On Error GoTo errorhandler
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		KeyChain = VB6.CopyArray(Veh.KeyManager.GetCurrentSubAssembliesKeys)
		
		If PerformanceType = PERFORMANCEWHEEL Then
			PerformanceType = Wheel
		ElseIf PerformanceType = PERFORMANCESKID Then 
			PerformanceType = Skid
		ElseIf PerformanceType = PERFORMANCETRACK Then 
			PerformanceType = Track
		ElseIf PerformanceType = PERFORMANCELEG Then 
			PerformanceType = Leg
		ElseIf PerformanceType = PERFORMANCEFLEX Then 
			PerformanceType = Body
		End If
		
		If UBound(KeyChain) = 1 And KeyChain(1) = "" Then
			GoTo errorhandler
		Else
			For i = 1 To UBound(KeyChain)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Datatype = Veh.Components(KeyChain(i)).Datatype
				If Datatype = PerformanceType Then
					GetMotiveAssemblyKey = KeyChain(i)
					Exit Function
				End If
			Next 
		End If
		'
		'For Each element In Veh.Components
		'    Select Case element.Datatype
		'        Case Wheel, Skid, Track, Leg, FlexibodyDrivetrain
		'            GetMotiveAssembly = element.Key
		'            Exit Function
		'    End Select
		'Next
		
errorhandler: 
		GetMotiveAssemblyKey = ""
	End Function
	
	
	Public Sub GetVehicleWeight(ByVal lngPerformanceType As Integer, ByVal sngPercentAuxVehicleWeight As Integer, ByVal sngPercentCargoWeight As Integer, ByVal sngPercentAmmunitionWeight As Integer, ByVal sngPercentHardpointWeight As Integer, ByVal sngPercentFuelWeight As Integer, ByVal sngPercentProvisionWeight As Integer, ByRef dblWeight As Double, ByRef dblMass As Double)
		
		Dim HardPointWeight As Single
		Dim FuelWeight As Single
		Dim ProvisionsWeight As Single
		Dim AmmoWeight As Single
		Dim GunCarriagesWeight As Single
		Dim CargoWeight As Single
		Dim AuxVehiclesWeight As Single
		Dim element As Object
		
		
		On Error Resume Next
		'//based on the options, this routine sets the m_VWeight and m_VMass variables
		If lngPerformanceType = PERFORMANCESUB Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dblWeight = Veh.Stats.SubmergedWeight
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dblWeight = Veh.Stats.HLoadedWeight 'MPJ 07/25/2000 Was using Loaded Weight instead of HardpointLoadedWeight
			' the user can chose to use a % of hardpoint weight (0 to 100%) in the performance
			' profile Edit dialog
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HardPointWeight = Veh.Stats.HLoadedWeight - Veh.Stats.LoadedWeight
		If HardPointWeight < 0 Then HardPointWeight = 0
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsFuelTank Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.FuelWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FuelWeight = element.FuelWeight
			ElseIf TypeOf element Is clsCargo Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.CargoWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CargoWeight = CargoWeight + element.CargoWeight
				
			ElseIf TypeOf element Is clsProvisions Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ProvisionsWeight = ProvisionsWeight + element.Weight
				'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsVehicleStorage Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.CraftWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AuxVehiclesWeight = AuxVehiclesWeight + element.CraftWeight
				'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			ElseIf TypeName(element) = "clsWeaponAmmunition" Then  'check for ammunition
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AmmoWeight = AmmoWeight + element.Weight
				'ElseIf TypeName(element) = "clsWeaponGun" Then 'check for guns with carriages
				'    If element.Carriage Then
				'        GunCarriagesWeight = GunCarriagesWeight + element.Weight
				'    End If
			End If
		Next element
		
		'//now lets subtract all of these from our weight
		dblWeight = dblWeight - CargoWeight - HardPointWeight - FuelWeight - ProvisionsWeight - AmmoWeight
		
		'//now add the percentages of the weight
		dblWeight = dblWeight + (AuxVehiclesWeight * sngPercentAuxVehicleWeight) + (CargoWeight * sngPercentCargoWeight) + (AmmoWeight * sngPercentAmmunitionWeight) + (HardPointWeight * sngPercentHardpointWeight) + (FuelWeight * sngPercentFuelWeight) + (ProvisionsWeight * sngPercentProvisionWeight)
		
		dblMass = dblWeight / 2000 'get our mass
		
	End Sub
	
	Public Function GetSlowestAnimalSpeed(ByRef KeyChain As Object) As Single
		Dim i As Integer
		Dim sKey As String
		Dim dType As Short
		Dim Slowest As Single
		On Error Resume Next
		
		For i = 1 To UBound(KeyChain)
			'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sKey = KeyChain(i)
			If sKey = "" Then Exit For
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = Veh.Components(sKey).Datatype
			'check for animals.  Max vehicle ground speed cant exceed slowest animal
			If (dType = RopeHarness) Or (dType = YokeandPoleHarness) Or (dType = ShaftandCollarHarness) Or (dType = WhiffletreeHarness) Then
				' set the first animal to the slowest
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Slowest = 0 Then Slowest = Veh.Components(sKey).Speed
				' check if this animal is slower than the current slowest
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(sKey).Speed < Slowest Then Slowest = Veh.Components(sKey).Speed
			End If
		Next 
		
		GetSlowestAnimalSpeed = Slowest
	End Function
	
	
	Public Function GetTiltRotorForwardThrust(ByVal sngPercentThrust As Single) As Single
		'//if the user has checked the TreatTiltRotorsAsPropellers, then
		'//we use their thrust to for forward thrust and NOT for lift
		Dim retval As Single
		Dim element As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsHelicopterDrivetrain Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					Case CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain
						'UPGRADE_WARNING: Couldn't resolve default property of object element.TiltRotor. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If element.TiltRotor Then
							'UPGRADE_WARNING: Couldn't resolve default property of object element.MotivePower. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							retval = retval + (3.5 * element.MotivePower * sngPercentThrust)
						End If
				End Select
			End If
		Next element
		
		GetTiltRotorForwardThrust = retval
		
	End Function
	
	Public Function CalcTotalStaticLift(ByRef KeyChain As Object, ByVal bTreatTiltRotorsAsPropellers As Boolean, ByVal bSEVSidewalls As Boolean, ByVal bGEVSkirt As Boolean, ByVal lngPerformanceType As Integer, ByVal sngPercentThrust As Single, ByVal dblWeight As Double) As Single
		
		Dim i As Short
		Dim Templift As Single
		Dim TempFanLift As Single
		Dim sKey As String
		Dim sngVehicleEmptyWeight As Single
		Dim dType As Integer
		Dim element As Object
		
		'//search the vehicle for liftin gas
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsLiftingGas Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Lift. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Templift = Templift + element.Lift
			End If
		Next element
		
		'next get the lift for Levitation (page 41)
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With Veh.Surface
			sngVehicleEmptyWeight = dblWeight
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .bMagicLevitation Then
				Templift = Templift + sngVehicleEmptyWeight
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .bAntigravityCoating Then
				Templift = Templift + sngVehicleEmptyWeight
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .bSuperScienceCoating Then
				Templift = Templift + sngVehicleEmptyWeight
			End If
		End With
		
		'exit if there are no propulsion systems in the keychain
		'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If KeyChain(1) = "" Then
			CalcTotalStaticLift = Templift
			Exit Function
		End If
		
		'cycle through each item in the array to find all static lift components
		For i = 1 To UBound(KeyChain)
			'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sKey = KeyChain(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = Veh.Components(sKey).Datatype
			
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			With Veh.Components(sKey)
				Select Case dType
					Case CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain
						If bTreatTiltRotorsAsPropellers Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If .TiltRotor Then
								'//we dont add the lift of a tilt rotor when its in forward flight mode
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Templift = Templift + .Lift * sngPercentThrust
							End If
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Templift = Templift + .Lift * sngPercentThrust
						End If
					Case OrnithopterDrivetrain
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Templift = Templift + .Lift * sngPercentThrust
					Case ContraGravGenerator
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Templift = Templift + .Lift '* sngPercentThrust
					Case DuctedFan
						'MPJ 06/30/2000  Fixed.  HoverFan option was not being run at all
						' since the old code was If .Liftengine then if .Hoveran rather than Elseif
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						TempFanLift = .MotiveThrust * sngPercentThrust
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .LiftEngine Then
							Templift = TempFanLift ' no multiplier.  LiftEngines cant be used for forward propulsion though
							
							'//if its a GEV skirt hovercraft then
							'//hoverfan lift is increased by 5
							'//for SEV sidewalls its increased by 4
							'//else its increased by 2
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ElseIf .HoverFan Then 
							If lngPerformanceType = PERFORMANCEHOVER Then
								If bSEVSidewalls Then
									Templift = Templift + (4 * TempFanLift)
								ElseIf bGEVSkirt Then 
									Templift = Templift + (5 * TempFanLift)
								Else
									Templift = Templift + (2 * TempFanLift)
								End If
							End If
						End If
					Case Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .LiftEngine Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If .Afterburner Then 'PPP this should only include AB if its checked!
								'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Templift = Templift + .ABThrust
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Templift = Templift + .MotiveThrust * sngPercentThrust
							End If
						End If
					Case LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .LiftEngine Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Templift = Templift + .MotiveThrust * sngPercentThrust
						End If
					Case OrionEngine
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .LiftEngine Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Templift = Templift + .MotiveThrust * sngPercentThrust
						End If
					Case StandardThruster, SuperThruster, MegaThruster
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .LiftEngine Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Templift = Templift + .MotiveThrust * sngPercentThrust
						End If
				End Select
			End With
		Next 
		
		CalcTotalStaticLift = Templift
	End Function
	
	Function NoPriorPeriscopeChildren(ByVal parentkey As String) As Boolean
		Dim element As Object
		Dim i As Short
		
		i = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If element.LogicalParent = parentkey Then
				i = i + 1
				If i >= 2 Then
					NoPriorPeriscopeChildren = False
					Exit Function
				End If
			End If
		Next element
		' if it makes it out of the loop, no children found and fuction returns TRUE
		NoPriorPeriscopeChildren = True
	End Function
	
	Function GetTotalHitPoints(ByVal Classname As String) As Double
		'returns the total number of hit points for all objects of a given class
		Dim element As Object
		Dim TempHitPoints As Double
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If TypeName(element) = Classname Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.HitPoints. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TempHitPoints = TempHitPoints + element.HitPoints
			End If
		Next element
		GetTotalHitPoints = TempHitPoints
	End Function
	
	'Function GetLowestDR() As Long
	'' this is used to find max speed restrictions for aerial performance
	''only DR from metal, composite or laminate armor counts
	''JAW 2000.06.06
	''
	' Dim element As Object
	' Dim lngRetval As Long
	' Dim lngTemp As Long
	' Dim bDRSet As Boolean
	' Dim lngArmorCount As Long
	' On Error Resume Next
	' '//must initialize the DR to 21 and from there we check for less
	' lngRetval = 21
	'
	'    For Each element In Veh
	'        If TypeOf element Is clsArmor Then
	'            InfoPrint 1, "Armor"
	'            If TypeOf Veh.Components(element.LogicalParent) Is clsPopTurret Then
	'                    'skip these DR's
	'                    InfoPrint 1, " - PopTurret"
	'            ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsSkid Then
	'                    InfoPrint 1, " - Skid"
	'                    If element.Retractable Then
	'                        'skip retractables
	'                    Else
	'                        lngTemp = element.GetLowestDR
	'                        bDRSet = True
	'                    End If
	'            ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsWheel Then
	'                    InfoPrint 1, " - Wheel"
	'                    If Veh.Components(element.LogicalParent).SubType = "retractable" Then
	'                        'skip retractables
	'                    Else
	'                        lngTemp = element.GetLowestDR
	'                        bDRSet = True
	'                    End If
	'            Else
	'                InfoPrint 1, " - other?"
	'                lngTemp = element.GetLowestDR
	'                bDRSet = True
	'            End If
	'        End If
	'        If bDRSet Then
	'            lngRetval = Minimum(lngRetval, lngTemp)
	'        End If
	'    Next
	'
	'    '//if armor was never set, then we no we have no armor
	'    If bDRSet = False Then
	'        lngRetval = 0
	'    End If
	'
	'    GetLowestDR = lngRetval
	'End Function
	
	Function GetLowestDR() As Integer
		' this is used to find max speed restrictions for aerial performance
		'only DR from metal, composite or laminate armor counts
		'JAW 2000.06.19
		'refined to count overall and armor-by-location together properly
		
		Dim element As Object
		Dim lngRetval As Integer
		Dim lngTemp As Integer
		Dim bDRSet As Boolean
		Dim lngArmorCount As Integer
		Dim MinDRX As Integer
		Dim OverallDR As Integer
		Dim DR1 As Integer
		Dim DR2 As Integer
		Dim DR3 As Integer
		Dim DR4 As Integer
		Dim DR5 As Integer
		Dim DR6 As Integer
		
		On Error Resume Next
		'//must initialize the DR to 21 and from there we check for less
		lngRetval = 999999999
		MinDRX = 999999999
		For	Each element In Veh
			If TypeOf element Is clsArmor Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If TypeOf Veh.Components(element.LogicalParent) Is clsPopTurret Then
					'skip these DR's
					'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsSkid Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Retractable. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.Retractable Then
						'skip retractables
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object element.GetLowestDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngTemp = element.GetLowestDR
						bDRSet = True
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsWheel Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Veh.Components(element.LogicalParent).SubType = "retractable" Then
						'skip retractables
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object element.GetLowestDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngTemp = element.GetLowestDR
						bDRSet = True
					End If
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object element. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AccumulateDR(OverallDR, DR1, DR2, DR3, DR4, DR5, DR6, element)
					'lngTemp = element.GetLowestDR
					'bDRSet = True
				End If
			End If
			If bDRSet Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = Minimum(lngRetval, lngTemp)
			End If
		Next element
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MinDRX = Minimum(DR1, MinDRX)
		'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MinDRX = Minimum(DR2, MinDRX)
		'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MinDRX = Minimum(DR3, MinDRX)
		'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MinDRX = Minimum(DR4, MinDRX)
		'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MinDRX = Minimum(DR5, MinDRX)
		'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MinDRX = Minimum(DR6, MinDRX)
		
		
		'if bDRSet is true, then exposed components were found.
		If bDRSet = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lngRetval = Minimum(lngRetval, MinDRX)
		Else
			lngRetval = MinDRX
		End If
		
		GetLowestDR = OverallDR + lngRetval
	End Function
	
	Sub AccumulateDR(ByRef AccumDR As Integer, ByRef AccumDR1 As Integer, ByRef AccumDR2 As Integer, ByRef AccumDR3 As Integer, ByRef AccumDR4 As Integer, ByRef AccumDR5 As Integer, ByRef AccumDR6 As Integer, ByRef ArmorChunk As clsArmor)
		'JAW 2000.06.19
		'mostly cribbed from the GetLowestDR member function of clsArmor
		
		On Error Resume Next
		'Dim lngRetval As Long
		With ArmorChunk
			
			Select Case .Datatype
				Case ArmorComplexFacing
					Select Case .Material1
						Case "metal", "composite", "laminate"
							AccumDR1 = AccumDR1 + .DR1
							Select Case .Material2
								Case "metal", "composite", "laminate"
									AccumDR2 = AccumDR2 + .DR2
									Select Case .Material3
										Case "metal", "composite", "laminate"
											AccumDR3 = AccumDR3 + .DR3
											Select Case .Material4
												Case "metal", "composite", "laminate"
													AccumDR4 = AccumDR4 + .DR4
													Select Case .Material5
														Case "metal", "composite", "laminate"
															AccumDR5 = AccumDR5 + .DR5
															Select Case .Material6
																Case "metal", "composite", "laminate"
																	AccumDR6 = AccumDR6 + .DR6
																Case Else
																	'lngRetval = 0
															End Select
														Case Else
															'lngRetval = 0
													End Select
												Case Else
													'lngRetval = 0
											End Select
										Case Else
											'lngRetval = 0
									End Select
								Case Else
									'lngRetval = 0
							End Select
						Case Else
							'lngRetval = 0
					End Select
					
				Case ArmorBasicFacing
					Select Case .Material
						Case "metal", "composite", "laminate"
							AccumDR1 = AccumDR1 + .DR1
							AccumDR2 = AccumDR2 + .DR2
							AccumDR3 = AccumDR3 + .DR3
							AccumDR4 = AccumDR4 + .DR4
							AccumDR5 = AccumDR5 + .DR5
							AccumDR6 = AccumDR6 + .DR6
							
						Case Else
							'lngRetval = 0
					End Select
					
				Case ArmorLocation, ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
					Select Case .Material
						Case "metal", "composite", "laminate"
							AccumDR = AccumDR + .DR
						Case Else
							'lngRetval = 0
					End Select
					
			End Select
			
		End With
		'GetLowestDR = lngRetval
	End Sub
	
	Function GetLowestTL(ByVal Classname As String) As Integer
		'returns the Tech Level of the lowest Tech Level objects of a given class
		Dim element As Object
		Dim TempTL As Integer
		Dim NewTempTL As Integer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If TypeName(element) = Classname Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.TL. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				NewTempTL = element.TL
			End If
			
			If TempTL = 0 Then
				TempTL = NewTempTL
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TempTL = Minimum(TempTL, NewTempTL)
			End If
		Next element
		GetLowestTL = TempTL
	End Function
	
	Function MinimumNonZero(ByVal x As Object, ByVal y As Object) As Object
		'compares x and y and returns the minimum non zero number
		
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If x = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object MinimumNonZero. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MinimumNonZero = y
			Exit Function
			'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf y = 0 Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object MinimumNonZero. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MinimumNonZero = x
			Exit Function
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MinimumNonZero. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If x < y Then
			MinimumNonZero = x
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object MinimumNonZero. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MinimumNonZero = y
		End If
		
	End Function
	
	
	Function Minimum(ByVal x As Object, ByVal y As Object) As Object
		'compares x and y and returns the minimum of the two
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Minimum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If x < y Then
			Minimum = x
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Minimum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Minimum = y
		End If
	End Function
	
	Function Maximum(ByVal x As Object, ByVal y As Object) As Object
		'compares x and y and returns the maximum of the two
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Maximum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If x > y Then
			Maximum = x
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Maximum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Maximum = y
		End If
	End Function
	
	Function GetHovercraftType() As String
		'returns whether the vehicle has hovercraft with sev sidewalls
		Dim element As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsHovercraft Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.SubType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.SubType = "SEV Sidewalls" Then
					GetHovercraftType = "SEV"
					Exit Function
				Else
					GetHovercraftType = "GEV"
					Exit Function
				End If
			End If
		Next element
		GetHovercraftType = "none"
	End Function
	
	Function VehicleHasResponsiveStruct() As Boolean
		'returns true if the vehicle has a responsive structure
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Veh.Components(BODY_KEY).Responsive Then
			VehicleHasResponsiveStruct = True
		Else
			VehicleHasResponsiveStruct = False
		End If
	End Function
	
	Function VehicleHasWings() As Boolean
		'returns true if the vehicle has wings
		Dim element As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsWing Then
				VehicleHasWings = True
				Exit Function
			End If
		Next element
		
		VehicleHasWings = False
		
	End Function
	
	Function VehicleHasFlarecraftWings() As Boolean
		'returns true if the vehilce has flarecraft wings on any wing assembly
		Dim element As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsWing Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.SubType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.SubType = "flarecraft" Then
					VehicleHasFlarecraftWings = True
					Exit Function
				End If
			End If
		Next element
		VehicleHasFlarecraftWings = False
	End Function
	
	Function VehicleHasRotors() As Boolean
		'returns true if the vehicle has rotors subassembly
		Dim element As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsRotor Then
				VehicleHasRotors = True
				Exit Function
			End If
		Next element
		VehicleHasRotors = False
	End Function
	
	
	Function VehicleHasNonTiltRotors() As Boolean
		'returns true if the vehicle has rotors subassembly
		Dim element As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsHelicopterDrivetrain Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.TiltRotor. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.TiltRotor = False Then
					VehicleHasNonTiltRotors = True
					Exit Function
				End If
			End If
		Next element
		VehicleHasNonTiltRotors = False
	End Function
	
	Function VehicleHasCoaxialRotors() As Boolean
		'returns true if the vehicle has coaxial rotors installs
		Dim element As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsRotor Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.Datatype = CARotor Then
					VehicleHasCoaxialRotors = True
					Exit Function
				End If
			End If
		Next element
		VehicleHasCoaxialRotors = False
	End Function
	
	Function VehiclehasElectORCompcontrols() As Boolean
		'returns true if the vehicle has EITHER electronic or computerized controls
		Dim element As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case ElectronicDivingControl, ComputerizedDivingControl, ElectronicManeuverControl, ComputerizedManeuverControl
					
					VehiclehasElectORCompcontrols = True
					Exit Function
			End Select
		Next element
		VehiclehasElectORCompcontrols = False
	End Function
	
	Function VehicleHasCompControls() As Boolean
		Dim element As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case ComputerizedDivingControl, ComputerizedManeuverControl
					
					VehicleHasCompControls = True
					Exit Function
			End Select
		Next element
		VehicleHasCompControls = False
	End Function
	
	Function AllWingsAreHighAgility() As Boolean
		Dim element As Object
		Dim All As Boolean
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsWing Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.SubType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.SubType = "high agility" Then
					All = True
				Else
					All = False
					Exit For
				End If
			End If
		Next element
		AllWingsAreHighAgility = All
	End Function
	
	Function AllWingsAreVariableSweep() As Boolean
		Dim element As Object
		Dim All As Boolean
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsWing Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.VariableSweep. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.VariableSweep <> "none" Then
					All = True
				Else
					All = False
					Exit For
				End If
			End If
		Next element
		AllWingsAreVariableSweep = All
	End Function
	
	Function AllWingsRotorsControlledInstability() As Boolean
		Dim element As Object
		Dim component As Short
		Dim All As Boolean
		'determines if all wings and or rotors have controlled instability set
		
		All = False 'init the flag
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			component = element.Datatype
			If component = Wing Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.ControlledInstability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.ControlledInstability Then
					All = True
				Else
					All = False
					Exit For
				End If
			ElseIf (component = AutogyroRotor) Or (component = CARotor) Or (component = MMRotor) Or (component = TTRotor) Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.ControlledInstability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.ControlledInstability Then
					All = True
				Else
					All = False
					Exit For
				End If
			End If
		Next element
		AllWingsRotorsControlledInstability = All
	End Function
	
	Function VehicleHasMMRRotors() As Boolean
		Dim element As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If element.Datatype = MMRotor Then
				VehicleHasMMRRotors = True
			End If
		Next element
		VehicleHasMMRRotors = False
	End Function
	
	Function VehicleHasBipeorTripWings() As Object
		Dim element As Object
		Dim All As Boolean
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If element.Datatype = Wing Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.SubType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (element.SubType = "biplane") Or (element.SubType = "triplane") Then
					All = True
				Else
					All = False
					Exit For
				End If
			End If
		Next element
		'UPGRADE_WARNING: Couldn't resolve default property of object VehicleHasBipeorTripWings. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		VehicleHasBipeorTripWings = All
	End Function
	
	Function VehicleHasOnlyStubWings() As Object
		Dim element As Object
		Dim All As Boolean
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsWing Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.SubType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.SubType = "stub" Then
					All = True
				Else
					All = False
					Exit For
				End If
			End If
		Next element
		'UPGRADE_WARNING: Couldn't resolve default property of object VehicleHasOnlyStubWings. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		VehicleHasOnlyStubWings = All
	End Function
End Module