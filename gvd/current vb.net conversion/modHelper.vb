Option Strict Off
Option Explicit On
Module modHelper
	
	Public InfoBox As Object
	Public p_sFormat As String ' todo: delete this var eventually.  This is the users format setting from the old config dialog.  I believe its obsolete now with our metric dialog
	
	
	'/////////////////////////////////////////////////
	'the Collection which will hold all of the components attached to this vehicle
	Public Veh As cVehicle 'todo: need this still?  check cVehicle.Initialize() first
	Public gVehicleTL As Short ' the vehicles tech level.  This should be a property of vehicle rather than global IF we do use Veh
	
	
	'//////////////////////////////////////////////////////////////////////
	'//////////////////////////////////////////////////////////////////////
	'Below are Helper functions called from the individual classes
	'//////////////////////////////////////////////////////////////////////
	'//////////////////////////////////////////////////////////////////////
	Public Function RoundUP(ByRef x As Object) As Single
		Dim Temp As Single
		Temp = Int(CDbl(x))
		'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Temp < x Then Temp = Temp + 1
		RoundUP = Temp
	End Function
	
	Public Sub InfoPrint(ByVal Code As Short, ByRef Message As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object InfoBox.Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InfoBox.Text = Message & vbNewLine & Left(InfoBox.Text, 2000)
	End Sub
	
	
	
	Public Function GetLogicalParent(ByVal parentkey As String) As String
		'this function gets the logical parent key.  A logical parent is not necessarily the parent node
		' of a component.  For instance, a Group  component can never be a logical parent.  If something is
		' attached to a Group component, the logical parent is the parent of that group component
		
		Dim retval As String
		Dim TempLocation As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		TempLocation = TypeName(Veh.Components(parentkey))
		
		Select Case TempLocation
			Case "clsGroup"
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				retval = GetLogicalParent(Veh.Components(parentkey).Parent)
				
			Case "clsWeaponLauncher"
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				retval = GetLogicalParent(Veh.Components(parentkey).Parent)
			Case "clsModule"
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				retval = GetLogicalParent(Veh.Components(parentkey).Parent)
				
			Case "clsBody"
				retval = parentkey
			Case "clsTurret"
				retval = parentkey
			Case "clsPopTurret"
				retval = parentkey
			Case "clsSuperStructure"
				retval = parentkey
			Case "clsArm"
				retval = parentkey
			Case "clsLeg"
				retval = parentkey
			Case "clsPod"
				retval = parentkey
			Case "clsWing"
				retval = parentkey
			Case "clsEquipmentPod"
				retval = parentkey
			Case "clsGasbag"
				retval = parentkey
			Case "clsHovercraft"
				retval = parentkey
			Case "clsHydrofoil"
				retval = parentkey
			Case "clsRotor"
				retval = parentkey
			Case "clsMast"
				retval = parentkey
			Case "clsHardPoint"
				retval = parentkey
			Case "clsOpenMount"
				retval = parentkey
			Case "clsWheel"
				retval = parentkey
			Case "clsTrack"
				retval = parentkey
			Case "clsSkid"
				retval = parentkey
			Case "clsSolarPanel"
				retval = parentkey
			Case Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				retval = Veh.Components(parentkey).Parent
		End Select
		
		GetLogicalParent = retval
	End Function
	
	
	
	Public Function NumericToString(ByVal nNumber As Object) As String
		'//this function accepts a number and if that number is between
		'1 and 10 it will convert them to "One" and "Ten" for instance. If
		'the number is greater than 10 it will just return the number formatted
		'as a string
		Dim retval As String
		
		'NOTE: This is currently only set up to handle longs and not decimals
		'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nNumber >= 1 And nNumber <= 10 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If nNumber = 1 Then
				'retval = "One"
				retval = "" 'if its a 1 we'll jsut leave it blank since its assumed to be 1 unless noted
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 2 Then 
				retval = "Two"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 3 Then 
				retval = "Three"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 4 Then 
				retval = "Four"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 5 Then 
				retval = "Five"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 6 Then 
				retval = "Six"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 7 Then 
				retval = "Seven"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 8 Then 
				retval = "Eight"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 9 Then 
				retval = "Nine"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 10 Then 
				retval = "Ten"
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			retval = "(" & VB6.Format(nNumber) & ")"
		End If
		
		NumericToString = retval
		
	End Function
	
	Public Function ConvertDamage(ByVal Damage As Object) As Object
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Double 'holds info to the left of the decimal
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Double 'holds info to the right of the decimal
		Dim LeftConvert As String
		Dim RightConvert As String
		Dim TempConvert As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Damage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Right_Renamed = Damage - Fix(Damage) 'get the number left of the decimal and retain the decimal
		Left_Renamed = Int(CDbl(Damage)) 'get number left of decimal
		LeftConvert = Str(Left_Renamed)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Damage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Damage < 1 Then 'fractional damages of les than 1d
			If Damage < 0.1 Then
				TempConvert = "No Damage"
			ElseIf Damage <= 0.2 Then 
				TempConvert = "1d-4"
			ElseIf Damage <= 0.4 Then 
				TempConvert = "1d-3"
			ElseIf Damage <= 0.6 Then 
				TempConvert = "1d-2"
			ElseIf Damage <= 0.8 Then 
				TempConvert = "1d-1"
			Else : TempConvert = "1d"
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object ConvertDamage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ConvertDamage = TempConvert
			'UPGRADE_WARNING: Couldn't resolve default property of object Damage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf Damage < 24 Then  'fractional damages between 1d and 24d
			If Right_Renamed <= 0.2 Then
				RightConvert = ""
			ElseIf Right_Renamed <= 0.4 Then 
				RightConvert = "+1"
			ElseIf Right_Renamed <= 0.6 Then 
				RightConvert = "+2"
			ElseIf Right_Renamed <= 0.8 Then 
				RightConvert = "-1"
				Left_Renamed = Left_Renamed + 1
			Else
				RightConvert = ""
				Left_Renamed = Left_Renamed + 1
			End If
			LeftConvert = Val(CStr(Left_Renamed)) & "d"
			TempConvert = LeftConvert & RightConvert
			'UPGRADE_WARNING: Couldn't resolve default property of object ConvertDamage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ConvertDamage = TempConvert
		Else 'fractional damages of 24d and more
			LeftConvert = "6d x"
			'UPGRADE_WARNING: Couldn't resolve default property of object Damage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Right_Renamed = System.Math.Round(Damage / 6, 0)
			RightConvert = CStr(Val(CStr(Right_Renamed)))
			TempConvert = LeftConvert & RightConvert
			'UPGRADE_WARNING: Couldn't resolve default property of object ConvertDamage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ConvertDamage = TempConvert
		End If
		
		
	End Function
	
	Public Function DecreaseMalf(ByVal Malf As String) As String
		Dim value As Short
		Dim TempMalf As String
		
		value = Val(Malf)
		If value = 1 Then
			TempMalf = "1" ' cant be less than 1
		ElseIf value > 1 Then 
			value = value - 1
			TempMalf = Str(value)
		Else
			If Malf = "Crit." Then
				TempMalf = "16"
			ElseIf Malf = "Ver." Then 
				TempMalf = "Crit."
			ElseIf Malf = "Ver.(Crit.)" Then 
				TempMalf = "Ver."
			End If
		End If
		DecreaseMalf = TempMalf
	End Function
	
	Public Function IncreaseMalf(ByVal Malf As String) As String
		Dim value As Short
		Dim TempMalf As String
		
		value = Val(Malf)
		If value = 16 Then
			TempMalf = "Crit."
		ElseIf value >= 1 Then 
			value = value + 1
			TempMalf = Str(value)
		Else
			If Malf = "Crit." Then
				TempMalf = "Ver."
			ElseIf Malf = "Ver." Then 
				TempMalf = "Ver.(Crit.)"
			ElseIf Malf = "Ver.(Crit.)" Then 
				TempMalf = "Ver.(Crit.)"
			End If
		End If
		
		IncreaseMalf = TempMalf
	End Function
	
	Public Function CalcAccessSpace(ByVal Key As String) As Double
		Dim element As Object
		Dim tempSpace As Double
		Dim sKey As String
		
		' find every child component to a subassembly, then determine if its uses access space
		' todo: if using an itterator and heirarchal vehicle structure, I can simply itterate through the logical parent's
		' sub notes to check JUST the children.
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If element.LogicalParent = Key Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					'these are the Power and Fuel systems
					Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, FlexibodyDrivetrain, FlexibodyDrivetrain, TrackedDrivetrain, LegDrivetrain, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, AerialPropeller, DuctedFan, PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel, MagLevLifter, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, StandardThruster, SuperThruster, MegaThruster, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, SolidRocketEngine, OrionEngine, TeleportationDrive, Hyperdrive, JumpDrive, WarpDrive, QuantumConveyor, SubQuantumConveyor, TwoQuantumConveyor 'ContraGravGenerator  'according to David Pulver, contragrav does not need it
						
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Volume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						tempSpace = tempSpace + element.Volume
						
						'these are the powered Propulsion and Lift Systems (NOTE: any change to this list
						' should also change list in the Gettotalpropulsioncost function
					Case GasolineEngine, HPGasolineEngine, TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, SuperHPGasolineEngine, StandardDieselEngine, TurboStandardDieselEngine, MarineDieselEngine, HPDieselEngine, TurboHPDieselEngine, CeramicEngine, TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, TurboHPCeramicEngine, SuperHPCeramicEngine, HydrogenCombustionEngine, EarlySteamEngine, ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine, StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, StandardMHDTurbine, HPMHDTurbine, FuelCell, FissionReactor, RTGReactor, NPU, FusionReactor, AntimatterReactor, TotalConversionPowerPlant, CosmicPowerPlant, Soulburner, ElementalFurnace, ManaEngine, Carnivore, Herbivore, Omnivore, Vampire
						'ClockWork, LeadAcidBattery, AdvancedBattery, Flywheel, _
						''RechargeablePowerCell, PowerCell, NitrousOxideBooster
						
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Volume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						tempSpace = tempSpace + element.Volume
				End Select
			End If
		Next element
		'apply option from the Option Dialog
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		tempSpace = tempSpace * Veh.Options.AccessSpaceVolumeMod
		
		CalcAccessSpace = System.Math.Round(tempSpace, 2)
		
	End Function
	
	Public Function GetTotalPropulsionCost() As Double
		'gets the cost of all propulsion systems.  Used by ManeuverControl cost calculations
		Dim TempCost As Double
		Dim element As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case GasolineEngine, HPGasolineEngine, TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, SuperHPGasolineEngine, StandardDieselEngine, TurboStandardDieselEngine, MarineDieselEngine, HPDieselEngine, TurboHPDieselEngine, CeramicEngine, TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, TurboHPCeramicEngine, SuperHPCeramicEngine, HydrogenCombustionEngine, EarlySteamEngine, ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine, StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, StandardMHDTurbine, HPMHDTurbine, FuelCell, FissionReactor, RTGReactor, NPU, FusionReactor, AntimatterReactor, TotalConversionPowerPlant, CosmicPowerPlant, Soulburner, ElementalFurnace, ManaEngine, Carnivore, Herbivore, Omnivore, Vampire, ClockWork, LeadAcidBattery, AdvancedBattery, Flywheel, RechargeablePowerCell, PowerCell, NitrousOxideBooster
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Cost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempCost = TempCost + element.Cost
					
			End Select
		Next element
		
		GetTotalPropulsionCost = TempCost
	End Function
	
	Public Function GetOriginalManeuverControlCost(ByVal ControlType As Short) As Double
		'finds the cost of an existing maneuver control
		Dim TempCost As Double
		Dim element As Object
		
		For	Each element In Veh
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If element.Datatype = ControlType Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Duplicate. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.Duplicate = False Then
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Cost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempCost = element.Cost
					Exit For
				End If
			End If
		Next element
		
		GetOriginalManeuverControlCost = TempCost
	End Function
	
	Public Function CalcCombinedVolume(ByVal ObjectKey As String) As Double
		Dim element As Object
		Dim NumObjects As Short
		Dim locationkey() As String
		Dim lkey As String
		Dim dType As Short
		
		On Error Resume Next
		
		
		' add up volumes of all objects attached to the subassembly
		For	Each element In Veh
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = element.Datatype
			If TypeOf element Is clsHarness Then
			ElseIf TypeOf element Is clsArmor Then 
			ElseIf TypeOf element Is clsSoftware Then 
			ElseIf TypeOf element Is clsWeaponLink Then 
			ElseIf TypeOf element Is clsBattlesuitSystem Then 
			ElseIf TypeOf element Is clsManeuverControl Then 
			ElseIf dType = Module_Renamed Then  'the ModularSocket holds the volume!
			ElseIf dType = SolarCellArray Then 
				'i know for instance that ALL other sub assemblies are not to be included in Combined Component volumes!!!!
			ElseIf (dType = Turret) Or (dType = Popturret) Or (dType = Wheel) Or (dType = Gasbag) Or (dType = Wing) Or (dType = Leg) Or (dType = Skid) Or (dType = Track) Or (dType = AutogyroRotor) Or (dType = TTRotor) Or (dType = CARotor) Or (dType = MMRotor) Or (dType = Hydrofoil) Or (dType = Hovercraft) Or (dType = Superstructure) Or (dType = Pod) Or (dType = Arm) Then 
				
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lkey = element.LogicalParent
				If lkey <> "" Then
					If lkey = ObjectKey Then
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Volume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CalcCombinedVolume = CalcCombinedVolume + element.Volume
					End If
				End If
			End If
		Next element
		
	End Function
	Public Function CalcRotationSpace(ByVal sKey As String) As Double
		Dim element As Object
		Dim tempSpace As Double
		
		' add up volumes of all objects attached to the subassembly
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If (TypeOf element Is clsTurret) Or (TypeOf element Is clsPopTurret) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.LogicalParent = sKey Then
					'important to make sure we are getting the latest stats here
					'UPGRADE_WARNING: Couldn't resolve default property of object element.StatsUpdate. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					element.StatsUpdate()
					'UPGRADE_WARNING: Couldn't resolve default property of object element.RotationSpace. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					tempSpace = tempSpace + element.RotationSpace
				End If
			End If
		Next element
		CalcRotationSpace = System.Math.Round(tempSpace, 2)
	End Function
	
	Public Function CalcSlopeMultiplier(ByVal sKey As String) As Single
		Dim Temp As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Temp = Val(Veh.Components(sKey).SlopeF) + Val(Veh.Components(sKey).SlopeB) + Val(Veh.Components(sKey).SlopeL) + Val(Veh.Components(sKey).SlopeR)
		
		Select Case Temp
			Case 0
				CalcSlopeMultiplier = 1
			Case 30
				CalcSlopeMultiplier = 1.1
			Case 60
				CalcSlopeMultiplier = 1.25
			Case 90
				CalcSlopeMultiplier = 1.4
			Case 120
				CalcSlopeMultiplier = 1.6
			Case 150
				CalcSlopeMultiplier = 2
			Case 180
				CalcSlopeMultiplier = 2.5
			Case 210
				CalcSlopeMultiplier = 3.3
			Case 240
				CalcSlopeMultiplier = 5
		End Select
		
	End Function
	
	
	Public Function CalcTotalContragravLift() As Double
		Dim element As Object
		Dim Templift As Double
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If TypeOf element Is clsContraGravGenerator Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Lift. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Templift = element.Lift
			End If
		Next element
		CalcTotalContragravLift = Templift
	End Function
	
	
	Public Function CalcTotalDeckArea(ByVal Area As Single, ByVal MyKey As String) As Single
		Dim temparea As Single
		Dim element As Object
		Dim AppendageArea As Double
		Dim Modifier As Boolean
		Dim dType As Short
		
		temparea = Area / 3
		Modifier = False 'init
		
		' First check for streamlining of "good" or better
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Veh.surface.StreamLining = "none") Or (Veh.surface.StreamLining = "fair") Then
		Else
			Modifier = True
		End If
		
		For	Each element In Veh
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = element.Datatype
			'check for streamlining, gasbag, and masts
			If Modifier Then
			Else
				If (dType = Mast) Or (dType = Gasbag) Then
					Modifier = True
				End If
			End If
			'get areas of any attached turret, superstructures or Open cargo
			'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If element.LogicalParent = MyKey Then
				If (dType = Turret) Or (dType = Popturret) Or (dType = Superstructure) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SurfaceArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AppendageArea = AppendageArea + (element.SurfaceArea / 5)
				ElseIf dType = Cargo Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SubType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.SubType = "open" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Volume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Debug.Print("CalcTotalDeckArea: " & element.Volume)
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Volume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AppendageArea = AppendageArea + (element.Volume / 10)
					End If
				End If
			End If
		Next element
		
		temparea = temparea + AppendageArea
		
		'If vehicle has Masts,gasbags or streamlining of good or better, divide this by 2
		If Modifier Then temparea = temparea / 2
		If temparea < 0 Then temparea = 0
		
		CalcTotalDeckArea = temparea
		
	End Function
	
	Public Function CalcDeckCost(ByVal FlightArea As Single, ByVal CoveredArea As Single, ByVal DeckOption As String) As Single
		Dim Modifier As Single
		Dim FlightCost As Single
		Dim CoveredCost As Single
		
		If DeckOption = "landing pad" Then
			Modifier = 0.5
		ElseIf DeckOption = "angled flight deck" Then 
			Modifier = 2
		Else
			Modifier = 1
		End If
		
		FlightCost = FlightArea * 1 * Modifier
		CoveredCost = CoveredArea * 0.5
		
		CalcDeckCost = System.Math.Round(FlightCost + CoveredCost, 2)
		
	End Function
	
	
	Public Function CalcDeckWeight(ByVal FlightArea As Single, ByVal CoveredArea As Single, ByVal DeckOption As String) As Single
		Dim Modifier As Single
		Dim FlightWeight As Single
		Dim CoveredWeight As Single
		
		If DeckOption = "landing pad" Then
			Modifier = 0.5
		Else
			Modifier = 1
		End If
		
		FlightWeight = FlightArea * 0.1 * Modifier
		CoveredWeight = CoveredArea * 0.05
		CalcDeckWeight = System.Math.Round(FlightWeight + CoveredWeight, 2)
		
	End Function
	
	Function GetRetractLocation() As String
		Dim retval As String
		Dim j As Integer
		Dim SubKeys() As String
		
		On Error Resume Next 'MPJ 07/04/2000 ' added resume next to avoid rare (and unresolved) problem where subassembly key is no longer valid
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SubKeys = VB6.CopyArray(Veh.KeyManager.GetCurrentSubAssembliesKeys)
		retval = "none"
		If SubKeys(1) <> "" Then
			For j = 1 To UBound(SubKeys)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (Veh.Components(SubKeys(j)).Datatype = Wheel) Or (Veh.Components(SubKeys(j)).Datatype = Skid) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					retval = Veh.Components(SubKeys(j)).RetractLocation
					Exit For
				End If
			Next 
		End If
		'todo: this must take the "worst" value.  For instance, what happens
		'if you have one set of skids set to retract into body and
		'then a set of wheels set to retract into body and wing?
		GetRetractLocation = retval
	End Function
	
	Public Function GetScanRating(ByVal Range As Single) As Integer
		Dim Scan As Single
		Dim Base As Single
		Dim I As Integer

		Scan = 0
		Base = 5
		I = 0
		
		Do While Scan = 0
			If Range <= 0.1 Then 'the only exception
				Scan = Base
				Exit Do
			End If
			
			If Range < 0.15 * 10 ^ I Then
				Scan = Base + (6 * I)
			ElseIf Range < 0.2 * 10 ^ I Then 
				Scan = Base + 1 + (6 * I)
			ElseIf Range < 0.3 * 10 ^ I Then 
				Scan = Base + 2 + (6 * I)
			ElseIf Range < 0.45 * 10 ^ I Then 
				Scan = Base + 3 + (6 * I)
			ElseIf Range < 0.7 * 10 ^ I Then 
				Scan = Base + 4 + (6 * I)
			ElseIf Range < 1 * 10 ^ I Then 
				Scan = Base + 5 + (6 * I)
			End If
			I = I + 1
		Loop 
		GetScanRating = Scan
	End Function
	
	
	Public Function CalcSurfaceArea(ByVal Volume As Single) As Double
		Dim Temp As Short
		Dim I As Integer
		If Volume = 0 Then Exit Function
		
		Temp = UBound(SurfaceAreaMatrix)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Veh.Options.UseSurfaceAreaTable) And (Volume < SurfaceAreaMatrix(Temp).Volume) Then
			
			If Volume < SurfaceAreaMatrix(1).Volume Then
				CalcSurfaceArea = SurfaceAreaMatrix(1).Area
				Exit Function
			End If
			
			For I = 2 To Temp
				If Volume < SurfaceAreaMatrix(I).Volume Then
					CalcSurfaceArea = SurfaceAreaMatrix(I).Area
					Exit Function
				End If
			Next 
		Else
			
			CalcSurfaceArea = System.Math.Round(((Volume ^ (1 / 3)) ^ 2) * 6, 2)
		End If
	End Function
	
	Public Function TechLevelModifier(ByVal TechLevel As Short) As Object
		
		Select Case TechLevel
			Case Is <= 5
				'UPGRADE_WARNING: Couldn't resolve default property of object TechLevelModifier. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TechLevelModifier = 12
			Case 6
				'UPGRADE_WARNING: Couldn't resolve default property of object TechLevelModifier. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TechLevelModifier = 8
			Case 7
				'UPGRADE_WARNING: Couldn't resolve default property of object TechLevelModifier. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TechLevelModifier = 6
			Case 8
				'UPGRADE_WARNING: Couldn't resolve default property of object TechLevelModifier. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TechLevelModifier = 4
			Case 9
				'UPGRADE_WARNING: Couldn't resolve default property of object TechLevelModifier. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TechLevelModifier = 3
			Case 10
				'UPGRADE_WARNING: Couldn't resolve default property of object TechLevelModifier. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TechLevelModifier = 2
			Case 11
				'UPGRADE_WARNING: Couldn't resolve default property of object TechLevelModifier. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TechLevelModifier = 1.5
			Case Is >= 12
				'UPGRADE_WARNING: Couldn't resolve default property of object TechLevelModifier. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TechLevelModifier = 1
		End Select
	End Function
	
	
	
	
	Public Function BasicDesignCost(ByVal SurfaceArea As Single, ByVal StructureTL As Short, ByVal sStrength As String, ByVal sMaterials As String, ByVal bResponsive As Boolean, ByVal bRobotic As Boolean, ByVal bBiomechanical As Boolean, ByVal bLivingMetal As Boolean) As Double
		'This produces structural COST and not BASIC cost!!!
		Dim StructureCost As Short ' Basic Design Cost
		Dim StrengthCost As Single ' Frame Strength Cost Multiplier
		Dim MaterialsCost As Single ' Materials Cost Multiplier
		Dim StreamLiningCost As Single ' Streamlined Structure Cost Multiplier
		Dim TotalSpecialCost As Single ' total of all Special Structure Cost Modifiers
		Dim TotalOtherCost As Single ' total of all Other Cost Modifiers
		'UPGRADE_ISSUE: clsSurface object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oSurface As clsSurface ' holds the body class for the vehicle
		Dim sStreamlining As String ' vehicle body's streamlining property
		Dim bSubmersible As Boolean ' vehicle body's submersible property
		Dim bWingsorRotors As Boolean ' vehicle body's wingsorRotors property
		Dim bLiftingBody As Boolean ' vehicle body's LiftingBody property
		Dim bFlexibody As Boolean ' vehcicle body's Flexibody option
		
		Dim element As Object
		
		Const ResponsiveCost As Double = 1.5 ' Responsive Structure Cost Multiplier
		Const RoboticCost As Short = 2 ' Robotic Structure Cost Multiplier
		Const BiomechanicalCost As Double = 1.5 ' Biomechanical Structure Cost Multiplier
		Const LivingMetalCost As Short = 2 ' Living Metal Structure Cost Multiplier
		Const SubmersibleCost As Short = 2 ' Submersible Structure Cost Multiplier
		Const WingsorRotorsCost As Short = 10 ' Wings or Rotors Cost Multiplier
		Const LiftingBodyCost As Double = 1.2 ' Lifting Body Cost Multiplier
		Const FlexibodyDriveCost As Short = 2 ' Flexibody Drive Train Cost Multiplier
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oSurface = Veh.surface
		
		With oSurface
			'UPGRADE_WARNING: Couldn't resolve default property of object oSurface.StreamLining. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sStreamlining = .StreamLining
			'UPGRADE_WARNING: Couldn't resolve default property of object oSurface.Submersible. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			bSubmersible = .Submersible
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			bLiftingBody = Veh.Components(BODY_KEY).LiftingBody
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			bFlexibody = Veh.Components(BODY_KEY).FlexibodyOption
		End With
		
		'Determine if wings or Rotors are on the vehicle
		If VehicleHasWings Or VehicleHasRotors Then
			bWingsorRotors = True
		End If
		
		Select Case StructureTL
			Case 0, 1
				StructureCost = 5
			Case 2, 3, 4
				StructureCost = 5
			Case 5
				StructureCost = 5
			Case 6
				StructureCost = 10
			Case 7
				StructureCost = 50
			Case 8
				StructureCost = 50
			Case 9
				StructureCost = 50
			Case 10
				StructureCost = 50
			Case 11
				StructureCost = 50
			Case Is >= 12
				StructureCost = 50
			Case Else
				StructureCost = 1
		End Select
		
		Select Case sStrength
			Case "super-light"
				StrengthCost = 0.1
			Case "extra-light"
				StrengthCost = 0.25
			Case "light"
				StrengthCost = 0.5
			Case "medium"
				StrengthCost = 1
			Case "heavy"
				StrengthCost = 2
			Case "extra-heavy"
				StrengthCost = 5
			Case Else
				StrengthCost = 1
		End Select
		
		Select Case sMaterials
			Case "very cheap"
				MaterialsCost = 0.2
			Case "cheap"
				MaterialsCost = 0.5
			Case "standard"
				MaterialsCost = 1
			Case "expensive"
				MaterialsCost = 2
			Case "very expensive"
				MaterialsCost = 5
			Case "advanced"
				MaterialsCost = 10
			Case Else
				MaterialsCost = 1
		End Select
		
		Select Case sStreamlining
			Case "none"
				StreamLiningCost = 1
			Case "fair"
				StreamLiningCost = 1.2
			Case "good"
				StreamLiningCost = 1.5
			Case "very good"
				StreamLiningCost = 2
			Case "superior"
				StreamLiningCost = 3
			Case "excellent"
				StreamLiningCost = 5
			Case "radical"
				StreamLiningCost = 10
			Case Else
				StreamLiningCost = 1
		End Select
		
		' Calculate total cost modifier value for Special Structures
		TotalSpecialCost = 1 ' initialize variable
		If bResponsive Then TotalSpecialCost = ResponsiveCost
		If bRobotic Then TotalSpecialCost = TotalSpecialCost * RoboticCost
		If bBiomechanical Then TotalSpecialCost = TotalSpecialCost * BiomechanicalCost
		If bLivingMetal Then TotalSpecialCost = TotalSpecialCost * LivingMetalCost
		
		' Calculate total cost modifier value for Other Modifiers
		TotalOtherCost = 1 ' initialize variable
		If bSubmersible Then TotalOtherCost = 2
		If bWingsorRotors Then TotalOtherCost = TotalOtherCost * 10
		If bLiftingBody Then TotalOtherCost = TotalOtherCost * 1.2
		If bFlexibody Then TotalOtherCost = TotalOtherCost * 2
		
		' Calculate Final Structural Cost
		BasicDesignCost = System.Math.Round(SurfaceArea * StructureCost * StrengthCost * MaterialsCost * StreamLiningCost * TotalSpecialCost * TotalOtherCost, 2)
	End Function
	
	Public Function BasicDesignWeight(ByVal SurfaceArea As Single, ByVal StructureTL As Short, ByVal sStrength As String, ByVal sMaterials As String) As Double
		'This produces the STRUCTURAL WEIGHT and NOT the BASIC DESIGN WEIGHT
		
		Dim StructureWeight As Single ' Basic Design Weight
		Dim StrengthWeight As Single ' Frame Strength Weight Multiplier
		Dim MaterialsWeight As Single ' Materials Weight Multiplier
		Dim TotalOtherWeight As Single ' Total value of Other Weight Multipliers
		
		Dim bSubmersible As Boolean
		Dim bFlexibody As Boolean
		
		Const SubmersibleCost As Short = 2 ' Submerisble Structure Weight Multiplier
		Const FlexibodyCost As Short = 2 ' Flexibody Drivetrain Weight Multiplier
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bSubmersible = Veh.surface.Submersible
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bFlexibody = Veh.Components(BODY_KEY).FlexibodyOption
		
		Select Case StructureTL
			Case 0, 1
				StructureWeight = 20
			Case 2, 3, 4
				StructureWeight = 18
			Case 5
				StructureWeight = 12
			Case 6
				StructureWeight = 8
			Case 7
				StructureWeight = 6
			Case 8
				StructureWeight = 4
			Case 9
				StructureWeight = 3
			Case 10
				StructureWeight = 2
			Case 11
				StructureWeight = 1.5
			Case Is >= 12
				StructureWeight = 1
		End Select
		
		Select Case sStrength
			Case "super-light"
				StrengthWeight = 0.1
			Case "extra-light"
				StrengthWeight = 0.25
			Case "light"
				StrengthWeight = 0.5
			Case "medium"
				StrengthWeight = 1
			Case "heavy"
				StrengthWeight = 1.5
			Case "extra-heavy"
				StrengthWeight = 2
		End Select
		
		Select Case sMaterials
			Case "very cheap"
				MaterialsWeight = 2
			Case "cheap"
				MaterialsWeight = 1.5
			Case "standard"
				MaterialsWeight = 1
			Case "expensive"
				MaterialsWeight = 0.75
			Case "very expensive"
				MaterialsWeight = 0.5
			Case "advanced"
				MaterialsWeight = 0.375
		End Select
		
		' Calculate total weight modifier value for Other Modifiers
		TotalOtherWeight = 1 ' initialize variable
		If bSubmersible Then TotalOtherWeight = SubmersibleCost
		If bFlexibody Then TotalOtherWeight = TotalOtherWeight * FlexibodyCost
		
		' Calculate Final Structural Weight
		BasicDesignWeight = System.Math.Round(SurfaceArea * StructureWeight * StrengthWeight * MaterialsWeight * TotalOtherWeight, 2)
	End Function
	
	
	Public Function CalcComponentHitpoints(ByVal nNum As Double) As Object
		'//functin calculates Hitpoints for all NON subassemblies
		Dim nTemp As Double
		
		nTemp = System.Math.Round(nNum, 0)
		If nTemp < 1 Then nTemp = 1
		
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalcComponentHitpoints = nTemp
		
	End Function
	
	'UPGRADE_WARNING: ParamArray NumberofWheelsTracksSkids was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Public Function CalcHitPoints(ByVal SubAssembly As String, ByVal FrameStrength As String, ByVal Area As Single, ParamArray ByVal NumberofWheelsTracksSkids() As Object) As Double
		'//This function calculates the hitpoints for Subassemblies only
		Dim TempHitPoints As Double
		
		Select Case SubAssembly
			Case "clsArm", "clsRotor"
				TempHitPoints = Area * 3
			Case "clsBody", "clsSuperStructure", "clsTurret", "clsPopTurret", "clsPod", "clsLeg", "clsWing"
				TempHitPoints = Area * 1.5
			Case "clsSkid", "clsTrack"
				TempHitPoints = (Area * 1.5) / (NumberofWheelsTracksSkids(0))
			Case "clsGasbag"
				TempHitPoints = Area * 0.01
			Case "clsMast", "clsOpenMount"
				TempHitPoints = Area * 2
			Case "clsWheel"
				'UPGRADE_WARNING: Couldn't resolve default property of object NumberofWheelsTracksSkids(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TempHitPoints = (Area * 3) / NumberofWheelsTracksSkids(0)
			Case "clsHovercraft"
				TempHitPoints = Area * 1.5
			Case "clsHydrofoil"
				TempHitPoints = Area * 1.5
			Case "clsSolarPanel"
				TempHitPoints = Area * 0.2
		End Select
		' /////////////////////////
		' Note: I removed the gasbag, open mount and mast sections and placed them directly in the
		' classes for those subassemblies.
		' /////////////////////////
		' Continue to calculate hitpoints for all other types of subassemblies
		Select Case FrameStrength
			Case "super-light"
				TempHitPoints = TempHitPoints / 10
			Case "extra-light"
				TempHitPoints = TempHitPoints / 4
			Case "light"
				TempHitPoints = TempHitPoints / 2
			Case "heavy"
				TempHitPoints = TempHitPoints * 2
			Case "extra-heavy"
				TempHitPoints = TempHitPoints * 4
		End Select
		' produce a rounded (whole) number
		CalcHitPoints = System.Math.Round(TempHitPoints, 0)
		' make sure that we always have at least 1 hit point... can't have 0
		If CalcHitPoints < 1 Then
			CalcHitPoints = 1
		End If
		
	End Function
End Module