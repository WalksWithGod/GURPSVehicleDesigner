Option Strict Off
Option Explicit On
Module modGUI_Performance
	
	Private p_CheckListKeys() As String
	Private m_sMappedKeys() As String ' since our listbox only holds indexes, we map this array by index to determine the key of the listitem
	Private m_sCurrent As String
	Private m_lngCheckListType As Integer
	
	Public Sub PopulateWeaponLinkCheckList(ByVal sCurrent As String)
		Dim m_oCurrentVeh As Object
		Dim element As Object
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case StoneThrower, BoltThrower, RepeatingBoltThrower, MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling, BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam, FlameThrower, WaterCannon, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, RevolverLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher, IronBomb, RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, ProximityMine, PressureTriggerMine, CommandTriggerMine, SmartTriggerMine, UnGuidedMissile, UnGuidedTorpedo, GuidedMissile, GuidedTorpedo
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddCheckListKey(element.Key, element.CustomDescription)
					
					'check mark existing
					Call CheckMarkExisting()
			End Select
		Next element
	End Sub
	
	Public Sub PopulateCheckList(ByVal sCurrent As String)
		Dim m_oCurrentVeh As Object
		Dim frmDesigner As Object
		'clear the list
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.lstPropulsionSystems. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.lstPropulsionSystems.Clear()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.ActiveCheckList. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sCurrent = m_oCurrentVeh.ActiveCheckList
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.ActiveCheckListType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_lngCheckListType = m_oCurrentVeh.ActiveCheckListType
		
		ReDim p_CheckListKeys(1)
		ReDim m_sMappedKeys(0)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.ActiveCheckListType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		System.Diagnostics.Debug.Assert(m_oCurrentVeh.ActiveCheckListType > 0, "")
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.ActiveCheckListType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If m_oCurrentVeh.ActiveCheckListType = PERFORMANCE_CHECKLIST Then
			Call PopulatePerformanceCheckList(sCurrent)
		Else
			Call PopulateWeaponLinkCheckList(sCurrent)
		End If
		
	End Sub
	Public Sub PopulatePerformanceCheckList(ByVal sCurrent As String)
		Dim m_oCurrentVeh As Object
		
		
		'todo: all this crap needs to be converted to check bitflag for propulsion system capabilities
		'fill the global propulsion system list based on Type of Performance Profile
		Dim sPType As String 'performance type
		Dim element As Object
		Dim sMType As String 'holds classname for the Ground MotiveAssembly
		
		
		' cycle through all objects in the Vehicle (NOTE: to optimize, a creating a keychain in the AddObject
		' routine to track propulsion systems would be helpful) and come up with list of propulsion systems that
		' are relevant to the performance profile type
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'//Wheels - only display elements which can be used with Wheeled Performance
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsPerformanceWheel Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
						
						'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call AddCheckListKey(element.Key, element.CustomDescription)
						
				End Select
				'//Skids
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsPerformanceSkid Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					Case DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
						
						'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call AddCheckListKey(element.Key, element.CustomDescription)
				End Select
				'//Tracks, Halftracks, Skitracks
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsPerformanceTrack Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					Case TrackedDrivetrain
						'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call AddCheckListKey(element.Key, element.CustomDescription)
				End Select
				
				'//Legs
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsPerformanceLeg Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					Case LegDrivetrain
						'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call AddCheckListKey(element.Key, element.CustomDescription)
				End Select
				
				'//Flexibody
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsPerformanceFlex Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					Case FlexibodyDrivetrain
						'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call AddCheckListKey(element.Key, element.CustomDescription)
				End Select
				
				'//WATER
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsPerformanceWater Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					Case TrackedDrivetrain, LegDrivetrain, FlexibodyDrivetrain, WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, RowingPositions, PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
						
						'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call AddCheckListKey(element.Key, element.CustomDescription)
				End Select
				
				'//Submerged
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsPerformanceSubmerged Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					Case TrackedDrivetrain, LegDrivetrain, FlexibodyDrivetrain, PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
						
						'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call AddCheckListKey(element.Key, element.CustomDescription)
				End Select
				
				'//Aerial Performance
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsPerformanceAir Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					Case ContraGravGenerator, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine
						
						'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call AddCheckListKey(element.Key, element.CustomDescription)
				End Select
				'//Mag-Lev Performance
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsPerformanceMagLev Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					Case ContraGravGenerator, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine, MagLevLifter
						
						'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call AddCheckListKey(element.Key, element.CustomDescription)
				End Select
				'//Hovercraft Performance
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsPerformanceHover Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					'note Ramjet's are not allowed since max speed for Hovercrafts is 300mph
					Case ContraGravGenerator, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine
						
						'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call AddCheckListKey(element.Key, element.CustomDescription)
				End Select
				'//Space Performance
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsPerformanceSpace Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					Case TeleportationDrive, Hyperdrive, JumpDrive, WarpDrive, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine, lightSail
						
						'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call AddCheckListKey(element.Key, element.CustomDescription)
				End Select
			End If
		Next element
		
		'check mark existing
		Call CheckMarkExisting()
		
	End Sub
	
	
	Private Sub CheckMarkExisting()
		Dim m_oCurrentVeh As Object
		Dim frmDesigner As Object
		
		'Check Mark any items that are already added
		Dim arrKeys() As String
		Dim i As Integer
		Dim j As Integer
		
		On Error GoTo err_Renamed
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.lstPropulsionSystems. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.lstPropulsionSystems.TAG = CHECKLIST_STATE_RESTORE
		
		If m_lngCheckListType = WEAPON_CHECKLIST Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.WeaponProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			arrKeys = m_oCurrentVeh.WeaponProfiles(m_sCurrent).getcurrentkeys
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			arrKeys = m_oCurrentVeh.PerformanceProfiles(m_sCurrent).getcurrentkeys
		End If
		
		If arrKeys(1) = "" Then
		Else
			For i = 1 To UBound(arrKeys)
				' if the keys inside the performance profile match any on the list, checkmark them
				' since they are already added
				For j = 0 To UBound(m_sMappedKeys)
					If arrKeys(i) = m_sMappedKeys(j) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.lstPropulsionSystems. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						frmDesigner.lstPropulsionSystems.Selected(j) = True
						Exit For
					End If
					'frmDesigner.lstPropulsionSystems.AddItem m_oCurrentVeh.Components(arrKeys(i)).customDescription, arrKeys(i)
				Next 
			Next 
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.lstPropulsionSystems. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.lstPropulsionSystems.TAG = ""
		Exit Sub
err_Renamed: 
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.lstPropulsionSystems. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.lstPropulsionSystems.TAG = ""
	End Sub
	Private Sub AddCheckListKey(ByRef sKey As String, ByRef sDescription As String)
		Dim frmDesigner As Object
		Dim Count As Integer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.lstPropulsionSystems. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.lstPropulsionSystems.AddItem(sDescription)
		p_CheckListKeys = mAddKey(p_CheckListKeys, sKey)
		
		Count = UBound(m_sMappedKeys)
		If (Count = 0) And (m_sMappedKeys(0) = "") Then
			Count = 0
		Else
			Count = Count + 1
		End If
		
		ReDim Preserve m_sMappedKeys(Count)
		
		m_sMappedKeys(Count) = sKey
		
	End Sub
	Public Sub UpdatePerformanceStats()
		Dim m_oCurrentVeh As Object
		' when a user adds/removes propulsion/drivetrain components from a profile, the stats
		' need to be adjusted in real time.
		Dim o As Object
		
		' re-calc vehicle performance figures
		'note: to optimize, i would only update figures for those profiles
		'which the user changed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each o In m_oCurrentVeh.PerformanceProfiles
			'UPGRADE_WARNING: Couldn't resolve default property of object o.CalcPerformance. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			o.CalcPerformance()
		Next o
		
		
	End Sub
	
	Public Sub PropulsionSelect(ByRef sCurrentProfile As String, ByRef iIndex As Integer)
		Dim m_oCurrentVeh As Object
		'add the item to the keychain of the Current Performance Profile (see clsPerformanceXXXXXX)
		If m_lngCheckListType = WEAPON_CHECKLIST Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.WeaponProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_oCurrentVeh.WeaponProfiles(m_sCurrent).AddKey(p_CheckListKeys(iIndex + 1))
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_oCurrentVeh.PerformanceProfiles(m_sCurrent).AddKey(p_CheckListKeys(iIndex + 1))
		End If
	End Sub
	
	Public Sub PropulsionDeSelect(ByRef sCurrentProfile As String, ByRef iIndex As Integer)
		Dim m_oCurrentVeh As Object
		'remove the item from the keychain of the Performance Profile
		If m_lngCheckListType = WEAPON_CHECKLIST Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.WeaponProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_oCurrentVeh.WeaponProfiles(m_sCurrent).removekey(p_CheckListKeys(iIndex + 1))
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_oCurrentVeh.PerformanceProfiles(m_sCurrent).removekey(p_CheckListKeys(iIndex + 1))
		End If
	End Sub
	
	Public Sub DeletePerformanceProfile()
		Dim m_oCurrentVeh As Object
		Dim frmDesigner As Object
		
		Dim sItem As String
		On Error GoTo errorhandler
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sItem = frmDesigner.treeVehicle.SelectedItem.Text
		
		If MsgBox("Are you sure you want to delete the profile '" & sItem & "' ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_oCurrentVeh.PerformanceProfiles.Remove(sItem) ' remove the item from the collection
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.treeVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmDesigner.treeVehicle.Nodes.Remove(sItem)
			p_bChangedFlag = True ' JAW 2000.05.07
		End If
		Exit Sub
errorhandler: 
		Exit Sub
	End Sub
End Module