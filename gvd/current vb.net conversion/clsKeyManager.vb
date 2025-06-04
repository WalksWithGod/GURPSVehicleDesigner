Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsKeyManager_NET.clsKeyManager")> Public Class clsKeyManager
	
	'///////////////////////////////////////////////////////
	'Keychains 'make sure all keychains start with 1
	Private mvarPowerSystemKeyChain As Object
	Private mvarPowerConsumptionKeyChain As Object
	
	Private mvarFuelUsingSystemKeyChain As Object
	Private mvarFuelStorageKeyChain As Object
	
	Private mvarLegsKeychain As Object
	
	Private mvarLegDrivetrainKeychain As Object
	Private mvarRotorDrivetrainKeychain As Object
	Private mvarOrnithopterDrivetrainKeychain As Object
	Private mvarOtherGroundDrivetrainKeychain As Object
	Private mvarRotorsKeychain As Object
	
	Private mvarSubAssembliesKeychain As Object
	
	
	
	
	'MPJ 03/30/02 OBSOLETE -- There are no more dependancies of Performance Profiles on Motive Subassemblies
	'Sub RemoveAnyDependantPerformanceProfiles(Key As String)
	'Dim profilearray() As String
	'Dim i As Long
	'
	'profilearray = Veh.Components(BODY_KEY).GetCurrentPerformanceProfileKeys
	'If profilearray(1) = "" Then
	'        Exit Sub
	'Else
	'    For i = 1 To UBound(profilearray)
	'        If Veh.Components(profilearray(i)).MotiveAssemblyKey = Key Then
	'            Veh.Components(BODY_KEY).RemovePerformanceProfileKey profilearray(i) 'remove the profile from the ProfileKeychain
	'            Veh.Components.Remove profilearray(i) 'remove the profile from the collection
	'        End If
	'    Next
	'End If
	'
	'End Sub
	Sub RemoveKeyChainKey(ByRef Key As String, ByRef Datatype As Short)
		Dim profilearray() As String
		Dim weaponlinkarray() As String
		Dim k As Integer
		Dim temparray() As String ' holds keys for Weapon in WeaponLInk
		Dim i As Integer
		
		Select Case Datatype
			'remove subassembly key references
			Case Leg
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemoveLegKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemoveSubAssemblyKey(Key)
				
			Case Wheel, Skid, Track, Arm, Hydrofoil, Hovercraft, AutogyroRotor, Wing, Mast, Superstructure, Turret, Popturret, OpenMount, Gasbag, Pod
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemoveSubAssemblyKey(Key)
				
			Case TTRotor, CARotor, MMRotor
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemoveRotorKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemoveSubAssemblyKey(Key)
				
			Case LegDrivetrain
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemoveLegDrivetrainKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemovePowerConsumptionKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveConsumerFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
				
			Case OrnithopterDrivetrain
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemoveOrnithopterDrivetrainKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemovePowerConsumptionKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveConsumerFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
				
			Case CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemoveRotorDrivetrainKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemovePowerConsumptionKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveConsumerFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
				
			Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, FlexibodyDrivetrain, TrackedDrivetrain
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemoveOtherGroundDrivetrainKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemovePowerConsumptionKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveConsumerFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
				
				'-----------------------------------------------------------------
				'REMOVE POWER CONSUMPTION KEY REFERENCES
			Case SimpleCustom, MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling, BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam, AerialPropeller, DuctedFan, PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel, MagLevLifter, StandardThruster, SuperThruster, MegaThruster, IonDrive, TeleportationDrive, Hyperdrive, JumpDrive, WarpDrive, QuantumConveyor, SubQuantumConveyor, TwoQuantumConveyor, ContraGravGenerator, RadioDirectionFinder, RadioCommunicator, TightBeamRadio, VLFRadio, CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator, Headlight, Searchlight, InfraredSearchlight, AstronomicalInstruments, Telescope, lightAmplification, LowlightTV
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemovePowerConsumptionKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveConsumerFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
				
			Case Radar, Ladar, NavigationalRadar, AntiCollisionRadar, AESA, LowResImagingRadar, HiResImagingRadar, ActiveSonar, PassiveSonar, PassiveInfrared, Thermograph, PassiveRadar, PESA, Geophone, MAD, MultiScanner, ChemScanner, RadScanner, BioScanner, GravScanner, RangingSoundDetector, SurveillanceSoundDetector, MeteorologicalInstruments, LowResPlanetarySurveyArray, MedResPlanetarySurveyArray, HighResPlanetarySurveyArray, SoundSystem, FlightRecorder, VehicleCamera, DigitalVehicleCamera, ReconCamera, DigitalReconCamera, NavigationInstruments, AutoPilot, IFF, Transponder, INS, GPS, MilitaryGPS, TFR, AdvancedOpticalBombSight, FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, LaserDesignator, LaserSpotTracker, RadarDetector, LaserSensor, LaserRadarDetector, AreaRadarJammer, DeceptiveRadarJammer, InfraredJammer, BlipEnhancer, TEMPEST, MacroFrame, MainFrame, MicroFrame, MiniComputer, SmallComputer, Terminal, SurgicalInterface, InterfaceWeb, AutoInterfaceWeb, SocketInterface, NeuralInductionField, DeflectorField, ForceScreen, VariableForceScreen
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemovePowerConsumptionKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveConsumerFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
				
			Case ArmMotor, BilgePump, CompleteWorkshop, MechanicWorkshop, EngineeringWorkshop, ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, MiniEngineeringWorkshop, MiniArmouryWorkshop, Crane, Winch, PowerShovel, WreckingCrane, ForkLift, VehicularBridge, LaunchCatapult, Bore, SuperBore, TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet, OperatingRoom, EmergencySupportUnit, EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable, Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, MovieScreenandProjectorSmall, HoloventureZone, Airlock, MembraneAirlock, Forcelock, TeleportProjector, BurglarAlarm, HighSecurityAlarm, MutableLicensePlate, RefuellingProbe, FuelElectrolysisSystem, AtmosphereProcessor, NuclearDamper, SmallRealityStabilizer, MediumRealityStabilizer, HeavyRealityStabilizer, TotalLifeSystem, ArtificialGravityUnit, EnvironmentalControl, NBCKit, LimitedLifeSystem, FullLifeSystem, GravityWeb, GravCompensator
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemovePowerConsumptionKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveConsumerFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
				
				'-----------------------------------------------------------------
				'DEVICES WHICH SUPPLY POWER
			Case SimpleCustom, MuscleEngine, ClockWork, LeadAcidBattery, AdvancedBattery, Flywheel, PowerCell, ElectricContactPower, LaserBeamedPowerReceiver, MaserBeamedPowerReceiver, FissionReactor, RTGReactor, NPU, FusionReactor, AntimatterReactor, TotalConversionPowerPlant, CosmicPowerPlant, SolarCellArray, Soulburner, ElementalFurnace, ManaEngine, Carnivore, Herbivore, Omnivore, Vampire
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemovePowerSystemKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveSupplierFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveSupplierFromAllProfiles(Key, POWER_PROFILE)
				
				'-----------------------------------------------------------------
				' RECHARGEABLE POWER SUPPLY DEVICE
			Case RechargeablePowerCell
				' this is the only power supply that also gets configured as a CONSUMER since it can be recharged
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemovePowerSystemKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveSupplierFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveSupplierFromAllProfiles(Key, POWER_PROFILE)
				
				' the extra calls since this is categorized as a consumer also
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemovePowerConsumptionKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveConsumerFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
				
				'-----------------------------------------------------------------
				'DEVICES WHICH SUPPLY POWER BUT ALSO CONSUME FUEL
			Case StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, StandardMHDTurbine, HPMHDTurbine, GasolineEngine, HPGasolineEngine, TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, SuperHPGasolineEngine, StandardDieselEngine, TurboStandardDieselEngine, MarineDieselEngine, HPDieselEngine, TurboHPDieselEngine, CeramicEngine, TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, TurboHPCeramicEngine, SuperHPCeramicEngine, HydrogenCombustionEngine, EarlySteamEngine, ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine, FuelCell, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemovePowerSystemKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemoveFuelUsingSystemKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveSupplierFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveSupplierFromAllProfiles(Key, POWER_PROFILE)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveConsumerFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveConsumerFromAllProfiles(Key, FUEL_PROFILE)
				
				'-----------------------------------------------------------------
				' DEVICES WHICH CONSUME FUEL
			Case LiquidFuelRocket, MOXRocket, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemoveFuelUsingSystemKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveConsumerFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveConsumerFromAllProfiles(Key, FUEL_PROFILE)
				
				'-----------------------------------------------------------------
				'DEVICES WHICH SUPPLY FUEL
			Case AntiMatterBay, CoalBunker, WoodBunker, StandardTank, lightTank, UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, UltralightSelfSealingTank
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Veh.KeyManager.RemoveFuelStorageKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.RemoveSupplierFromAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.RemoveSupplierFromAllProfiles(Key, FUEL_PROFILE)
		End Select
		
		'/////now remove the key of this deleted propulsion system from every Performance Profile
		Dim O As Object
		Dim oWL As clsWeaponLink
		Select Case Datatype
			Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, FlexibodyDrivetrain, TrackedDrivetrain, LegDrivetrain, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, AerialPropeller, DuctedFan, PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, StandardThruster, SuperThruster, MegaThruster, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, SolidRocketEngine, OrionEngine, TeleportationDrive, Hyperdrive, JumpDrive, WarpDrive, QuantumConveyor, SubQuantumConveyor, TwoQuantumConveyor, RowingPositions, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, lightSail, ContraGravGenerator, MagLevLifter
				
				'profilearray = Veh.KeyManager.GetCurrentPerformanceProfileKeys
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For	Each O In Veh.PerformanceProfiles
					'UPGRADE_WARNING: Couldn't resolve default property of object O.GetCurrentKeys. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					temparray = VB6.CopyArray(O.GetCurrentKeys)
					For i = 1 To UBound(temparray)
						If temparray(i) = Key Then
							'UPGRADE_WARNING: Couldn't resolve default property of object O.RemoveKey. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							O.RemoveKey(Key)
							Exit For
						End If
					Next 
				Next O
				
				
				'            For k = 1 To UBound(profilearray)
				'                If profilearray(1) = "" Then
				'                    Exit For
				'                End If
				'                temparray = Veh.PerformanceProfiles(profilearray(k)).GetCurrentKeys
				'                For i = 1 To UBound(temparray)
				'                    If temparray(i) = Key Then
				'                        Veh.PerformanceProfiles(profilearray(k)).RemoveKey Key
				'                        Exit For
				'                    End If
				'                Next
				'            Next
				
				'// Remove the key of this deleted weapon component from every weapon link that might be referencing it
			Case StoneThrower, BoltThrower, RepeatingBoltThrower, MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling, BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam, FlameThrower, WaterCannon, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, RevolverLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher, IronBomb, RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, ProximityMine, PressureTriggerMine, CommandTriggerMine, SmartTriggerMine, UnGuidedMissile, UnGuidedTorpedo, GuidedMissile, GuidedTorpedo
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.WeaponProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For	Each oWL In Veh.WeaponProfiles
					temparray = VB6.CopyArray(oWL.GetCurrentKeys)
					For i = 1 To UBound(temparray)
						If temparray(i) = Key Then
							oWL.RemoveKey(Key)
							Exit For
						End If
					Next 
				Next oWL
				
				'            For k = 1 To UBound(weaponlinkarray)
				'                If weaponlinkarray(1) = "" Then
				'                    Exit For
				'                End If
				'                temparray = Veh.WeaponProfiles(weaponlinkarray(k)).GetCurrentKeys
				'                For i = 1 To UBound(temparray)
				'                    If temparray(i) = Key Then
				'                        Veh.Components(weaponlinkarray(k)).RemoveKey Key
				'                        Exit For
				'                    End If
				'                Next
				'            Next
		End Select
	End Sub
	
	
	Public Sub AddKeyChainKeys(ByVal Key As String)
		Dim Datatype As Short
		
		'todo: this will all be deleted i think. The way i want to handle this is to simply
		' check the supported interfaces of a class.  If a class has a cIPowerConsume interface
		' then it gets added to .AddconsumerToAllProfiles()
		' and if that same component also has a cIWeapon interface, it gets added for weapon link configuring purposes
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Datatype = Veh.Components(Key).Datatype
		
		Select Case Datatype
			'add Power Consumption Key Reference
			Case SimpleCustom, MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling, BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam, WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, FlexibodyDrivetrain, TrackedDrivetrain, LegDrivetrain, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, AerialPropeller, DuctedFan, PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel, MagLevLifter, StandardThruster, SuperThruster, MegaThruster, IonDrive, TeleportationDrive, Hyperdrive, JumpDrive, WarpDrive, QuantumConveyor, SubQuantumConveyor, TwoQuantumConveyor, ContraGravGenerator, RadioDirectionFinder, RadioCommunicator, TightBeamRadio, VLFRadio, CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator, Headlight, Searchlight, InfraredSearchlight, AstronomicalInstruments, Telescope, lightAmplification, LowlightTV
				
				Call AddPowerConsumptionKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.AddConsumerToAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.AddConsumerToAllProfiles(Key, POWER_PROFILE)
				
			Case Radar, Ladar, NavigationalRadar, AntiCollisionRadar, AESA, LowResImagingRadar, HiResImagingRadar, ActiveSonar, PassiveSonar, PassiveInfrared, Thermograph, PassiveRadar, PESA, Geophone, MAD, MultiScanner, ChemScanner, RadScanner, BioScanner, GravScanner, RangingSoundDetector, SurveillanceSoundDetector, MeteorologicalInstruments, LowResPlanetarySurveyArray, MedResPlanetarySurveyArray, HighResPlanetarySurveyArray, SoundSystem, FlightRecorder, VehicleCamera, DigitalVehicleCamera, ReconCamera, DigitalReconCamera, NavigationInstruments, AutoPilot, IFF, Transponder, INS, GPS, MilitaryGPS, TFR, AdvancedOpticalBombSight, FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, LaserDesignator, LaserSpotTracker, RadarDetector, LaserSensor, LaserRadarDetector, AreaRadarJammer, DeceptiveRadarJammer, InfraredJammer, BlipEnhancer, TEMPEST, MacroFrame, MainFrame, MicroFrame, MiniComputer, SmallComputer, Terminal, SurgicalInterface, InterfaceWeb, AutoInterfaceWeb, SocketInterface, NeuralInductionField, DeflectorField, ForceScreen, VariableForceScreen
				
				Call AddPowerConsumptionKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.AddConsumerToAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.AddConsumerToAllProfiles(Key, POWER_PROFILE)
				
			Case ArmMotor, BilgePump, CompleteWorkshop, MechanicWorkshop, EngineeringWorkshop, ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, MiniEngineeringWorkshop, MiniArmouryWorkshop, Crane, Winch, PowerShovel, WreckingCrane, ForkLift, VehicularBridge, LaunchCatapult, Bore, SuperBore, TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet, OperatingRoom, EmergencySupportUnit, EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable, Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, MovieScreenandProjectorSmall, HoloventureZone, Airlock, MembraneAirlock, Forcelock, TeleportProjector, BurglarAlarm, HighSecurityAlarm, MutableLicensePlate, RefuellingProbe, FuelElectrolysisSystem, AtmosphereProcessor, NuclearDamper, SmallRealityStabilizer, MediumRealityStabilizer, HeavyRealityStabilizer, TotalLifeSystem, ArtificialGravityUnit, EnvironmentalControl, NBCKit, LimitedLifeSystem, FullLifeSystem, GravityWeb, GravCompensator
				
				Call AddPowerConsumptionKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.AddConsumerToAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.AddConsumerToAllProfiles(Key, POWER_PROFILE)
				
				'add Power System Key Reference
			Case SimpleCustom, MuscleEngine, GasolineEngine, HPGasolineEngine, TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, SuperHPGasolineEngine, StandardDieselEngine, TurboStandardDieselEngine, MarineDieselEngine, HPDieselEngine, TurboHPDieselEngine, CeramicEngine, TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, TurboHPCeramicEngine, SuperHPCeramicEngine, HydrogenCombustionEngine, EarlySteamEngine, ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine, StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, StandardMHDTurbine, HPMHDTurbine, FuelCell, FissionReactor, RTGReactor, NPU, FusionReactor, AntimatterReactor, TotalConversionPowerPlant, CosmicPowerPlant, Soulburner, ElementalFurnace, ManaEngine, Carnivore, Herbivore, Omnivore, Vampire, ClockWork, LeadAcidBattery, AdvancedBattery, Flywheel, PowerCell, ElectricContactPower, LaserBeamedPowerReceiver, MaserBeamedPowerReceiver, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, SolarCellArray
				
				Call AddPowerSystemKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.AddSupplierToAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.AddSupplierToAllProfiles(Key, POWER_PROFILE)
				
			Case RechargeablePowerCell ' 07/13/02 MPJ
				' this rechargeable power supplier also gets added as a consumer
				Call AddPowerConsumptionKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.AddConsumerToAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.AddConsumerToAllProfiles(Key, POWER_PROFILE)
				Call AddPowerSystemKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.AddSupplierToAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.AddSupplierToAllProfiles(Key, POWER_PROFILE)
		End Select
		
		'todo: the following items even though they are Power Generating Systems
		'can also be recharged when assigned to another power generating system
		'ClockWork , LeadAcidBattery, AdvancedBattery, Flywheel, _
		''RechargeablePowerCell , PowerCell
		
		'/////// Now add any keychain references for FuelUsingSystems
		
		Select Case Datatype
			Case StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, StandardMHDTurbine, HPMHDTurbine, GasolineEngine, HPGasolineEngine, TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, SuperHPGasolineEngine, StandardDieselEngine, TurboStandardDieselEngine, MarineDieselEngine, HPDieselEngine, TurboHPDieselEngine, CeramicEngine, TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, TurboHPCeramicEngine, SuperHPCeramicEngine, HydrogenCombustionEngine, EarlySteamEngine, ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FuelCell
				
				Call AddFuelUsingSystemKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.AddConsumerToAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.AddConsumerToAllProfiles(Key, FUEL_PROFILE)
				
		End Select
		'///////Now add any keychain references for Fuel Storage components
		Select Case Datatype
			Case AntiMatterBay, CoalBunker, WoodBunker, StandardTank, lightTank, UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, UltralightSelfSealingTank
				
				Call AddFuelStorageKey(Key)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.AddSupplierToAllProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.AddSupplierToAllProfiles(Key, FUEL_PROFILE)
				
		End Select
		
		'/////////Now add any keychain references for Drivetrains
		Select Case Datatype
			'remove keychain references for Wheeled drivetrains, Tracked Drivetrains, Flexibody Drivetrains
			Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, FlexibodyDrivetrain, TrackedDrivetrain
				
				Call AddOtherGroundDrivetrainKey(Key)
				
			Case MMRRotorDrivetrain, TTRRotorDrivetrain, CARRotorDrivetrain
				
				Call AddRotorDrivetrainKey(Key)
				
			Case LegDrivetrain
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.KeyManager.AddLegDrivetrainKey(Key)
				
			Case OrnithopterDrivetrain
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.KeyManager.AddOrnithopterDrivetrainKey(Key)
				
			Case CARotor, TTRotor, MMRotor
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.KeyManager.AddRotorKey(Key)
				
		End Select
		'/////////Now for assemblies
		'if its a subassembly, add it to the subassembly keychain
		Select Case Datatype
			Case Wheel, Skid, Track, Leg, Arm, Hydrofoil, Hovercraft, AutogyroRotor, TTRotor, CARotor, MMRotor, Wing, Mast, Superstructure, Turret, Popturret, OpenMount, Gasbag, Pod
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call Veh.KeyManager.AddSubAssemblyKey(Key)
		End Select
		
		'if its a leg, _also_ add it to the leg keychain
		If Datatype = Leg Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call Veh.KeyManager.AddLegKey(Key)
		End If
	End Sub
	
	'//////////////////Subassemblies keychain management
	Public Function GetCurrentSubAssembliesKeys() As String()
		GetCurrentSubAssembliesKeys = VariantArrayToStringArray(mvarSubAssembliesKeychain)
	End Function
	
	Public Sub AddSubAssemblyKey(ByRef subassemblykey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarSubAssembliesKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarSubAssembliesKeychain = mAddKey(mvarSubAssembliesKeychain, subassemblykey)
	End Sub
	
	Public Sub RemoveSubAssemblyKey(ByRef subassemblykey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarSubAssembliesKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarSubAssembliesKeychain = mRemoveKey(mvarSubAssembliesKeychain, subassemblykey)
	End Sub
	
	
	'//////////////////LegsKeychain management
	Public Function GetCurrentLegKeys() As String()
		GetCurrentLegKeys = VariantArrayToStringArray(mvarLegsKeychain)
	End Function
	
	Public Sub AddLegKey(ByRef legkey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarLegsKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarLegsKeychain = mAddKey(mvarLegsKeychain, legkey)
	End Sub
	
	Public Sub RemoveLegKey(ByRef legkey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarLegsKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarLegsKeychain = mRemoveKey(mvarLegsKeychain, legkey)
	End Sub
	
	
	'//////////////////Leg Drivetrain Keychain Management
	Public Function GetCurrentLegDrivetrainKeys() As String()
		GetCurrentLegDrivetrainKeys = VariantArrayToStringArray(mvarLegDrivetrainKeychain)
	End Function
	
	Public Sub AddLegDrivetrainKey(ByRef legdrivetrainkey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarLegDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarLegDrivetrainKeychain = mAddKey(mvarLegDrivetrainKeychain, legdrivetrainkey)
	End Sub
	
	Public Sub RemoveLegDrivetrainKey(ByRef legdrivetrainkey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarLegDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarLegDrivetrainKeychain = mRemoveKey(mvarLegDrivetrainKeychain, legdrivetrainkey)
	End Sub
	
	'//////////////////Rotorkeys management
	Public Function GetCurrentRotorKeys() As String()
		GetCurrentRotorKeys = VariantArrayToStringArray(mvarRotorsKeychain)
	End Function
	
	Public Sub AddRotorKey(ByRef Rotorkey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRotorsKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarRotorsKeychain = mAddKey(mvarRotorsKeychain, Rotorkey)
	End Sub
	
	Public Sub RemoveRotorKey(ByRef Rotorkey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRotorsKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarRotorsKeychain = mRemoveKey(mvarRotorsKeychain, Rotorkey)
	End Sub
	
	
	'//////////////////Helicopter Drivetrain Keychain Management
	Public Function GetCurrentRotorDrivetrainKeys() As String()
		GetCurrentRotorDrivetrainKeys = VariantArrayToStringArray(RotorDrivetrainKeychain)
	End Function
	
	Public Sub AddRotorDrivetrainKey(ByRef rotordrivetrainkey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRotorDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarRotorDrivetrainKeychain = mAddKey(mvarRotorDrivetrainKeychain, rotordrivetrainkey)
	End Sub
	
	Public Sub RemoveRotorDrivetrainKey(ByRef rotordrivetrainkey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRotorDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarRotorDrivetrainKeychain = mRemoveKey(mvarRotorDrivetrainKeychain, rotordrivetrainkey)
	End Sub
	
	'//////////////////Ornithopter Drivetrain Keychain Management
	Public Function GetCurrentOrnithopterDrivetrainKeys() As String()
		GetCurrentOrnithopterDrivetrainKeys = VariantArrayToStringArray(mvarOrnithopterDrivetrainKeychain)
	End Function
	
	Public Sub AddOrnithopterDrivetrainKey(ByRef Ornithopterdrivetrainkey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarOrnithopterDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarOrnithopterDrivetrainKeychain = mAddKey(mvarOrnithopterDrivetrainKeychain, Ornithopterdrivetrainkey)
	End Sub
	
	Public Sub RemoveOrnithopterDrivetrainKey(ByRef Ornithopterdrivetrainkey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarOrnithopterDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarOrnithopterDrivetrainKeychain = mRemoveKey(mvarOrnithopterDrivetrainKeychain, Ornithopterdrivetrainkey)
	End Sub
	
	
	'//////////////////Other Ground Drivetrain Keychain Management
	Public Function GetCurrentOtherGroundDrivetrainKeys() As String()
		GetCurrentOtherGroundDrivetrainKeys = VariantArrayToStringArray(mvarOtherGroundDrivetrainKeychain)
	End Function
	
	Public Sub AddOtherGroundDrivetrainKey(ByRef OtherGroundDrivetrainkey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarOtherGroundDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarOtherGroundDrivetrainKeychain = mAddKey(mvarOtherGroundDrivetrainKeychain, OtherGroundDrivetrainkey)
	End Sub
	
	Public Sub RemoveOtherGroundDrivetrainKey(ByRef OtherGroundDrivetrainkey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarOtherGroundDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarOtherGroundDrivetrainKeychain = mRemoveKey(mvarOtherGroundDrivetrainKeychain, OtherGroundDrivetrainkey)
	End Sub
	
	'//////////////////PerformanceProfileKeychain management
	'MPJ 07/09/02  OBSOLETE -- Performance Profiles are now in a seperate collection so we dont need to track them via keychains
	'Public Function GetCurrentPerformanceProfileKeys() As String()
	'    GetCurrentPerformanceProfileKeys = VariantArrayToStringArray(mvarperformanceprofilekeychain)
	'End Function
	'
	'Public Sub AddPerformanceProfileKey(PerformanceProfileKey As String)
	'    mvarperformanceprofilekeychain = mAddKey(mvarperformanceprofilekeychain, PerformanceProfileKey)
	'End Sub
	'
	'Public Sub RemovePerformanceProfileKey(PerformanceProfileKey As String)
	'    mvarperformanceprofilekeychain = mRemoveKey(mvarperformanceprofilekeychain, PerformanceProfileKey)
	'End Sub
	
	
	'//////////////////PowerSystemKeychain management
	Public Function GetCurrentPowerSystemKeys() As String()
		GetCurrentPowerSystemKeys = VariantArrayToStringArray(mvarPowerSystemKeyChain)
	End Function
	
	Public Sub AddPowerSystemKey(ByRef PowerSystemKey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPowerSystemKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarPowerSystemKeyChain = mAddKey(mvarPowerSystemKeyChain, PowerSystemKey)
	End Sub
	
	Public Sub RemovePowerSystemKey(ByRef PowerSystemKey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPowerSystemKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarPowerSystemKeyChain = mRemoveKey(mvarPowerSystemKeyChain, PowerSystemKey)
	End Sub
	
	
	'//////////////////PowerConsumptionKeyChain management
	Public Function GetCurrentPowerConsumptionKeys() As String()
		GetCurrentPowerConsumptionKeys = VariantArrayToStringArray(mvarPowerConsumptionKeyChain)
	End Function
	
	Public Sub AddPowerConsumptionKey(ByRef PowerConsumptionKey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPowerConsumptionKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarPowerConsumptionKeyChain = mAddKey(mvarPowerConsumptionKeyChain, PowerConsumptionKey)
	End Sub
	
	Public Sub RemovePowerConsumptionKey(ByRef PowerConsumptionKey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPowerConsumptionKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarPowerConsumptionKeyChain = mRemoveKey(mvarPowerConsumptionKeyChain, PowerConsumptionKey)
	End Sub
	
	'//////////////////FuelUsingSystemKeychain management
	Public Function GetCurrentFuelUsingSystemKeys() As String()
		GetCurrentFuelUsingSystemKeys = VariantArrayToStringArray(mvarFuelUsingSystemKeyChain)
	End Function
	
	Public Sub AddFuelUsingSystemKey(ByRef FuelUsingSystemKey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarFuelUsingSystemKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarFuelUsingSystemKeyChain = mAddKey(mvarFuelUsingSystemKeyChain, FuelUsingSystemKey)
	End Sub
	
	Public Sub RemoveFuelUsingSystemKey(ByRef FuelUsingSystemKey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarFuelUsingSystemKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarFuelUsingSystemKeyChain = mRemoveKey(mvarFuelUsingSystemKeyChain, FuelUsingSystemKey)
	End Sub
	
	
	'//////////////////FuelStorageKeyChain management
	Public Function GetCurrentFuelStorageKeys() As String()
		GetCurrentFuelStorageKeys = VariantArrayToStringArray(mvarFuelStorageKeyChain)
	End Function
	
	Public Sub AddFuelStorageKey(ByRef FuelStorageKey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarFuelStorageKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarFuelStorageKeyChain = mAddKey(mvarFuelStorageKeyChain, FuelStorageKey)
	End Sub
	
	Public Sub RemoveFuelStorageKey(ByRef FuelStorageKey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarFuelStorageKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarFuelStorageKeyChain = mRemoveKey(mvarFuelStorageKeyChain, FuelStorageKey)
	End Sub
	
	
	
	Public Property PowerSystemKeyChain() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPowerSystemKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PowerSystemKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PowerSystemKeyChain = mvarPowerSystemKeyChain
		End Get
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object vdata. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPowerSystemKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPowerSystemKeyChain = Value
		End Set
	End Property
	
	
	Public Property PowerConsumptionKeyChain() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPowerConsumptionKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PowerConsumptionKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PowerConsumptionKeyChain = mvarPowerConsumptionKeyChain
		End Get
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object vdata. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPowerConsumptionKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPowerConsumptionKeyChain = Value
		End Set
	End Property
	
	
	
	Public Property FuelUsingSystemKeyChain() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarFuelUsingSystemKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FuelUsingSystemKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FuelUsingSystemKeyChain = mvarFuelUsingSystemKeyChain
		End Get
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object vdata. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarFuelUsingSystemKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarFuelUsingSystemKeyChain = Value
		End Set
	End Property
	
	
	Public Property FuelStorageKeyChain() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarFuelStorageKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object FuelStorageKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FuelStorageKeyChain = mvarFuelStorageKeyChain
		End Get
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object vdata. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarFuelStorageKeyChain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarFuelStorageKeyChain = Value
		End Set
	End Property
	
	
	
	Public Property SubAssembliesKeychain() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSubAssembliesKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object SubAssembliesKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SubAssembliesKeychain = mvarSubAssembliesKeychain
		End Get
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object vdata. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSubAssembliesKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarSubAssembliesKeychain = Value
		End Set
	End Property
	
	
	Public Property LegsKeychain() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarLegsKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object LegsKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LegsKeychain = mvarLegsKeychain
		End Get
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object vdata. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarLegsKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarLegsKeychain = Value
		End Set
	End Property
	
	
	Public Property LegDrivetrainKeychain() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarLegDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object LegDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LegDrivetrainKeychain = mvarLegDrivetrainKeychain
		End Get
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object vdata. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarLegDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarLegDrivetrainKeychain = Value
		End Set
	End Property
	
	
	Public Property RotorDrivetrainKeychain() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarRotorDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object RotorDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RotorDrivetrainKeychain = mvarRotorDrivetrainKeychain
		End Get
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object vdata. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarRotorDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarRotorDrivetrainKeychain = Value
		End Set
	End Property
	
	
	Public Property OrnithopterDrivetrainKeychain() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarOrnithopterDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object OrnithopterDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OrnithopterDrivetrainKeychain = mvarOrnithopterDrivetrainKeychain
		End Get
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object vdata. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarOrnithopterDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarOrnithopterDrivetrainKeychain = Value
		End Set
	End Property
	
	
	Public Property OtherGroundDrivetrainKeychain() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarOtherGroundDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object OtherGroundDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OtherGroundDrivetrainKeychain = mvarOtherGroundDrivetrainKeychain
		End Get
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object vdata. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarOtherGroundDrivetrainKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarOtherGroundDrivetrainKeychain = Value
		End Set
	End Property
	
	
	Public Property RotorsKeychain() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarRotorsKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object RotorsKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RotorsKeychain = mvarRotorsKeychain
		End Get
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object vdata. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarRotorsKeychain. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarRotorsKeychain = Value
		End Set
	End Property
	
	'=================================================================================================
	'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'=================================================================================================
	''/////////These next routines are used to hold the
	''///////// keys of the power systems they are drawing power from
	'Public Function GetCurrentPowerSystemKeys() As String()
	'GetCurrentPowerSystemKeys = VariantArrayToStringArray(mvarPowerSystemKeyChain)
	'End Function
	'
	'Public Sub AddPowerSystemKey(PowerSystemKey As String)
	'mvarPowerSystemKeyChain = mAddKey(mvarPowerSystemKeyChain, PowerSystemKey)
	'End Sub
	'
	'Public Sub RemovePowerSystemKey(PowerSystemKey As String)
	'mvarPowerSystemKeyChain = mRemoveKey(mvarPowerSystemKeyChain, PowerSystemKey)
	'End Sub
	'
	'Public Sub AddPowerConsumptionKeyReference()
	''When this object is created, it needs to add its OWN key
	''to the powerconsumptionkeychain in the hull.
	'Veh.Components(BODY_KEY).AddPowerConsumptionKey mvarKey
	'End Sub
	'
	'Public Sub RemoveAllKeyReferences()
	''When this object is deleted from the Vehicle,
	''its OWN key needs to be removed from the body's PowerConsumptionkeychain which
	''tracks all power consuming components in the vehicle
	'Veh.Components(BODY_KEY).RemovePowerConsumptionKey mvarKey
	'
	''remove its key from the power systems that are providing power to this component
	'mRemoveReferencedKeys mvarPowerSystemKeyChain, mvarKey, "RemovePowerConsumptionKey"
	'
	'End Sub
	'
	'
	'
	''/////////These next routines are used to hold the
	''///////// keys of all components that are consuming power from this power system
	'Public Function GetCurrentConsumptionSystemKeys() As String()
	'GetCurrentConsumptionSystemKeys = VariantArrayToStringArray(mvarPowerConsumptionKeyChain)
	'End Function
	'
	'Public Sub AddPowerConsumptionKey(ConsumptionKey As String)
	'mvarPowerConsumptionKeyChain = mAddKey(mvarPowerConsumptionKeyChain, ConsumptionKey)
	'End Sub
	'
	'Public Sub RemovePowerConsumptionKey(ConsumptionKey As String)
	'mvarPowerConsumptionKeyChain = mRemoveKey(mvarPowerConsumptionKeyChain, ConsumptionKey)
	'End Sub
	'
	''//////When this object is created, it needs to add its OWN key
	''////// to the powersystemkeychain in the hull.  When this object
	''////// is deleted from the Vehicle either during build process
	''////// or if its destroyed in battle or removed to be traded
	''////// its OWN key needs to be removed from the body's powersystemkeychain
	'Public Sub AddPowerSystemKeyReference()
	'Veh.Components(BODY_KEY).AddPowerSystemKey mvarKey
	'End Sub
	'
	'
	'Public Sub RemoveAllKeyReferences()
	'
	'Veh.Components(BODY_KEY).RemovePowerSystemKey mvarKey
	'
	'mRemoveReferencedKeys mvarPowerConsumptionKeyChain, mvarKey, "RemovePowerSystemKey"
	'End Sub
	'
	''/////////These next routines are used to hold the
	''///////// keys of all components that are consuming power from this power system
	'Public Function GetCurrentConsumptionSystemKeys() As String()
	'GetCurrentConsumptionSystemKeys = VariantArrayToStringArray(mvarPowerConsumptionKeyChain)
	'End Function
	'
	'Public Sub AddPowerConsumptionKey(ConsumptionKey As String)
	'mvarPowerConsumptionKeyChain = mAddKey(mvarPowerConsumptionKeyChain, ConsumptionKey)
	'End Sub
	'
	'Public Sub RemovePowerConsumptionKey(ConsumptionKey As String)
	'mvarPowerConsumptionKeyChain = mRemoveKey(mvarPowerConsumptionKeyChain, ConsumptionKey)
	'End Sub
	'
	'
	''/////////These next routines are used to hold the
	''///////// keys of all fuel storage devices that are linked to this
	'Public Function GetCurrentFuelStorageKeys() As String()
	'GetCurrentFuelStorageKeys = VariantArrayToStringArray(mvarFuelStorageKeyChain)
	'End Function
	'
	'Public Sub AddFuelStorageKey(FuelStorageKey As String)
	'mvarFuelStorageKeyChain = mAddKey(mvarFuelStorageKeyChain, FuelStorageKey)
	'StatsUpdate
	'End Sub
	'
	'Public Sub RemoveFuelStorageKey(FuelStorageKey As String)
	'mvarFuelStorageKeyChain = mRemoveKey(mvarFuelStorageKeyChain, FuelStorageKey)
	'StatsUpdate
	'End Sub
	'
	''//////When this object is created, it needs to add its OWN key
	''////// to the powersystemkeychain in the hull.  When this object
	''////// is deleted from the Vehicle either during build process
	''////// or if its destroyed in battle or removed to be traded
	''////// its OWN key needs to be removed from the body's powersystemkeychain
	'Public Sub AddPowerSystemKeyReference()
	'Veh.Components(BODY_KEY).AddPowerSystemKey mvarKey
	'End Sub
	'
	''same with regard to its fuel using
	'Public Sub AddFuelUsingSystemKeyReference()
	'Veh.Components(BODY_KEY).AddFuelUsingSystemKey mvarKey
	'End Sub
	'
	'Public Sub RemoveAllKeyReferences()
	'Veh.Components(BODY_KEY).RemoveFuelUsingSystemKey mvarKey
	'
	'Veh.Components(BODY_KEY).RemovePowerSystemKey mvarKey
	'
	'mRemoveReferencedKeys mvarFuelStorageKeyChain, mvarKey, "RemoveFuelUsingSystemKey"
	'
	'mRemoveReferencedKeys mvarPowerConsumptionKeyChain, mvarKey, "RemovePowerSystemKey"
	'End Sub
	'
	'Public Sub RemoveAllKeyReferences()
	''When this object is deleted from the Vehicle,
	''its OWN key needs to be removed from the body's PowerConsumptionkeychain which
	''tracks all power consuming components in the vehicle
	'Veh.Components(BODY_KEY).RemovePowerConsumptionKey mvarKey
	'
	''remove its key from the power systems that are providing power to this component
	'mRemoveReferencedKeys mvarPowerSystemKeyChain, mvarKey, "RemovePowerConsumptionKey"
	'
	''/////////Now remove any keychain references for Drivetrains and Assemblies
	'Select Case Datatype
	' 'remove keychain references for Wheeled drivetrains, Tracked Drivetrains, Flexibody Drivetrains
	'    Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, FlexibodyDrivetrain, TrackedDrivetrain
	'
	'        Veh.Components(BODY_KEY).RemoveOtherGroundDrivetrainKey Key
	'
	'
	'    Case LegDrivetrain
	'
	'        Veh.Components(BODY_KEY).RemoveLegDrivetrainKey Key
	'
	'    Case OrnithopterDrivetrain
	'        Veh.Components(BODY_KEY).RemoveOrnithopterDrivetrainKey Key
	'
	'
	'
	'End Select
	'End Sub
	'
	'
	'Public Sub RemoveAllKeyReferences()
	'
	'
	'
	'
	'If mvarDatatype <> FusionAirRam Then ' Fusion Air Ram's dont use Fuel, they use internal material
	'    Veh.Components(BODY_KEY).RemoveFuelUsingSystemKey mvarKey
	'    mRemoveReferencedKeys mvarFuelStorageKeyChain, mvarKey, "RemoveFuelUsingSystemKey"
	'End If
	'
	'' remove our key from the Body's keychain which tracks
	'' all power consuming devices
	'Veh.Components(BODY_KEY).RemovePowerSystemKey mvarKey
	'' call function which will have our key removed from their
	'' keychains. (e.g. all Power Plants will have our key removed as
	'' an item which is consumin power from them)
	'mRemoveReferencedKeys mvarPowerConsumptionKeyChain, mvarKey, "RemovePowerSystemKey"
	'End Sub
End Class