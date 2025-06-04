VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'///////////////////////////////////////////////////////
'Keychains
Option Base 1  'make sure all keychains start with 1
Private mvarPowerSystemKeyChain As Variant
Private mvarPowerConsumptionKeyChain As Variant

Private mvarFuelUsingSystemKeyChain As Variant
Private mvarFuelStorageKeyChain As Variant

Private mvarLegsKeychain As Variant

Private mvarLegDrivetrainKeychain As Variant
Private mvarRotorDrivetrainKeychain As Variant
Private mvarOrnithopterDrivetrainKeychain As Variant
Private mvarOtherGroundDrivetrainKeychain As Variant
Private mvarRotorsKeychain As Variant

Private mvarSubAssembliesKeychain As Variant



    
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
Sub RemoveKeyChainKey(Key As String, Datatype As Integer)
    Dim profilearray() As String
    Dim weaponlinkarray() As String
    Dim k As Long
    Dim temparray() As String ' holds keys for Weapon in WeaponLInk
    Dim i As Long
    
    Select Case Datatype
        'remove subassembly key references
        Case Leg
    
            Veh.KeyManager.RemoveLegKey Key
            Veh.KeyManager.RemoveSubAssemblyKey Key
        
        Case Wheel, Skid, Track, Arm, Hydrofoil, Hovercraft, AutogyroRotor, _
            Wing, Mast, Superstructure, Turret, Popturret, OpenMount, Gasbag, Pod
            Veh.KeyManager.RemoveSubAssemblyKey Key
        
        Case TTRotor, CARotor, MMRotor
            Veh.KeyManager.RemoveRotorKey Key
            Veh.KeyManager.RemoveSubAssemblyKey Key
            
        Case LegDrivetrain
            Veh.KeyManager.RemoveLegDrivetrainKey (Key)
            Veh.KeyManager.RemovePowerConsumptionKey (Key)
            Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
            
        Case OrnithopterDrivetrain
            Veh.KeyManager.RemoveOrnithopterDrivetrainKey (Key)
            Veh.KeyManager.RemovePowerConsumptionKey (Key)
            Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
            
        Case CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain
            Veh.KeyManager.RemoveRotorDrivetrainKey (Key)
            Veh.KeyManager.RemovePowerConsumptionKey (Key)
            Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
        
        Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, FlexibodyDrivetrain, _
            TrackedDrivetrain
    
            Veh.KeyManager.RemoveOtherGroundDrivetrainKey (Key)
            Veh.KeyManager.RemovePowerConsumptionKey (Key)
            Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
                    
        '-----------------------------------------------------------------
        'REMOVE POWER CONSUMPTION KEY REFERENCES
        Case SimpleCustom, MuzzleLoader, BreechLoader, ManualRepeater, _
            Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, _
            lightAutomatic, HeavyAutomatic, ElectricGatling, _
            BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, _
            ChargedParticleBeam, NeutralParticleBeam, _
            Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, _
            GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, _
            BeamedPowerTransmitter, MilitaryParalysisBeam, _
            AerialPropeller, DuctedFan, PaddleWheel, ScrewPropeller, lightScrewPropeller, _
            DuctedPropeller, Hydrojet, MHDTunnel, _
            MagLevLifter, StandardThruster, SuperThruster, MegaThruster, _
            IonDrive, TeleportationDrive, Hyperdrive, JumpDrive, WarpDrive, _
            QuantumConveyor, SubQuantumConveyor, TwoQuantumConveyor, _
            ContraGravGenerator, _
            RadioDirectionFinder, RadioCommunicator, TightBeamRadio, VLFRadio, _
            CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, _
            LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator, _
            Headlight, Searchlight, InfraredSearchlight, _
            AstronomicalInstruments, Telescope, lightAmplification, LowlightTV
            
            Veh.KeyManager.RemovePowerConsumptionKey (Key)
            Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
           
         Case Radar, Ladar, NavigationalRadar, AntiCollisionRadar, AESA, _
            LowResImagingRadar, HiResImagingRadar, ActiveSonar, PassiveSonar, _
            PassiveInfrared, Thermograph, PassiveRadar, PESA, _
            Geophone, MAD, MultiScanner, ChemScanner, RadScanner, BioScanner, _
            GravScanner, RangingSoundDetector, SurveillanceSoundDetector, _
            MeteorologicalInstruments, LowResPlanetarySurveyArray, _
            MedResPlanetarySurveyArray, HighResPlanetarySurveyArray, _
            SoundSystem, FlightRecorder, VehicleCamera, DigitalVehicleCamera, _
            ReconCamera, DigitalReconCamera, NavigationInstruments, AutoPilot, _
            IFF, Transponder, INS, GPS, MilitaryGPS, TFR, _
            AdvancedOpticalBombSight, _
            FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, _
            LaserDesignator, LaserSpotTracker, RadarDetector, LaserSensor, _
            LaserRadarDetector, AreaRadarJammer, DeceptiveRadarJammer, _
            InfraredJammer, BlipEnhancer, TEMPEST, MacroFrame, MainFrame, _
            MicroFrame, MiniComputer, SmallComputer, Terminal, SurgicalInterface, _
            InterfaceWeb, AutoInterfaceWeb, SocketInterface, NeuralInductionField, _
            DeflectorField, ForceScreen, VariableForceScreen
            
            Veh.KeyManager.RemovePowerConsumptionKey (Key)
            Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
              
        Case ArmMotor, _
            BilgePump, CompleteWorkshop, MechanicWorkshop, EngineeringWorkshop, _
            ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, _
            ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, _
            MiniEngineeringWorkshop, MiniArmouryWorkshop, _
            Crane, Winch, PowerShovel, WreckingCrane, ForkLift, _
            VehicularBridge, LaunchCatapult, Bore, SuperBore, _
            TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet, _
            OperatingRoom, EmergencySupportUnit, _
            EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable, _
            Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, _
            MovieScreenandProjectorSmall, HoloventureZone, _
            Airlock, MembraneAirlock, Forcelock, TeleportProjector, _
            BurglarAlarm, HighSecurityAlarm, MutableLicensePlate, RefuellingProbe, _
            FuelElectrolysisSystem, AtmosphereProcessor, NuclearDamper, _
            SmallRealityStabilizer, MediumRealityStabilizer, HeavyRealityStabilizer, _
            TotalLifeSystem, ArtificialGravityUnit, EnvironmentalControl, _
            NBCKit, LimitedLifeSystem, FullLifeSystem, GravityWeb, GravCompensator
            
            Veh.KeyManager.RemovePowerConsumptionKey (Key)
            Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
             
        '-----------------------------------------------------------------
        'DEVICES WHICH SUPPLY POWER
        Case SimpleCustom, MuscleEngine, ClockWork, LeadAcidBattery, _
            AdvancedBattery, Flywheel, PowerCell, _
            ElectricContactPower, LaserBeamedPowerReceiver, MaserBeamedPowerReceiver, _
            FissionReactor, _
            RTGReactor, NPU, FusionReactor, AntimatterReactor, _
            TotalConversionPowerPlant, CosmicPowerPlant, SolarCellArray, _
            Soulburner, ElementalFurnace, ManaEngine, Carnivore, _
            Herbivore, Omnivore, Vampire
         
            Veh.KeyManager.RemovePowerSystemKey Key
            Call Veh.RemoveSupplierFromAllProfiles(Key, POWER_PROFILE)
            
        '-----------------------------------------------------------------
        ' RECHARGEABLE POWER SUPPLY DEVICE
        Case RechargeablePowerCell
            ' this is the only power supply that also gets configured as a CONSUMER since it can be recharged
            Veh.KeyManager.RemovePowerSystemKey Key
            Call Veh.RemoveSupplierFromAllProfiles(Key, POWER_PROFILE)
        
            ' the extra calls since this is categorized as a consumer also
            Veh.KeyManager.RemovePowerConsumptionKey (Key)
            Call Veh.RemoveConsumerFromAllProfiles(Key, POWER_PROFILE)
            
        '-----------------------------------------------------------------
        'DEVICES WHICH SUPPLY POWER BUT ALSO CONSUME FUEL
        Case StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, _
            StandardMHDTurbine, HPMHDTurbine, GasolineEngine, HPGasolineEngine, _
            TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, _
            SuperHPGasolineEngine, StandardDieselEngine, _
            TurboStandardDieselEngine, MarineDieselEngine, HPDieselEngine, _
            TurboHPDieselEngine, CeramicEngine, TurboCeramicEngine, _
            SuperCeramicEngine, HPCeramicEngine, TurboHPCeramicEngine, _
            SuperHPCeramicEngine, HydrogenCombustionEngine, EarlySteamEngine, _
            ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine, _
            FuelCell, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan
            
            Veh.KeyManager.RemovePowerSystemKey Key
            Veh.KeyManager.RemoveFuelUsingSystemKey Key
            Call Veh.RemoveSupplierFromAllProfiles(Key, POWER_PROFILE)
            Call Veh.RemoveConsumerFromAllProfiles(Key, FUEL_PROFILE)
        
         '-----------------------------------------------------------------
        ' DEVICES WHICH CONSUME FUEL
             Case LiquidFuelRocket, MOXRocket, FissionRocket, FusionRocket, _
            OptimizedFusion, AntimatterThermal, AntimatterPion
            
            Veh.KeyManager.RemoveFuelUsingSystemKey (Key)
            Call Veh.RemoveConsumerFromAllProfiles(Key, FUEL_PROFILE)
            
        '-----------------------------------------------------------------
        'DEVICES WHICH SUPPLY FUEL
        Case AntiMatterBay, CoalBunker, WoodBunker, StandardTank, lightTank, _
             UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, _
             UltralightSelfSealingTank
             
            Veh.KeyManager.RemoveFuelStorageKey (Key)
            Call Veh.RemoveSupplierFromAllProfiles(Key, FUEL_PROFILE)
    End Select
    
    '/////now remove the key of this deleted propulsion system from every Performance Profile
    Select Case Datatype
        Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, FlexibodyDrivetrain, _
            TrackedDrivetrain, LegDrivetrain, CARRotorDrivetrain, _
            MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, _
            AerialPropeller, DuctedFan, _
            PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, _
            Hydrojet, MHDTunnel, _
            RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, _
            WhiffletreeHarness, _
            Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, _
            StandardThruster, SuperThruster, MegaThruster, _
            LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
            OptimizedFusion, AntimatterThermal, AntimatterPion, _
            SolidRocketEngine, OrionEngine, _
            TeleportationDrive, Hyperdrive, JumpDrive, WarpDrive, _
            QuantumConveyor, SubQuantumConveyor, TwoQuantumConveyor, _
            RowingPositions, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, lightSail, _
            ContraGravGenerator, MagLevLifter
    
            'profilearray = Veh.KeyManager.GetCurrentPerformanceProfileKeys
            Dim O As Object
            For Each O In Veh.PerformanceProfiles
                temparray = O.GetCurrentKeys
                For i = 1 To UBound(temparray)
                    If temparray(i) = Key Then
                        O.RemoveKey Key
                        Exit For
                    End If
                Next
            Next
            
            
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
        Case StoneThrower, BoltThrower, _
            RepeatingBoltThrower, MuzzleLoader, BreechLoader, ManualRepeater, _
            Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, _
            lightAutomatic, HeavyAutomatic, ElectricGatling, _
            BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, _
            ChargedParticleBeam, NeutralParticleBeam, _
            Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, _
            GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, _
            BeamedPowerTransmitter, MilitaryParalysisBeam, _
            FlameThrower, WaterCannon, DisposableLauncher, _
            MuzzleloadingLauncher, BreechloadingLauncher, _
            ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, _
            RevolverLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher, _
            IronBomb, RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, _
            ProximityMine, PressureTriggerMine, CommandTriggerMine, _
            SmartTriggerMine, _
            UnGuidedMissile, UnGuidedTorpedo, GuidedMissile, GuidedTorpedo
    
            Dim oWL As clsWeaponLink
            For Each oWL In Veh.WeaponProfiles
                temparray = oWL.GetCurrentKeys
                For i = 1 To UBound(temparray)
                    If temparray(i) = Key Then
                        oWL.RemoveKey Key
                        Exit For
                    End If
                Next
            Next
            
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
    Dim Datatype As Integer
    
    'todo: this will all be deleted i think. The way i want to handle this is to simply
    ' check the supported interfaces of a class.  If a class has a cIPowerConsume interface
    ' then it gets added to .AddconsumerToAllProfiles()
    ' and if that same component also has a cIWeapon interface, it gets added for weapon link configuring purposes
    Datatype = Veh.Components(Key).Datatype
    
    Select Case Datatype
        'add Power Consumption Key Reference
        Case SimpleCustom, MuzzleLoader, BreechLoader, ManualRepeater, _
            Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, _
            lightAutomatic, HeavyAutomatic, ElectricGatling, _
            BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, _
            ChargedParticleBeam, NeutralParticleBeam, _
            Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, _
            GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, _
            BeamedPowerTransmitter, MilitaryParalysisBeam, _
            WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, FlexibodyDrivetrain, _
            TrackedDrivetrain, LegDrivetrain, CARRotorDrivetrain, _
            MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, _
            AerialPropeller, DuctedFan, PaddleWheel, ScrewPropeller, lightScrewPropeller, _
            DuctedPropeller, Hydrojet, MHDTunnel, _
            MagLevLifter, StandardThruster, SuperThruster, MegaThruster, _
            IonDrive, TeleportationDrive, Hyperdrive, JumpDrive, WarpDrive, _
            QuantumConveyor, SubQuantumConveyor, TwoQuantumConveyor, _
            ContraGravGenerator, _
            RadioDirectionFinder, RadioCommunicator, TightBeamRadio, VLFRadio, _
            CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, _
            LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator, _
            Headlight, Searchlight, InfraredSearchlight, _
            AstronomicalInstruments, Telescope, lightAmplification, LowlightTV
    
            Call AddPowerConsumptionKey(Key)
            Call Veh.AddConsumerToAllProfiles(Key, POWER_PROFILE)
            
         Case Radar, Ladar, NavigationalRadar, AntiCollisionRadar, AESA, _
            LowResImagingRadar, HiResImagingRadar, ActiveSonar, PassiveSonar, _
            PassiveInfrared, Thermograph, PassiveRadar, PESA, _
            Geophone, MAD, MultiScanner, ChemScanner, RadScanner, BioScanner, _
            GravScanner, RangingSoundDetector, SurveillanceSoundDetector, _
            MeteorologicalInstruments, LowResPlanetarySurveyArray, _
            MedResPlanetarySurveyArray, HighResPlanetarySurveyArray, _
            SoundSystem, FlightRecorder, VehicleCamera, DigitalVehicleCamera, _
            ReconCamera, DigitalReconCamera, NavigationInstruments, AutoPilot, _
            IFF, Transponder, INS, GPS, MilitaryGPS, TFR, _
            AdvancedOpticalBombSight, _
            FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, _
            LaserDesignator, LaserSpotTracker, RadarDetector, LaserSensor, _
            LaserRadarDetector, AreaRadarJammer, DeceptiveRadarJammer, _
            InfraredJammer, BlipEnhancer, TEMPEST, MacroFrame, MainFrame, _
            MicroFrame, MiniComputer, SmallComputer, Terminal, SurgicalInterface, _
            InterfaceWeb, AutoInterfaceWeb, SocketInterface, NeuralInductionField, _
            DeflectorField, ForceScreen, VariableForceScreen
            
            Call AddPowerConsumptionKey(Key)
            Call Veh.AddConsumerToAllProfiles(Key, POWER_PROFILE)
            
        Case ArmMotor, _
            BilgePump, CompleteWorkshop, MechanicWorkshop, EngineeringWorkshop, _
            ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, _
            ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, _
            MiniEngineeringWorkshop, MiniArmouryWorkshop, _
            Crane, Winch, PowerShovel, WreckingCrane, ForkLift, _
            VehicularBridge, LaunchCatapult, Bore, SuperBore, _
            TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet, _
            OperatingRoom, EmergencySupportUnit, _
            EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable, _
            Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, _
            MovieScreenandProjectorSmall, HoloventureZone, _
            Airlock, MembraneAirlock, Forcelock, TeleportProjector, _
            BurglarAlarm, HighSecurityAlarm, MutableLicensePlate, RefuellingProbe, _
            FuelElectrolysisSystem, AtmosphereProcessor, NuclearDamper, _
            SmallRealityStabilizer, MediumRealityStabilizer, HeavyRealityStabilizer, _
            TotalLifeSystem, ArtificialGravityUnit, EnvironmentalControl, _
            NBCKit, LimitedLifeSystem, FullLifeSystem, GravityWeb, GravCompensator
            
            Call AddPowerConsumptionKey(Key)
            Call Veh.AddConsumerToAllProfiles(Key, POWER_PROFILE)
            
        'add Power System Key Reference
        Case SimpleCustom, MuscleEngine, GasolineEngine, HPGasolineEngine, _
            TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, _
            SuperHPGasolineEngine, StandardDieselEngine, _
            TurboStandardDieselEngine, MarineDieselEngine, HPDieselEngine, _
            TurboHPDieselEngine, CeramicEngine, TurboCeramicEngine, _
            SuperCeramicEngine, HPCeramicEngine, TurboHPCeramicEngine, _
            SuperHPCeramicEngine, HydrogenCombustionEngine, EarlySteamEngine, _
            ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine, _
            StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, _
            StandardMHDTurbine, HPMHDTurbine, FuelCell, FissionReactor, _
            RTGReactor, NPU, FusionReactor, AntimatterReactor, _
            TotalConversionPowerPlant, CosmicPowerPlant, _
            Soulburner, ElementalFurnace, ManaEngine, Carnivore, _
            Herbivore, Omnivore, Vampire, ClockWork, LeadAcidBattery, _
            AdvancedBattery, Flywheel, PowerCell, _
            ElectricContactPower, LaserBeamedPowerReceiver, MaserBeamedPowerReceiver, _
            Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, _
            SolarCellArray
            
            Call AddPowerSystemKey(Key)
            Call Veh.AddSupplierToAllProfiles(Key, POWER_PROFILE)
            
        Case RechargeablePowerCell ' 07/13/02 MPJ
            ' this rechargeable power supplier also gets added as a consumer
            Call AddPowerConsumptionKey(Key)
            Call Veh.AddConsumerToAllProfiles(Key, POWER_PROFILE)
            Call AddPowerSystemKey(Key)
            Call Veh.AddSupplierToAllProfiles(Key, POWER_PROFILE)
    End Select
    
        'todo: the following items even though they are Power Generating Systems
        'can also be recharged when assigned to another power generating system
            'ClockWork , LeadAcidBattery, AdvancedBattery, Flywheel, _
            'RechargeablePowerCell , PowerCell
    
    '/////// Now add any keychain references for FuelUsingSystems
    
    Select Case Datatype
        Case StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, _
            StandardMHDTurbine, HPMHDTurbine, GasolineEngine, HPGasolineEngine, _
            TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, _
            SuperHPGasolineEngine, StandardDieselEngine, TurboStandardDieselEngine, _
            MarineDieselEngine, HPDieselEngine, TurboHPDieselEngine, CeramicEngine, _
            TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, _
            TurboHPCeramicEngine, SuperHPCeramicEngine, HydrogenCombustionEngine, _
            EarlySteamEngine, ForcedDraftSteamEngine, TripleExpansionSteamEngine, _
            SteamTurbine, _
            LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
            OptimizedFusion, AntimatterThermal, AntimatterPion, _
            Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FuelCell
        
            Call AddFuelUsingSystemKey(Key)
            Call Veh.AddConsumerToAllProfiles(Key, FUEL_PROFILE)
            
    End Select
    '///////Now add any keychain references for Fuel Storage components
    Select Case Datatype
        Case AntiMatterBay, CoalBunker, WoodBunker, StandardTank, lightTank, _
             UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, _
             UltralightSelfSealingTank
             
            Call AddFuelStorageKey(Key)
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
            
           Call Veh.KeyManager.AddLegDrivetrainKey(Key)
        
        Case OrnithopterDrivetrain
           Call Veh.KeyManager.AddOrnithopterDrivetrainKey(Key)
        
        Case CARotor, TTRotor, MMRotor
           Call Veh.KeyManager.AddRotorKey(Key)
            
    End Select
    '/////////Now for assemblies
     'if its a subassembly, add it to the subassembly keychain
    Select Case Datatype
          Case Wheel, Skid, Track, Leg, Arm, Hydrofoil, Hovercraft, AutogyroRotor, TTRotor, CARotor, _
               MMRotor, Wing, Mast, Superstructure, Turret, Popturret, OpenMount, Gasbag, Pod
                    
              Call Veh.KeyManager.AddSubAssemblyKey(Key)
    End Select
                
    'if its a leg, _also_ add it to the leg keychain
    If Datatype = Leg Then
       Call Veh.KeyManager.AddLegKey(Key)
    End If
End Sub

'//////////////////Subassemblies keychain management
Public Function GetCurrentSubAssembliesKeys() As String()
    GetCurrentSubAssembliesKeys = VariantArrayToStringArray(mvarSubAssembliesKeychain)
End Function

Public Sub AddSubAssemblyKey(subassemblykey As String)
    mvarSubAssembliesKeychain = mAddKey(mvarSubAssembliesKeychain, subassemblykey)
End Sub

Public Sub RemoveSubAssemblyKey(subassemblykey As String)
    mvarSubAssembliesKeychain = mRemoveKey(mvarSubAssembliesKeychain, subassemblykey)
End Sub


'//////////////////LegsKeychain management
Public Function GetCurrentLegKeys() As String()
    GetCurrentLegKeys = VariantArrayToStringArray(mvarLegsKeychain)
End Function

Public Sub AddLegKey(legkey As String)
    mvarLegsKeychain = mAddKey(mvarLegsKeychain, legkey)
End Sub

Public Sub RemoveLegKey(legkey As String)
    mvarLegsKeychain = mRemoveKey(mvarLegsKeychain, legkey)
End Sub


'//////////////////Leg Drivetrain Keychain Management
Public Function GetCurrentLegDrivetrainKeys() As String()
    GetCurrentLegDrivetrainKeys = VariantArrayToStringArray(mvarLegDrivetrainKeychain)
End Function

Public Sub AddLegDrivetrainKey(legdrivetrainkey As String)
    mvarLegDrivetrainKeychain = mAddKey(mvarLegDrivetrainKeychain, legdrivetrainkey)
End Sub

Public Sub RemoveLegDrivetrainKey(legdrivetrainkey As String)
    mvarLegDrivetrainKeychain = mRemoveKey(mvarLegDrivetrainKeychain, legdrivetrainkey)
End Sub

'//////////////////Rotorkeys management
Public Function GetCurrentRotorKeys() As String()
    GetCurrentRotorKeys = VariantArrayToStringArray(mvarRotorsKeychain)
End Function

Public Sub AddRotorKey(Rotorkey As String)
    mvarRotorsKeychain = mAddKey(mvarRotorsKeychain, Rotorkey)
End Sub

Public Sub RemoveRotorKey(Rotorkey As String)
    mvarRotorsKeychain = mRemoveKey(mvarRotorsKeychain, Rotorkey)
End Sub


'//////////////////Helicopter Drivetrain Keychain Management
Public Function GetCurrentRotorDrivetrainKeys() As String()
    GetCurrentRotorDrivetrainKeys = VariantArrayToStringArray(RotorDrivetrainKeychain)
End Function

Public Sub AddRotorDrivetrainKey(rotordrivetrainkey As String)
    mvarRotorDrivetrainKeychain = mAddKey(mvarRotorDrivetrainKeychain, rotordrivetrainkey)
End Sub

Public Sub RemoveRotorDrivetrainKey(rotordrivetrainkey As String)
    mvarRotorDrivetrainKeychain = mRemoveKey(mvarRotorDrivetrainKeychain, rotordrivetrainkey)
End Sub

'//////////////////Ornithopter Drivetrain Keychain Management
Public Function GetCurrentOrnithopterDrivetrainKeys() As String()
    GetCurrentOrnithopterDrivetrainKeys = VariantArrayToStringArray(mvarOrnithopterDrivetrainKeychain)
End Function

Public Sub AddOrnithopterDrivetrainKey(Ornithopterdrivetrainkey As String)
    mvarOrnithopterDrivetrainKeychain = mAddKey(mvarOrnithopterDrivetrainKeychain, Ornithopterdrivetrainkey)
End Sub

Public Sub RemoveOrnithopterDrivetrainKey(Ornithopterdrivetrainkey As String)
    mvarOrnithopterDrivetrainKeychain = mRemoveKey(mvarOrnithopterDrivetrainKeychain, Ornithopterdrivetrainkey)
End Sub


'//////////////////Other Ground Drivetrain Keychain Management
Public Function GetCurrentOtherGroundDrivetrainKeys() As String()
    GetCurrentOtherGroundDrivetrainKeys = VariantArrayToStringArray(mvarOtherGroundDrivetrainKeychain)
End Function

Public Sub AddOtherGroundDrivetrainKey(OtherGroundDrivetrainkey As String)
    mvarOtherGroundDrivetrainKeychain = mAddKey(mvarOtherGroundDrivetrainKeychain, OtherGroundDrivetrainkey)
End Sub

Public Sub RemoveOtherGroundDrivetrainKey(OtherGroundDrivetrainkey As String)
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

Public Sub AddPowerSystemKey(PowerSystemKey As String)
    mvarPowerSystemKeyChain = mAddKey(mvarPowerSystemKeyChain, PowerSystemKey)
End Sub

Public Sub RemovePowerSystemKey(PowerSystemKey As String)
    mvarPowerSystemKeyChain = mRemoveKey(mvarPowerSystemKeyChain, PowerSystemKey)
End Sub


'//////////////////PowerConsumptionKeyChain management
Public Function GetCurrentPowerConsumptionKeys() As String()
    GetCurrentPowerConsumptionKeys = VariantArrayToStringArray(mvarPowerConsumptionKeyChain)
End Function

Public Sub AddPowerConsumptionKey(PowerConsumptionKey As String)
    mvarPowerConsumptionKeyChain = mAddKey(mvarPowerConsumptionKeyChain, PowerConsumptionKey)
End Sub

Public Sub RemovePowerConsumptionKey(PowerConsumptionKey As String)
    mvarPowerConsumptionKeyChain = mRemoveKey(mvarPowerConsumptionKeyChain, PowerConsumptionKey)
End Sub

'//////////////////FuelUsingSystemKeychain management
Public Function GetCurrentFuelUsingSystemKeys() As String()
    GetCurrentFuelUsingSystemKeys = VariantArrayToStringArray(mvarFuelUsingSystemKeyChain)
End Function

Public Sub AddFuelUsingSystemKey(FuelUsingSystemKey As String)
    mvarFuelUsingSystemKeyChain = mAddKey(mvarFuelUsingSystemKeyChain, FuelUsingSystemKey)
End Sub

Public Sub RemoveFuelUsingSystemKey(FuelUsingSystemKey As String)
    mvarFuelUsingSystemKeyChain = mRemoveKey(mvarFuelUsingSystemKeyChain, FuelUsingSystemKey)
End Sub


'//////////////////FuelStorageKeyChain management
Public Function GetCurrentFuelStorageKeys() As String()
    GetCurrentFuelStorageKeys = VariantArrayToStringArray(mvarFuelStorageKeyChain)
End Function

Public Sub AddFuelStorageKey(FuelStorageKey As String)
    mvarFuelStorageKeyChain = mAddKey(mvarFuelStorageKeyChain, FuelStorageKey)
End Sub

Public Sub RemoveFuelStorageKey(FuelStorageKey As String)
    mvarFuelStorageKeyChain = mRemoveKey(mvarFuelStorageKeyChain, FuelStorageKey)
End Sub


Public Property Let PowerSystemKeyChain(ByVal vdata As Variant)
    mvarPowerSystemKeyChain = vdata
End Property

Public Property Get PowerSystemKeyChain() As Variant
    PowerSystemKeyChain = mvarPowerSystemKeyChain
End Property

Public Property Let PowerConsumptionKeyChain(ByVal vdata As Variant)
    mvarPowerConsumptionKeyChain = vdata
End Property

Public Property Get PowerConsumptionKeyChain() As Variant
    PowerConsumptionKeyChain = mvarPowerConsumptionKeyChain
End Property


Public Property Let FuelUsingSystemKeyChain(ByVal vdata As Variant)
    mvarFuelUsingSystemKeyChain = vdata
End Property

Public Property Get FuelUsingSystemKeyChain() As Variant
    FuelUsingSystemKeyChain = mvarFuelUsingSystemKeyChain
End Property

Public Property Let FuelStorageKeyChain(ByVal vdata As Variant)
    mvarFuelStorageKeyChain = vdata
End Property

Public Property Get FuelStorageKeyChain() As Variant
    FuelStorageKeyChain = mvarFuelStorageKeyChain
End Property


Public Property Let SubAssembliesKeychain(ByVal vdata As Variant)
    mvarSubAssembliesKeychain = vdata
End Property

Public Property Get SubAssembliesKeychain() As Variant
    SubAssembliesKeychain = mvarSubAssembliesKeychain
End Property

Public Property Let LegsKeychain(ByVal vdata As Variant)
    mvarLegsKeychain = vdata
End Property

Public Property Get LegsKeychain() As Variant
    LegsKeychain = mvarLegsKeychain
End Property

Public Property Let LegDrivetrainKeychain(ByVal vdata As Variant)
    mvarLegDrivetrainKeychain = vdata
End Property

Public Property Get LegDrivetrainKeychain() As Variant
    LegDrivetrainKeychain = mvarLegDrivetrainKeychain
End Property

Public Property Let RotorDrivetrainKeychain(ByVal vdata As Variant)
    mvarRotorDrivetrainKeychain = vdata
End Property

Public Property Get RotorDrivetrainKeychain() As Variant
    RotorDrivetrainKeychain = mvarRotorDrivetrainKeychain
End Property

Public Property Let OrnithopterDrivetrainKeychain(ByVal vdata As Variant)
    mvarOrnithopterDrivetrainKeychain = vdata
End Property

Public Property Get OrnithopterDrivetrainKeychain() As Variant
    OrnithopterDrivetrainKeychain = mvarOrnithopterDrivetrainKeychain
End Property

Public Property Let OtherGroundDrivetrainKeychain(ByVal vdata As Variant)
    mvarOtherGroundDrivetrainKeychain = vdata
End Property

Public Property Get OtherGroundDrivetrainKeychain() As Variant
    OtherGroundDrivetrainKeychain = mvarOtherGroundDrivetrainKeychain
End Property

Public Property Let RotorsKeychain(ByVal vdata As Variant)
    mvarRotorsKeychain = vdata
End Property

Public Property Get RotorsKeychain() As Variant
    RotorsKeychain = mvarRotorsKeychain
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
