VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArmor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit





'local variable(s) to hold property value(s)
Private mvarMaterial As String
Private mvarQuality As String
Private mvarDR As Long
Private mvarRAP As Boolean
Private mvarElectrified As Boolean
Private mvarThermal As Boolean
Private mvarRadiation As Boolean
Private mvarCoating As String
Private mvarPD As Integer
Private mvarWeight As Double
Private mvarCost As Double
Private mvarTL As Integer
'Private mvarCost1 as single
'Private mvarCost2 as single
'Private mvarCost3 as single
'Private mvarCost4 as single
'Private mvarCost5 as single
'Private mvarCost6 as single
'Private mvarWeight1 as single
'Private mvarWeight2 as single
'Private mvarWeight3 as single
'Private mvarWeight4 as single
'Private mvarWeight5 as single
'Private mvarWeight6 as single
Private mvarMaterial1 As String
Private mvarMaterial2 As String
Private mvarMaterial3 As String
Private mvarMaterial4 As String
Private mvarMaterial5 As String
Private mvarMaterial6 As String
Private mvarQuality1 As String
Private mvarQuality2 As String
Private mvarQuality3 As String
Private mvarQuality4 As String
Private mvarQuality5 As String
Private mvarQuality6 As String
Private mvarEffectiveDR1 As Long
Private mvarEffectiveDR2 As Long
Private mvarEffectiveDR3 As Long
Private mvarEffectiveDR4 As Long
Private mvarEffectiveDR5 As Long
Private mvarEffectiveDR6 As Long
Private mvarDR1 As Long
Private mvarDR2 As Long
Private mvarDR3 As Long
Private mvarDR4 As Long
Private mvarDR5 As Long
Private mvarDR6 As Long
Private mvarPD1 As Integer
Private mvarPD2 As Integer
Private mvarPD3 As Integer
Private mvarPD4 As Integer
Private mvarPD5 As Integer
Private mvarPD6 As Integer
Private mvarParent As String
Private mvarLocation As String
Private mvarKey As String
Private mvarImage As Integer
Private mvarSelectedImage As Integer
Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarAverageDR As Single
Private mvarAveragePD As Single
Private mvarComment As String
Private mvarCName As String
Private mvarPrintOutput As String
Private mvarZZInit As Byte
Private mvarLogicalParent As String

Public Property Let LogicalParent(ByVal vdata As String)
    mvarLogicalParent = vdata
End Property

Public Property Get LogicalParent() As String
    LogicalParent = mvarLogicalParent
End Property


Public Property Let PrintOutput(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrintOutput = 5
    mvarPrintOutput = vdata
End Property


Public Property Get PrintOutput() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrintOutput
    PrintOutput = mvarPrintOutput
End Property



Public Property Let CName(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CName = 5
    mvarCName = vdata
End Property


Public Property Get CName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CName
    CName = mvarCName
End Property



Public Property Let Comment(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Comment = 5
    mvarComment = vdata
End Property


Public Property Get Comment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Comment
    Comment = mvarComment
End Property

Public Function LocationCheck() As Boolean
Dim TempCheck As Boolean
Dim InstallPoint As Integer
InstallPoint = Veh.Components(mvarParent).Datatype

'NOTE: Because Layering of armor is allowed,  user can have Overall armor along with location armor for instance on individual components!
'So there is no need to check for existance of other armor types
Select Case mvarDatatype

    Case ArmorComplexFacing, ArmorBasicFacing
        If (InstallPoint = Body) Or (InstallPoint = Superstructure) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Then
            TempCheck = True
        Else
            InfoPrint 1, "Armor by Facing can only be applied to a Body, Superstructure, Turret or Popturret."
            TempCheck = False
        End If
    Case ArmorOverall
        If InstallPoint = Body Then
            TempCheck = True
        Else
            InfoPrint 1, "Overall Armor can only be placed on the hull."
            TempCheck = False
        End If
    Case ArmorLocation
        If (InstallPoint = Body) Or (InstallPoint = Mast) Or (InstallPoint = Skid) Or (InstallPoint = Pod) Or (InstallPoint = Hovercraft) Or (InstallPoint = Hydrofoil) Or (InstallPoint = Wheel) Or (InstallPoint = Track) Or (InstallPoint = Leg) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = AutogyroRotor) Or (InstallPoint = CARotor) Or (InstallPoint = TTRotor) Or (InstallPoint = MMRotor) Or (InstallPoint = Gasbag) Or (InstallPoint = Superstructure) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Then
            TempCheck = True
        Else
            InfoPrint 1, "Location armor can only be applied to valid subassemblies."
            TempCheck = False
        End If
        
    Case ArmorOpenFrame
        Select Case InstallPoint
            Case Body, Mast, Skid, Pod, Hovercraft, Hydrofoil, Wheel, Track, _
                 Leg, Arm, Wing, AutogyroRotor, CARotor, TTRotor, MMRotor, Gasbag, _
                 Superstructure, Turret, Popturret
            
                TempCheck = True
        
            Case Else
                InfoPrint 1, "Location armor can only be applied to valid subassemblies."
                TempCheck = False
        End Select
        
    Case ArmorComponent
        Select Case InstallPoint
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
                PartialStabilizationGear, FullStabilizationGear, _
                UniversalMount, CasemateMount, DoorMount, Cyberslave, AntiBlastMagazine, _
                HardPoint, WeaponBay, Cargo, equipmentPod
        
                TempCheck = True
                
             Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, FlexibodyDrivetrain, _
                TrackedDrivetrain, LegDrivetrain, CARRotorDrivetrain, _
                MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, _
                AerialPropeller, DuctedFan, _
                PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, _
                Hydrojet, MHDTunnel, _
                MagLevLifter, _
                Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, _
                StandardThruster, SuperThruster, MegaThruster, _
                LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                OptimizedFusion, AntimatterThermal, AntimatterPion, _
                SolidRocketEngine, OrionEngine, _
                TeleportationDrive, Hyperdrive, JumpDrive, WarpDrive, _
                QuantumConveyor, SubQuantumConveyor, TwoQuantumConveyor, _
                ContraGravGenerator, _
                RadioDirectionFinder, RadioCommunicator, TightBeamRadio, VLFRadio, _
                CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, _
                LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator, _
                Headlight, Searchlight, InfraredSearchlight, _
                AstronomicalInstruments, Telescope, lightAmplification, LowlightTV, _
                ExtendableSensorPeriscope
            
                TempCheck = True
                
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
                ImprovedOpticalBombSight, AdvancedOpticalBombSight, OpticalBombSight, _
                FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, _
                LaserDesignator, LaserSpotTracker, RadarDetector, LaserSensor, _
                LaserRadarDetector, AreaRadarJammer, DeceptiveRadarJammer, _
                InfraredJammer, ChaffDecoyDischarger, SmokeDecoyDischarger, _
                FlareDecoyDischarger, SonarDecoyDischarger, HotSmokeDecoyDischarger, _
                PrismDecoyDischarger, BlackOutGasDecoyDischarger, RadarReflector, _
                BlipEnhancer, TEMPEST, MacroFrame, MainFrame, MicroFrame, MiniComputer, _
                SmallComputer, Terminal, _
                SurgicalInterface, InterfaceWeb, AutoInterfaceWeb, SocketInterface, _
                NeuralInductionField
                
                TempCheck = True
            
            Case ArmMotor, FireExtinguisherSystem, FullFireSuppressionSystem, _
                CompactFireSuppressionSystem, _
                BilgePump, CompleteWorkshop, MechanicWorkshop, EngineeringWorkshop, _
                ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, _
                ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, _
                MiniEngineeringWorkshop, MiniArmouryWorkshop, _
                ExtendableLadder, Crane, Winch, PowerShovel, WreckingCrane, ForkLift, _
                VehicularBridge, LaunchCatapult, SkyHook, Bore, SuperBore, _
                TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet, _
                OperatingRoom, StretcherPallet, EmergencySupportUnit, _
                EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable, _
                Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, _
                MovieScreenandProjectorSmall, HoloventureZone
                
                TempCheck = True
        
            Case CargoRamp, Airlock, MembraneAirlock, Forcelock, PassageTube, _
                ArmoredPassageTube, TeleportProjector, _
                BrigsandRestraints, BurglarAlarm, HighSecurityAlarm, _
                MutableLicensePlate, OilSprayer, PaintSprayer, _
                SmokeScreen, SpikeDropper, _
                VehicleBay, HangerBay, DryDock, SpaceDock, ExternalCradle, _
                ArrestorHook, VehicularParachute, RefuellingProbe, RefuellingDrogue, _
                FuelElectrolysisSystem, HydrogenFuelScoop, AtmosphereProcessor, _
                NuclearDamper, _
                SmallRealityStabilizer, MediumRealityStabilizer, HeavyRealityStabilizer, _
                ModularSocket, Module
            
                TempCheck = True
                
            Case PrimitiveManeuverControl, ElectronicDivingControl, ComputerizedDivingControl, _
                MechanicalManeuverControl, ElectronicManeuverControl, _
                ComputerizedManeuverControl, MechanicalDivingControl, _
                CrampedCrewStation, NormalCrewStation, RoomyCrewStation, _
                CycleCrewStation, HarnessCrewStation, CrampedSeat, NormalSeat, _
                RoomySeat, CrampedStandingRoom, NormalStandingRoom, RoomyStandingRoom, _
                CycleSeat, Hammock, Bunk, Cabin, LuxuryCabin, Suite, LuxurySuite, SmallGalley, _
                TotalLifeSystem, ArtificialGravityUnit, EnvironmentalControl, _
                NBCKit, LimitedLifeSystem, FullLifeSystem, _
                EjectionSeat, CrewEscapeCapsule, Airbag, CrashWeb, WombTank, _
                GravityWeb, GravCompensator
                
                
                TempCheck = True
               
            
            Case MuscleEngine, GasolineEngine, HPGasolineEngine, _
                TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, _
                SuperHPGasolineEngine, StandardDieselEngine, TurboStandardDieselEngine, _
                MarineDieselEngine, HPDieselEngine, TurboHPDieselEngine, CeramicEngine, _
                TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, _
                TurboHPCeramicEngine, SuperHPCeramicEngine, HydrogenCombustionEngine, _
                EarlySteamEngine, ForcedDraftSteamEngine, TripleExpansionSteamEngine, _
                SteamTurbine, StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, _
                StandardMHDTurbine, HPMHDTurbine, FuelCell, _
                FissionReactor, RTGReactor, NPU, FusionReactor, AntimatterReactor, _
                TotalConversionPowerPlant, CosmicPowerPlant, Soulburner, _
                ElementalFurnace, ManaEngine, Carnivore, Herbivore, Omnivore, Vampire, _
                ClockWork, LeadAcidBattery, AdvancedBattery, Flywheel, _
                RechargeablePowerCell, PowerCell, _
                AntiMatterBay, CoalBunker, WoodBunker, StandardTank, lightTank, _
                UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, _
                UltralightSelfSealingTank, Snorkel, ElectricContactPower, _
                LaserBeamedPowerReceiver, MaserBeamedPowerReceiver, NitrousOxideBooster
        
                TempCheck = True
            Case Else
                TempCheck = False
                InfoPrint 1, "Component armor cannot be applied to this vehicle component."

        End Select

    Case ArmorGunShield
        If InstallPoint = OpenMount Then
            TempCheck = True
        Else
            InfoPrint 1, "Gun Shields can only be applied to Open Mounts."
            TempCheck = False
        End If
        
    Case ArmorWheelGuard
        If (InstallPoint = Wheel) Or (InstallPoint = Track) Or (InstallPoint = Hovercraft) Then
            TempCheck = True
        Else
            InfoPrint 1, "Wheel Guards and Armored Skirts can only be applied to Wheels, Tracks, and Hovercraft subassemblies."
            TempCheck = False
        End If
        
End Select
    

If TempCheck Then SetLogicalParent
LocationCheck = TempCheck
End Function


Private Function GetLocation() As String
    On Error Resume Next
    If mvarLogicalParent = "" Then SetLogicalParent
    GetLocation = Veh.Components(mvarLogicalParent).Abbrev
    
End Function

Public Sub SetLogicalParent()
    ' in the case of Armor, its logical parent is always the same as its parent
    ' armor cannot be attached to GROUP components
    mvarLogicalParent = mvarParent
End Sub


Public Property Let Custom(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Custom = 5
    mvarCustom = vdata
End Property


Public Property Get Custom() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Custom
    Custom = mvarCustom
End Property

Public Property Let CustomDescription(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CustomDescription = 5
    mvarCustomDescription = vdata
End Property


Public Property Get CustomDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CustomDescription
    CustomDescription = mvarCustomDescription
End Property



Public Property Let Description(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vdata
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property


Public Property Let ParentDatatype(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ParentDatatype = 5
    mvarParentDatatype = vdata
End Property


Public Property Get ParentDatatype() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ParentDatatype
    ParentDatatype = mvarParentDatatype
End Property



Public Property Let Datatype(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Datatype = 5
    mvarDatatype = vdata
End Property


Public Property Get Datatype() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Datatype
    Datatype = mvarDatatype
End Property


Public Property Let Coating(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Coating = 5
    mvarCoating = vdata
End Property


Public Property Get Coating() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Coating
    Coating = mvarCoating
End Property



Public Property Let Radiation(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Radiation = 5
    mvarRadiation = vdata
End Property


Public Property Get Radiation() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Radiation
    Radiation = mvarRadiation
End Property



Public Property Let SelectedImage(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SelectedImage = 5
    mvarSelectedImage = vdata
End Property


Public Property Get SelectedImage() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedImage
    SelectedImage = mvarSelectedImage
End Property



Public Property Let Image(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Image = 5
    mvarImage = vdata
End Property


Public Property Get Image() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Image
    Image = mvarImage
End Property



Public Property Let Key(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.key = 5
    mvarKey = vdata
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.key
    Key = mvarKey
End Property



Public Property Let Location(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Location = 5
    mvarLocation = vdata
End Property


Public Property Get Location() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Location
    Location = mvarLocation
End Property



Public Property Let Parent(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Parent = 5
    mvarParent = vdata
End Property


Public Property Get Parent() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parent
    Parent = mvarParent
End Property


Public Property Let TL(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TL = 5

    mvarTL = vdata
End Property


Public Property Get TL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    TL = mvarTL
End Property



Public Property Let Cost(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cost = 5
    mvarCost = vdata
End Property


Public Property Get Cost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cost
    Cost = mvarCost
End Property



Public Property Let Weight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Weight = 5
    mvarWeight = vdata
End Property


Public Property Get Weight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Weight
    Weight = mvarWeight
End Property


Public Property Let PD(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PD = 5
    mvarPD = vdata
End Property


Public Property Get PD() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PD
    PD = mvarPD
End Property

Public Property Let PD1(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PD1 = 5
    mvarPD1 = vdata
End Property


Public Property Get PD1() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PD1
    PD1 = mvarPD1
End Property
Public Property Let PD2(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PD2 = 5
    mvarPD2 = vdata
End Property


Public Property Get PD2() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PD2
    PD2 = mvarPD2
End Property
Public Property Let PD3(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PD3 = 5
    mvarPD3 = vdata
End Property


Public Property Get PD3() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PD3
    PD3 = mvarPD3
End Property
Public Property Let PD4(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PD4 = 5
    mvarPD4 = vdata
End Property


Public Property Get PD4() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PD4
    PD4 = mvarPD4
End Property
Public Property Let PD5(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PD5 = 5
    mvarPD5 = vdata
End Property


Public Property Get PD5() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PD5
    PD5 = mvarPD5
End Property
Public Property Let PD6(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PD6 = 5
    mvarPD6 = vdata
End Property


Public Property Get PD6() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PD6
    PD6 = mvarPD6
End Property

Public Property Let AveragePD(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AveragePD = 5
    mvarAveragePD = vdata
End Property


Public Property Get AveragePD() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AveragePD
    AveragePD = mvarAveragePD
End Property



Public Property Let Thermal(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Thermal = 5
    mvarThermal = vdata
End Property


Public Property Get Thermal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Thermal
    Thermal = mvarThermal
End Property



Public Property Let Electrified(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Electrified = 5
    mvarElectrified = vdata
End Property


Public Property Get Electrified() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Electrified
    Electrified = mvarElectrified
End Property



Public Property Let RAP(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RAP = 5
    mvarRAP = vdata
End Property


Public Property Get RAP() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RAP
    RAP = mvarRAP
End Property



Public Property Let DR(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR = 5
    mvarDR = vdata
End Property


Public Property Get DR() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DR
    DR = mvarDR
End Property

Public Property Let DR1(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR1 = 5
    mvarDR1 = vdata
End Property


Public Property Get DR1() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DR1
    DR1 = mvarDR1
End Property

Public Property Let DR2(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR2 = 5
    mvarDR2 = vdata
End Property


Public Property Get DR2() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DR2
    DR2 = mvarDR2
End Property

Public Property Let DR3(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR3= 5
    mvarDR3 = vdata
End Property


Public Property Get DR3() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DR3
    DR3 = mvarDR3
End Property

Public Property Let DR4(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR4 = 5
    mvarDR4 = vdata
End Property


Public Property Get DR4() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DR4
    DR4 = mvarDR4
End Property

Public Property Let DR5(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR5 = 5
    mvarDR5 = vdata
End Property


Public Property Get DR5() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DR5
    DR5 = mvarDR5
End Property

Public Property Let DR6(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR6 = 5
    mvarDR6 = vdata
End Property


Public Property Get DR6() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DR6
    DR6 = mvarDR6
End Property

Public Property Let AverageDR(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AverageDR = 5
    mvarAverageDR = vdata
End Property


Public Property Get AverageDR() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AverageDR
    AverageDR = mvarAverageDR
End Property

Public Property Let EffectiveDR1(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EffectiveDR1 = 5
    mvarEffectiveDR1 = vdata
End Property


Public Property Get EffectiveDR1() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EffectiveDR1
    EffectiveDR1 = mvarEffectiveDR1
End Property

Public Property Let EffectiveDR2(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EffectiveDR2 = 5
    mvarEffectiveDR2 = vdata
End Property


Public Property Get EffectiveDR2() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EffectiveDR2
    EffectiveDR2 = mvarEffectiveDR2
End Property

Public Property Let EffectiveDR3(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EffectiveDR3= 5
    mvarEffectiveDR3 = vdata
End Property


Public Property Get EffectiveDR3() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EffectiveDR3
    EffectiveDR3 = mvarEffectiveDR3
End Property

Public Property Let EffectiveDR4(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EffectiveDR4 = 5
    mvarEffectiveDR4 = vdata
End Property


Public Property Get EffectiveDR4() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EffectiveDR4
    EffectiveDR4 = mvarEffectiveDR4
End Property

Public Property Let EffectiveDR5(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EffectiveDR5 = 5
    mvarEffectiveDR5 = vdata
End Property


Public Property Get EffectiveDR5() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EffectiveDR5
    EffectiveDR5 = mvarEffectiveDR5
End Property

Public Property Let EffectiveDR6(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EffectiveDR6 = 5
    mvarEffectiveDR6 = vdata
End Property


Public Property Get EffectiveDR6() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EffectiveDR6
    EffectiveDR6 = mvarEffectiveDR6
End Property

Public Property Let Quality(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quality = 5
    mvarQuality = vdata
End Property


Public Property Get Quality() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quality
    Quality = mvarQuality
End Property

Public Property Let Quality1(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quality1 = 5
    mvarQuality1 = vdata
End Property


Public Property Get Quality1() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quality1
    Quality1 = mvarQuality1
End Property
Public Property Let Quality2(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quality2 = 5
    mvarQuality2 = vdata
End Property


Public Property Get Quality2() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quality2
    Quality2 = mvarQuality2
End Property
Public Property Let Quality3(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quality3 = 5
    mvarQuality3 = vdata
End Property


Public Property Get Quality3() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quality3
    Quality3 = mvarQuality3
End Property
Public Property Let Quality4(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quality4 = 5
    mvarQuality4 = vdata
End Property


Public Property Get Quality4() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quality4
    Quality4 = mvarQuality4
End Property
Public Property Let Quality5(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quality5 = 5
    mvarQuality5 = vdata
End Property


Public Property Get Quality5() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quality5
    Quality5 = mvarQuality5
End Property
Public Property Let Quality6(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quality6 = 5
    mvarQuality6 = vdata
End Property


Public Property Get Quality6() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quality6
    Quality6 = mvarQuality6
End Property

Public Property Let Material(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Material = 5
    mvarMaterial = vdata
End Property


Public Property Get Material() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Material
    Material = mvarMaterial
End Property


Public Property Let Material1(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Material1 = 5
    mvarMaterial1 = vdata
End Property


Public Property Get Material1() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Material1
    Material1 = mvarMaterial1
End Property

Public Property Let Material2(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Material2 = 5
    mvarMaterial2 = vdata
End Property


Public Property Get Material2() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Material2
    Material2 = mvarMaterial2
End Property

Public Property Let Material3(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Material3 = 5
    mvarMaterial3 = vdata
End Property


Public Property Get Material3() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Material3
    Material3 = mvarMaterial3
End Property

Public Property Let Material4(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Material4 = 5
    mvarMaterial4 = vdata
End Property


Public Property Get Material4() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Material4
    Material4 = mvarMaterial4
End Property

Public Property Let Material5(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Material5 = 5
    mvarMaterial5 = vdata
End Property


Public Property Get Material5() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Material5
    Material5 = mvarMaterial5
End Property

Public Property Let Material6(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Material = 5
    mvarMaterial6 = vdata
End Property


Public Property Get Material6() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Material6
    Material6 = mvarMaterial6
End Property

Public Sub Init()


Select Case mvarDatatype
    
      Case ArmorComplexFacing
        Material = "wood"
        Quality = "standard"
        DR = 1
        
        Material1 = "wood"
        Material2 = "wood"
        Material3 = "wood"
        Material4 = "wood"
        Material5 = "wood"
        Material6 = "wood"
        Quality1 = "standard"
        Quality2 = "standard"
        Quality3 = "standard"
        Quality4 = "standard"
        Quality5 = "standard"
        Quality6 = "standard"
        DR1 = 1
        DR2 = 1
        DR3 = 1
        DR4 = 1
        DR5 = 1
        If Veh.Components(mvarParent).Datatype = Body Then
            DR6 = 1
        Else
            DR6 = 0
        End If
    Case ArmorBasicFacing
       
        Material = "wood"
        Quality = "standard"
        DR1 = 1
        DR2 = 1
        DR3 = 1
        DR4 = 1
        DR5 = 1
        If Veh.Components(mvarParent).Datatype = Body Then
            DR6 = 1
        Else
            DR6 = 0
        End If
    Case ArmorOpenFrame
       
        Material = "wood"
        Quality = "standard"
        DR = 1

    Case ArmorGunShield
       
        Material = "wood"
        Quality = "standard"
        DR = 1

    Case ArmorLocation
       
        Material = "wood"
        Quality = "standard"
        DR = 1

    Case ArmorComponent
        
        Material = "wood"
        Quality = "standard"
        DR = 1
        
    Case ArmorOverall
        
        Material = "wood"
        Quality = "standard"
        DR = 1

    Case ArmorWheelGuard
        Material = "wood"
        Quality = "standard"
        DR = 1
        
End Select

mvarTL = Veh.Components(mvarParent).TL

End Sub

Private Sub Class_Initialize()

mvarCoating = "none"
mvarRadiation = False
mvarThermal = False
mvarRAP = False
mvarElectrified = False

End Sub

Public Sub StatsUpdate()
Dim component As Integer
Dim SlopeR As String
Dim SlopeL As String
Dim SlopeF As String
Dim SlopeB As String
Dim sCompare(1 To 6) As String
Dim sSides() As String
Dim i As Long
Dim j As Long
Dim count As Long

mvarZZInit = 1
mvarPrintOutput = "" ' reinit this var

mvarLocation = GetLocation

component = Veh.Components(mvarParent).Datatype

If (mvarDatatype = ArmorBasicFacing) Or (mvarDatatype = ArmorComplexFacing) Then
    SlopeR = Veh.Components(mvarParent).SlopeR
    SlopeL = Veh.Components(mvarParent).SlopeL
    SlopeF = Veh.Components(mvarParent).SlopeF
    SlopeB = Veh.Components(mvarParent).SlopeB
    
    
    CalcByFacingArmorWeightCost
    'get the PD
    ' Get the PD for each side (used for both Complex and Basic)
    mvarPD1 = CalcPD(mvarDR1, SlopeR, mvarMaterial1)
    mvarPD2 = CalcPD(mvarDR2, SlopeL, mvarMaterial2)
    mvarPD3 = CalcPD(mvarDR3, SlopeF, mvarMaterial3)
    mvarPD4 = CalcPD(mvarDR4, SlopeB, mvarMaterial4)
    mvarPD5 = CalcPD(mvarDR5, "none", mvarMaterial5) 'the top and underside dont have slope
    mvarPD6 = CalcPD(mvarDR6, "none", mvarMaterial6)
        
    'get the effective DR
    mvarEffectiveDR1 = CalcEffectiveDR(mvarDR1, SlopeR)
    mvarEffectiveDR2 = CalcEffectiveDR(mvarDR2, SlopeL)
    mvarEffectiveDR3 = CalcEffectiveDR(mvarDR3, SlopeF)
    mvarEffectiveDR4 = CalcEffectiveDR(mvarDR4, SlopeB)
    mvarEffectiveDR5 = mvarDR5
    mvarEffectiveDR6 = mvarDR6
    
        
ElseIf mvarDatatype = ArmorLocation Then
    
    CalcArmorWeightCost
    If component = Body Or component = Superstructure Or component = Turret Or component = Popturret Then
        
        SlopeR = Veh.Components(mvarParent).SlopeR
        SlopeL = Veh.Components(mvarParent).SlopeL
        SlopeF = Veh.Components(mvarParent).SlopeF
        SlopeB = Veh.Components(mvarParent).SlopeB
        ' Get the PD for each side (used for both Complex and Basic)
        mvarPD1 = CalcPD(mvarDR, SlopeR, mvarMaterial)
        mvarPD2 = CalcPD(mvarDR, SlopeL, mvarMaterial)
        mvarPD3 = CalcPD(mvarDR, SlopeF, mvarMaterial)
        mvarPD4 = CalcPD(mvarDR, SlopeB, mvarMaterial)
        mvarPD5 = CalcPD(mvarDR, "none", mvarMaterial)
        mvarPD6 = CalcPD(mvarDR, "none", mvarMaterial)
    Else
        mvarPD = CalcPD(mvarDR, "none", mvarMaterial)
    End If
Else
    CalcArmorWeightCost
    mvarPD = CalcPD(mvarDR, "none", mvarMaterial)
End If

'print output
If mvarRAP Then
End If
If mvarElectrified Then
End If
If mvarThermal Then
End If
If mvarRadiation Then
End If
If mvarCoating <> "none" Then
End If
If mvarPD Then
End If

Select Case mvarDatatype
    Case ArmorComplexFacing
        'can have different everything
        'get each side
        sCompare(1) = " PD " + Format(mvarPD1) + ", DR " + Format(mvarDR1) + " " + mvarQuality1 + " " + mvarMaterial1 + ". "
        sCompare(2) = " PD " + Format(mvarPD2) + ", DR " + Format(mvarDR2) + " " + mvarQuality2 + " " + mvarMaterial2 + ". "
        sCompare(3) = " PD " + Format(mvarPD3) + ", DR " + Format(mvarDR3) + " " + mvarQuality3 + " " + mvarMaterial3 + ". "
        sCompare(4) = " PD " + Format(mvarPD4) + ", DR " + Format(mvarDR4) + " " + mvarQuality4 + " " + mvarMaterial4 + ". "
        sCompare(5) = " PD " + Format(mvarPD5) + ", DR " + Format(mvarDR5) + " " + mvarQuality5 + " " + mvarMaterial5 + ". "
        sCompare(6) = " PD " + Format(mvarPD6) + ", DR " + Format(mvarDR6) + " " + mvarQuality6 + " " + mvarMaterial6 + ". "
        'move the first side (the Right) into the sSides array
        ReDim sSides(2, 1)
        sSides(1, 1) = sCompare(1)
        sSides(2, 1) = "R"
        count = 1
        'compare each side to see if it can be grouped with another
        For j = 2 To 6
            For i = 1 To count
                If sCompare(j) = sSides(1, i) Then
                    sSides(1, i) = sCompare(j)
                    sSides(2, i) = sSides(2, i) + "," + GetSideLetterFromNumber(j)
                ElseIf i = count Then
                    ReDim Preserve sSides(2, i + 1)
                    count = count + 1
                    sSides(1, i + 1) = sCompare(j)
                    sSides(2, i + 1) = GetSideLetterFromNumber(j)
                End If
            Next
        Next
        'get final string and include the surface options to the armor
        For i = 1 To count
            mvarPrintOutput = mvarPrintOutput + " " + sSides(2, i) + ": " + sSides(1, i)
        Next
        mvarPrintOutput = mvarPrintOutput + " (" + Format(mvarWeight, p_sFormat) + " lbs., $" + Format(mvarCost, p_sFormat) + ")."
        
    Case ArmorBasicFacing
        'same material and quality but different DR's and PD's
        sCompare(1) = " PD " + Format(mvarPD1) + ", DR " + Format(mvarDR1) + " "
        sCompare(2) = " PD " + Format(mvarPD2) + ", DR " + Format(mvarDR2) + " "
        sCompare(3) = " PD " + Format(mvarPD3) + ", DR " + Format(mvarDR3) + " "
        sCompare(4) = " PD " + Format(mvarPD4) + ", DR " + Format(mvarDR4) + " "
        sCompare(5) = " PD " + Format(mvarPD5) + ", DR " + Format(mvarDR5) + " "
        sCompare(6) = " PD " + Format(mvarPD6) + ", DR " + Format(mvarDR6) + " "
        'move the first side (the Right) into the sSides array
        ReDim sSides(2, 1)
        sSides(1, 1) = sCompare(1)
        sSides(2, 1) = "R"
        count = 1
        'compare each side to see if it can be grouped with another
        For j = 2 To 6
            For i = 1 To count
                If sCompare(j) = sSides(1, i) Then
                    sSides(1, i) = sCompare(j)
                    sSides(2, i) = sSides(2, i) + "," + GetSideLetterFromNumber(j)
                ElseIf i = count Then
                    ReDim Preserve sSides(2, i + 1)
                    count = count + 1
                    sSides(1, i + 1) = sCompare(j)
                    sSides(2, i + 1) = GetSideLetterFromNumber(j)
                End If
            Next
        Next
        'get final string and include the surface options to the armor
        For i = 1 To count
            mvarPrintOutput = mvarPrintOutput + " " + sSides(2, i) + ": " + sSides(1, i)
        Next
        mvarPrintOutput = mvarQuality + " " + mvarMaterial + mvarPrintOutput + " (" + Format(mvarWeight, p_sFormat) + " lbs., $" + Format(mvarCost, p_sFormat) + ")."
 
    
    Case ArmorLocation 'this can still have different PD's on Body and Turrets do to slope differences
        sCompare(1) = " PD " + Format(mvarPD1) + ", DR " + Format(mvarDR) + " "
        sCompare(2) = " PD " + Format(mvarPD2) + ", DR " + Format(mvarDR) + " "
        sCompare(3) = " PD " + Format(mvarPD3) + ", DR " + Format(mvarDR) + " "
        sCompare(4) = " PD " + Format(mvarPD4) + ", DR " + Format(mvarDR) + " "
        sCompare(5) = " PD " + Format(mvarPD5) + ", DR " + Format(mvarDR) + " "
        sCompare(6) = " PD " + Format(mvarPD6) + ", DR " + Format(mvarDR) + " "
        'move the first side (the Right) into the sSides array
        ReDim sSides(2, 1)
        sSides(1, 1) = sCompare(1)
        sSides(2, 1) = "R"
        count = 1
        'compare each side to see if it can be grouped with another
        For j = 2 To 6
            For i = 1 To count
                If sCompare(j) = sSides(1, i) Then
                    sSides(1, i) = sCompare(j)
                    sSides(2, i) = sSides(2, i) + "," + GetSideLetterFromNumber(j)
                ElseIf i = count Then
                    ReDim Preserve sSides(2, i + 1)
                    count = count + 1
                    sSides(1, i + 1) = sCompare(j)
                    sSides(2, i + 1) = GetSideLetterFromNumber(j)
                End If
            Next
        Next
        'get final string and include the surface options to the armor
        mvarPrintOutput = "DR " + Format(mvarDR) + " "
        For i = 1 To count
            mvarPrintOutput = mvarPrintOutput + " " + sSides(2, i) + ": " + sSides(1, i)
        Next
        mvarPrintOutput = mvarPrintOutput + mvarQuality + " " + mvarMaterial + " (" + Format(mvarWeight, p_sFormat) + " lbs., $" + Format(mvarCost, p_sFormat) + ")."
 
 
    Case ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
        mvarPrintOutput = "PD " + Format(mvarPD) + ", DR " + Format(mvarDR) + " " + mvarQuality + " " + mvarMaterial + " (" + Format(mvarWeight, p_sFormat) + " lbs., $" + Format(mvarCost, p_sFormat) + ")."
 
End Select

End Sub

Private Function GetSideLetterFromNumber(i As Long) As String
    Select Case i
        Case 1
            GetSideLetterFromNumber = "R"
        Case 2
            GetSideLetterFromNumber = "L"
        Case 3
            GetSideLetterFromNumber = "F"
        Case 4
            GetSideLetterFromNumber = "B"
        Case 5
            GetSideLetterFromNumber = "T"
        Case 6
            GetSideLetterFromNumber = "U"
    End Select
End Function

Public Function FillMaterial() As String()
' populate the material combo
Dim materialarray() As String
ReDim materialarray(1)

    If mvarTL <= 6 Then
            materialarray = mAddKey(materialarray, "wood")
            materialarray = mAddKey(materialarray, "metal")
            materialarray = mAddKey(materialarray, "nonrigid")
    Else ' if its greater than or equal to 7
            materialarray = mAddKey(materialarray, "wood")
            materialarray = mAddKey(materialarray, "metal")
            materialarray = mAddKey(materialarray, "ablative")
            materialarray = mAddKey(materialarray, "fireproof ablative")
            materialarray = mAddKey(materialarray, "nonrigid")
            materialarray = mAddKey(materialarray, "composite")
            materialarray = mAddKey(materialarray, "laminate")
    End If

FillMaterial = materialarray
End Function

Public Function FillQuality(sMaterial As String) As String()
Dim MaterialCombo As ComboBox
Dim Selected As String ' holds the users selected Armor material
Dim arrQuality() As Integer 'holds the list of suitable quality
Dim iSelected As Integer 'holds converted Selected string
Dim element As Variant 'one element of the arrQuality array
Dim i As Integer 'counter
Dim count As Integer 'another counter
Dim qualityarray() As String
Dim TempTL As Integer

ReDim qualityarray(1)

Const Cheap = 1
Const Standard = 2
Const Expensive = 3
Const Advanced = 4

'get the type of armor that the user selected
Selected = sMaterial
'convert the Selected into an integer
Select Case Selected
    Case "wood"
        iSelected = 1
    Case "metal"
        iSelected = 2
    Case "ablative"
        iSelected = 3
    Case "fireproof ablative"
        iSelected = 4
    Case "nonrigid"
        iSelected = 5
    Case "composite"
        iSelected = 6
    Case "laminate"
        iSelected = 7
End Select

count = 1 ' init the counter
'given the tech level, produce list of quality types
TempTL = Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
TempTL = Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
For i = 1 To UBound(ArmorMatrix)
    If ArmorMatrix(i).TL = TempTL Then
        If ArmorMatrix(i).MaterialType = iSelected Then
            If ArmorMatrix(i).WeightMod <> 0 Then
                ReDim Preserve arrQuality(count)
                arrQuality(count) = ArmorMatrix(i).Quality
                count = count + 1
                If count > 5 Then Exit For Else
            End If
        End If
    End If
Next

'fill the Quality combo with the list of available items
For Each element In arrQuality
    Select Case element
        Case Cheap
            qualityarray = mAddKey(qualityarray, "cheap")
            'If TempText = "cheap" Then TextFlag = True
        Case Standard
            qualityarray = mAddKey(qualityarray, "standard")
            'If TempText = "standard" Then TextFlag = True
        Case Expensive
            qualityarray = mAddKey(qualityarray, "expensive")
            'If TempText = "expensive" Then TextFlag = True
        Case Advanced
            qualityarray = mAddKey(qualityarray, "advanced")
            'If TempText = "advanced" Then TextFlag = True
    End Select
Next

FillQuality = qualityarray

End Function

Sub CalcArmorWeightCost()
Dim TempTL As Integer
' This routine calculates the Cost and Weight of the armor.
    'These contstant values must match those in the module "modArmor" since
    'the armormatrix uses integers and not string names for the material and quality
    Const Cheap = 1
    Const Standard = 2
    Const Expensive = 3
    Const Advanced = 4
    
    Const Wood = 1
    Const Metal = 2
    Const Ablative = 3
    Const FireproofAblative = 4
    Const NonRigid = 5
    Const Composite = 6
    Const Laminate = 7

    Dim Area As Single 'holds the surface area
    Dim CostModifier As Single
    Dim WeightModifier As Single
    Dim SelectedMaterial As Integer
    Dim SelectedQuality As Integer
    Dim i As Integer ' counter
    
    ' Get the surface area based on the armor datatype being used
    If (mvarDatatype = ArmorWheelGuard) Then
        Area = Veh.Components(mvarParent).SurfaceArea / 2
    ElseIf (mvarDatatype = ArmorGunShield) Or (mvarDatatype = ArmorOpenFrame) Then
        Area = Veh.Components(mvarParent).SurfaceArea / 5
    ElseIf (mvarDatatype = ArmorLocation) Or (mvarDatatype = ArmorComponent) Then
        Area = Veh.Components(mvarParent).SurfaceArea
    ElseIf mvarDatatype = ArmorOverall Then
        Area = Veh.Stats.StructuralSurfaceArea
    End If
    
    'Determine Material
    If mvarMaterial = "wood" Then
        SelectedMaterial = Wood
    ElseIf mvarMaterial = "metal" Then SelectedMaterial = Metal
    ElseIf mvarMaterial = "ablative" Then SelectedMaterial = Ablative
    ElseIf mvarMaterial = "fireproof ablative" Then SelectedMaterial = FireproofAblative
    ElseIf mvarMaterial = "nonrigid" Then SelectedMaterial = NonRigid
    ElseIf mvarMaterial = "composite" Then SelectedMaterial = Composite
    ElseIf mvarMaterial = "laminate" Then SelectedMaterial = Laminate
    End If
        
    'Determine Quality
    If mvarQuality = "cheap" Then
        SelectedQuality = Cheap
    ElseIf mvarQuality = "standard" Then SelectedQuality = Standard
    ElseIf mvarQuality = "expensive" Then SelectedQuality = Expensive
    ElseIf mvarQuality = "advanced" Then SelectedQuality = Advanced
    End If

   ' Get the Cost and Weight Modifiers
       TempTL = Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
       TempTL = Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
       For i = 1 To UBound(ArmorMatrix)
           If ArmorMatrix(i).TL = TempTL Then
               If ArmorMatrix(i).MaterialType = SelectedMaterial Then
                   If ArmorMatrix(i).Quality = SelectedQuality Then
                       CostModifier = ArmorMatrix(i).Cost
                       WeightModifier = ArmorMatrix(i).WeightMod
                       Exit For
                    End If
                End If
           End If
       Next
          
       
    mvarWeight = mvarDR * Area * WeightModifier
    mvarCost = mvarWeight * CostModifier
    
    'get the final weight and cost by adding the cost/weight of the surface features
    CalcSurfaceFeaturesCostandWeight Area
End Sub

Sub CalcByFacingArmorWeightCost()
' This routine calculates the Cost and Weight of the armor
Dim Area As Single 'holds the surface area
Dim CostModifier(0 To 6) As Single
Dim WeightModifier(0 To 6) As Single
Dim iWeight(0 To 6) As Single
Dim iCost(0 To 6) As Single
Dim SelectedMaterial(0 To 6) As String
Dim SelectedQuality(0 To 6) As String
Dim iSelectedQuality(0 To 6) As Integer
Dim iSelectedMaterial(0 To 6) As Integer
Dim count As Integer
Dim i As Integer
Dim arrMaterial(0 To 6) As String
Dim arrQuality(0 To 6) As String
Dim arrDR(0 To 6) As Long
Dim TempCost As Single
Dim TempWeight As Single
Dim TempTL As Integer
'fill the arrMaterial array and arrQuality
arrMaterial(0) = mvarMaterial1
arrMaterial(1) = mvarMaterial2
arrMaterial(2) = mvarMaterial3
arrMaterial(3) = mvarMaterial4
arrMaterial(4) = mvarMaterial5
arrMaterial(5) = mvarMaterial6

arrQuality(0) = mvarQuality1
arrQuality(1) = mvarQuality2
arrQuality(2) = mvarQuality3
arrQuality(3) = mvarQuality4
arrQuality(4) = mvarQuality5
arrQuality(5) = mvarQuality6

'fill the arrDR aray
arrDR(0) = mvarDR1
arrDR(1) = mvarDR2
arrDR(2) = mvarDR3
arrDR(3) = mvarDR4
arrDR(4) = mvarDR5
arrDR(5) = mvarDR6

' re-init variables
TempCost = 0
TempWeight = 0

Area = Veh.Components(mvarParent).SurfaceArea
' There are just two paths, one for Complex and one for Basic
Select Case mvarDatatype
    Case ArmorComplexFacing
        For count = 0 To 5
            ' Get the quality and material of the armor
            SelectedMaterial(count) = arrMaterial(count)
            SelectedQuality(count) = arrQuality(count)
            'convert the Selected into an integer
            Select Case SelectedMaterial(count)
                Case "wood"
                    iSelectedMaterial(count) = 1
                Case "metal"
                    iSelectedMaterial(count) = 2
                Case "ablative"
                    iSelectedMaterial(count) = 3
                Case "fireproof ablative"
                    iSelectedMaterial(count) = 4
                Case "nonrigid"
                    iSelectedMaterial(count) = 5
                Case "composite"
                    iSelectedMaterial(count) = 6
                Case "laminate"
                    iSelectedMaterial(count) = 7
            End Select
            Select Case SelectedQuality(count)
                Case "cheap"
                    iSelectedQuality(count) = 1
                Case "standard"
                    iSelectedQuality(count) = 2
                Case "expensive"
                    iSelectedQuality(count) = 3
                Case "advanced"
                    iSelectedQuality(count) = 4
            End Select
            ' Get the Cost and Weight Modifiers
                TempTL = Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
                TempTL = Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
                For i = 1 To UBound(ArmorMatrix)
                    If ArmorMatrix(i).TL = TempTL Then
                        If ArmorMatrix(i).MaterialType = iSelectedMaterial(count) Then
                            If ArmorMatrix(i).Quality = iSelectedQuality(count) Then
                                CostModifier(count) = ArmorMatrix(i).Cost
                                WeightModifier(count) = ArmorMatrix(i).WeightMod
                                Exit For
                             End If
                         End If
                    End If
                Next
                'get the average DR
                CalcAverageDr
                ' Get the Cost and weight of each face
                If mvarParent = "1_" Then
                    iWeight(count) = Val(arrDR(count)) * (Area / 6) * WeightModifier(count)
                Else
                    iWeight(count) = Val(arrDR(count)) * (Area / 5) * WeightModifier(count)
                End If
                iCost(count) = iWeight(count) * CostModifier(count)
                TempWeight = TempWeight + iWeight(count)
                TempCost = TempCost + iCost(count)
            Next
            
    Case ArmorBasicFacing

        'convert the Selected into an integer
        Select Case mvarMaterial
            Case "wood"
                iSelectedMaterial(0) = 1
            Case "metal"
                iSelectedMaterial(0) = 2
            Case "ablative"
                iSelectedMaterial(0) = 3
            Case "fireproof ablative"
                iSelectedMaterial(0) = 4
            Case "nonrigid"
                iSelectedMaterial(0) = 5
            Case "composite"
                iSelectedMaterial(0) = 6
            Case "laminate"
                iSelectedMaterial(0) = 7
        End Select
        Select Case mvarQuality
            Case "cheap"
                iSelectedQuality(0) = 1
            Case "standard"
                iSelectedQuality(0) = 2
            Case "expensive"
                iSelectedQuality(0) = 3
            Case "advanced"
                iSelectedQuality(0) = 4
       End Select
       ' Get the Cost and Weight Modifiers
       TempTL = Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
       TempTL = Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
       For i = 1 To UBound(ArmorMatrix)
           If ArmorMatrix(i).TL = TempTL Then
               If ArmorMatrix(i).MaterialType = iSelectedMaterial(0) Then
                   If ArmorMatrix(i).Quality = iSelectedQuality(0) Then
                       CostModifier(0) = ArmorMatrix(i).Cost
                       WeightModifier(0) = ArmorMatrix(i).WeightMod
                    End If
                End If
           End If
       Next
       'call routine to calc averagedr
       CalcAverageDr
       ' Get the Final Cost and Final Weight
       TempWeight = AverageDR * Area * WeightModifier(0)
       TempCost = TempWeight * CostModifier(0)
End Select


'save these cost and weight results to the armor class
mvarCost = TempCost
mvarWeight = TempWeight

'get the final cost which includes the cost of the surface features
CalcSurfaceFeaturesCostandWeight Area
End Sub


Function CalcEffectiveDR(DR As Long, Slope As String)
Dim Modifier As Single

If Slope = "none" Then
    Modifier = 1
ElseIf Slope = "30 degrees" Then
    Modifier = 1.5
ElseIf Slope = "60 degrees" Then
    Modifier = 2
End If

CalcEffectiveDR = DR * Modifier
End Function
Function CalcPD(DR As Long, Slope As String, Material As String) As Long
Dim PD As Integer

If DR = 0 Then
    PD = 0
ElseIf DR = 1 Then
    PD = 1
ElseIf DR <= 4 Then
    PD = 2
ElseIf DR <= 15 Then
    PD = 3
ElseIf DR >= 16 Then
    PD = 4
End If
    
'check for max values for Wood and Nonrigid armor
If Material = "nonrigid" Then
    If PD > 2 Then PD = 2
ElseIf Material = "wood" Then
    If PD > 3 Then PD = 3
End If

'add bonus for slope
'Note: the bonus's are placed below the checks for  max values for wood and nonrigid.
'if users request, it can be moved above it
If Slope = "none" Then
ElseIf Slope = "30 degrees" Then
    PD = PD + 1
ElseIf Slope = "60 degrees" Then
    PD = PD + 2
End If
    
CalcPD = PD
End Function


Function GetLowestDR() As Long
    '//this function only gets called during Aerial performance calculations.
    '//its job is to return the DR of the armor. In the case of seperate DR's
    '//for each face, then  it will return the lowest one.
    '//only DR from metal, composite or laminate armor counts.. all other types
    '//return 0.
    On Error Resume Next
    Dim lngRetval As Long
    
    Select Case mvarDatatype
       Case ArmorComplexFacing
            Select Case mvarMaterial1
                Case "metal", "composite", "laminate"
                    lngRetval = mvarDR1
                    Select Case mvarMaterial2
                        Case "metal", "composite", "laminate"
                            lngRetval = Minimum(mvarDR1, mvarDR2)
                            Select Case mvarMaterial3
                                Case "metal", "composite", "laminate"
                                    lngRetval = Minimum(lngRetval, mvarDR3)
                                    Select Case mvarMaterial4
                                        Case "metal", "composite", "laminate"
                                            lngRetval = Minimum(lngRetval, mvarDR4)
                                            Select Case mvarMaterial5
                                                Case "metal", "composite", "laminate"
                                                    lngRetval = Minimum(lngRetval, mvarDR5)
                                                    Select Case mvarMaterial6
                                                        Case "metal", "composite", "laminate"
                                                            lngRetval = Minimum(lngRetval, mvarDR6)
                                                        Case Else
                                                            lngRetval = 0
                                                    End Select
                                                Case Else
                                                    lngRetval = 0
                                            End Select
                                        Case Else
                                            lngRetval = 0
                                    End Select
                                Case Else
                                    lngRetval = 0
                            End Select
                        Case Else
                            lngRetval = 0
                    End Select
                Case Else
                    lngRetval = 0
            End Select
           
       Case ArmorBasicFacing
           Select Case mvarMaterial
                Case "metal", "composite", "laminate"
                    lngRetval = Minimum(mvarDR1, mvarDR2)
                    lngRetval = Minimum(lngRetval, mvarDR3)
                    lngRetval = Minimum(lngRetval, mvarDR4)
                    lngRetval = Minimum(lngRetval, mvarDR5)
                    lngRetval = Minimum(lngRetval, mvarDR6)
                
                Case Else
                    lngRetval = 0
            End Select
       
       Case ArmorLocation, ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
            Select Case mvarMaterial
                Case "metal", "composite", "laminate"
                    lngRetval = mvarDR
                Case Else
                    lngRetval = 0
            End Select
          
    End Select


    GetLowestDR = lngRetval
End Function

Function GetLowestCrushDepthDR() As Long

    On Error Resume Next
    Dim lngRetval As Long
    
    Select Case mvarDatatype
       Case ArmorComplexFacing
           
                lngRetval = mvarDR1
                lngRetval = Minimum(mvarDR1, mvarDR2)
                lngRetval = Minimum(lngRetval, mvarDR3)
                lngRetval = Minimum(lngRetval, mvarDR4)
                lngRetval = Minimum(lngRetval, mvarDR5)
                lngRetval = Minimum(lngRetval, mvarDR6)
    
       Case ArmorBasicFacing
          
                lngRetval = Minimum(mvarDR1, mvarDR2)
                lngRetval = Minimum(lngRetval, mvarDR3)
                lngRetval = Minimum(lngRetval, mvarDR4)
                lngRetval = Minimum(lngRetval, mvarDR5)
                lngRetval = Minimum(lngRetval, mvarDR6)
    
       
       Case ArmorLocation, ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
               lngRetval = mvarDR
              
    End Select


    GetLowestCrushDepthDR = lngRetval
End Function

Private Sub CalcAverageDr()
Dim i As Integer
Dim TempAverage As Long
Dim Divisor As Integer
Dim arrDR(0 To 5) As Long

arrDR(0) = Val(mvarDR1)
arrDR(1) = Val(mvarDR2)
arrDR(2) = Val(mvarDR3)
arrDR(3) = Val(mvarDR4)
arrDR(4) = Val(mvarDR5)
arrDR(5) = Val(mvarDR6)

TempAverage = 0

For i = 0 To 5
    TempAverage = TempAverage + arrDR(i)
Next

'find the Average DR (this is done only for Basic Armor by facing
If mvarParent = "1_" Then
    Divisor = 6
Else
    Divisor = 5
End If

mvarAverageDR = Round(TempAverage / Divisor, 2)

End Sub


Sub CalcSurfaceFeaturesCostandWeight(SurfaceArea As Single)
Const RadWeight = 2
Const RadCost = 20
Const ReflectCost = 30
Const RetroCost = 150
Const ThermCost = 250
Const ThermWeight = 0.25
Const RAPCost = 20
Const RAPWeight = 8
Const ElectCost = 10
Const ElectWeight = 0.2
Dim TempWeight As Single
Dim TempCost As Single

If SurfaceArea = 0 Then Exit Sub

If mvarRadiation Then
    TempWeight = RadWeight * SurfaceArea
    TempCost = RadCost * SurfaceArea
End If

If mvarCoating = "reflective" Then
    TempCost = TempCost + (ReflectCost * SurfaceArea)
ElseIf mvarCoating = "retro-reflective" Then
    TempCost = TempCost + (RetroCost * SurfaceArea)
End If

If mvarThermal Then
    TempWeight = TempWeight + (ThermWeight * SurfaceArea)
    TempCost = TempCost + (ThermCost * SurfaceArea)
End If

If mvarRAP Then
    TempWeight = TempWeight + (RAPWeight * SurfaceArea)
    TempCost = TempCost + (RAPCost * SurfaceArea)
End If

If mvarElectrified Then
    TempWeight = TempWeight + (ElectWeight * SurfaceArea)
    TempCost = TempCost + (ElectCost * SurfaceArea)
End If

mvarCost = Round(mvarCost + TempCost, 2)
mvarWeight = Round(mvarWeight + TempWeight, 2)

End Sub


Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If

End Sub

Public Sub QueryChild()
' if the object has children, query it and check to see if
' more stats/property updates are needed for other objects in the collection
' (see the StatusUpdate property for help on checking for childeren.  Can i use that one in place of this?)

End Sub
