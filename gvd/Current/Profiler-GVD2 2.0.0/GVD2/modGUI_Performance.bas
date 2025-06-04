Attribute VB_Name = "modGUI_Performance"
Option Explicit

Private p_CheckListKeys() As String
Private m_sMappedKeys() As String ' since our listbox only holds indexes, we map this array by index to determine the key of the listitem
Private m_sCurrent As String
Private m_lngCheckListType As Long

Public Sub PopulateWeaponLinkCheckList(ByVal sCurrent As String)
vbwProfiler.vbwProcIn 224
    Dim element As Object


vbwProfiler.vbwExecuteLine 4413
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 4414
        Select Case element.Datatype
'vbwLine 4415:            Case StoneThrower, BoltThrower, RepeatingBoltThrower, MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling, BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam, FlameThrower, WaterCannon, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, RevolverLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher, IronBomb, RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, ProximityMine, PressureTriggerMine, CommandTriggerMine, SmartTriggerMine, UnGuidedMissile, UnGuidedTorpedo, GuidedMissile, GuidedTorpedo
            Case IIf(vbwProfiler.vbwExecuteLine(4415), VBWPROFILER_EMPTY, _
        StoneThrower), BoltThrower, RepeatingBoltThrower, MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling, BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam, FlameThrower, WaterCannon, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, RevolverLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher, IronBomb, RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, ProximityMine, PressureTriggerMine, CommandTriggerMine, SmartTriggerMine, UnGuidedMissile, UnGuidedTorpedo, GuidedMissile, GuidedTorpedo

vbwProfiler.vbwExecuteLine 4416
                AddCheckListKey element.Key, element.CustomDescription

                'check mark existing
vbwProfiler.vbwExecuteLine 4417
                Call CheckMarkExisting
        End Select
vbwProfiler.vbwExecuteLine 4418 'B
vbwProfiler.vbwExecuteLine 4419
    Next
vbwProfiler.vbwProcOut 224
vbwProfiler.vbwExecuteLine 4420
End Sub

Public Sub PopulateCheckList(ByVal sCurrent As String)
    'clear the list
vbwProfiler.vbwProcIn 225
vbwProfiler.vbwExecuteLine 4421
    frmDesigner.lstPropulsionSystems.Clear

vbwProfiler.vbwExecuteLine 4422
    m_sCurrent = m_oCurrentVeh.ActiveCheckList
vbwProfiler.vbwExecuteLine 4423
    m_lngCheckListType = m_oCurrentVeh.ActiveCheckListType

vbwProfiler.vbwExecuteLine 4424
    ReDim p_CheckListKeys(1)
vbwProfiler.vbwExecuteLine 4425
    ReDim m_sMappedKeys(0)

vbwProfiler.vbwExecuteLine 4426
     Debug.Assert m_oCurrentVeh.ActiveCheckListType > 0
vbwProfiler.vbwExecuteLine 4427
    If m_oCurrentVeh.ActiveCheckListType = PERFORMANCE_CHECKLIST Then
vbwProfiler.vbwExecuteLine 4428
        Call PopulatePerformanceCheckList(sCurrent)
    Else
vbwProfiler.vbwExecuteLine 4429 'B
vbwProfiler.vbwExecuteLine 4430
        Call PopulateWeaponLinkCheckList(sCurrent)
    End If
vbwProfiler.vbwExecuteLine 4431 'B

vbwProfiler.vbwProcOut 225
vbwProfiler.vbwExecuteLine 4432
End Sub
Public Sub PopulatePerformanceCheckList(ByVal sCurrent As String)
vbwProfiler.vbwProcIn 226


'todo: all this crap needs to be converted to check bitflag for propulsion system capabilities
'fill the global propulsion system list based on Type of Performance Profile
    Dim sPType As String 'performance type
    Dim element As Object
    Dim sMType As String 'holds classname for the Ground MotiveAssembly


    ' cycle through all objects in the Vehicle (NOTE: to optimize, a creating a keychain in the AddObject
    ' routine to track propulsion systems would be helpful) and come up with list of propulsion systems that
    ' are relevant to the performance profile type
vbwProfiler.vbwExecuteLine 4433
    For Each element In m_oCurrentVeh.Components
        '//Wheels - only display elements which can be used with Wheeled Performance
vbwProfiler.vbwExecuteLine 4434
        If TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsPerformanceWheel Then
vbwProfiler.vbwExecuteLine 4435
            Select Case element.Datatype
'vbwLine 4436:                Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
                Case IIf(vbwProfiler.vbwExecuteLine(4436), VBWPROFILER_EMPTY, _
        WheeledDrivetrain), AllWheelDriveWheeledDrivetrain, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine

vbwProfiler.vbwExecuteLine 4437
                Call AddCheckListKey(element.Key, element.CustomDescription)

            End Select
vbwProfiler.vbwExecuteLine 4438 'B
        '//Skids
'vbwLine 4439:        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceskid Then
        ElseIf vbwProfiler.vbwExecuteLine(4439) Or TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceskid Then
vbwProfiler.vbwExecuteLine 4440
            Select Case element.Datatype
'vbwLine 4441:               Case DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
               Case IIf(vbwProfiler.vbwExecuteLine(4441), VBWPROFILER_EMPTY, _
        DuctedFan), AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine

vbwProfiler.vbwExecuteLine 4442
                Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
vbwProfiler.vbwExecuteLine 4443 'B
        '//Tracks, Halftracks, Skitracks
'vbwLine 4444:        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceTrack Then
        ElseIf vbwProfiler.vbwExecuteLine(4444) Or TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceTrack Then
vbwProfiler.vbwExecuteLine 4445
            Select Case element.Datatype
'vbwLine 4446:                Case TrackedDrivetrain
                Case IIf(vbwProfiler.vbwExecuteLine(4446), VBWPROFILER_EMPTY, _
        TrackedDrivetrain)
vbwProfiler.vbwExecuteLine 4447
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
vbwProfiler.vbwExecuteLine 4448 'B

        '//Legs
'vbwLine 4449:        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceLeg Then
        ElseIf vbwProfiler.vbwExecuteLine(4449) Or TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceLeg Then
vbwProfiler.vbwExecuteLine 4450
            Select Case element.Datatype
'vbwLine 4451:                Case LegDrivetrain
                Case IIf(vbwProfiler.vbwExecuteLine(4451), VBWPROFILER_EMPTY, _
        LegDrivetrain)
vbwProfiler.vbwExecuteLine 4452
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
vbwProfiler.vbwExecuteLine 4453 'B

        '//Flexibody
'vbwLine 4454:        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceflex Then
        ElseIf vbwProfiler.vbwExecuteLine(4454) Or TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceflex Then
vbwProfiler.vbwExecuteLine 4455
            Select Case element.Datatype
'vbwLine 4456:                Case FlexibodyDrivetrain
                Case IIf(vbwProfiler.vbwExecuteLine(4456), VBWPROFILER_EMPTY, _
        FlexibodyDrivetrain)
vbwProfiler.vbwExecuteLine 4457
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
vbwProfiler.vbwExecuteLine 4458 'B

        '//WATER
'vbwLine 4459:        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancewater Then
        ElseIf vbwProfiler.vbwExecuteLine(4459) Or TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancewater Then
vbwProfiler.vbwExecuteLine 4460
            Select Case element.Datatype
'vbwLine 4461:                Case TrackedDrivetrain, LegDrivetrain, FlexibodyDrivetrain, WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, RowingPositions, PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
                Case IIf(vbwProfiler.vbwExecuteLine(4461), VBWPROFILER_EMPTY, _
        TrackedDrivetrain), LegDrivetrain, FlexibodyDrivetrain, WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, RowingPositions, PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine

vbwProfiler.vbwExecuteLine 4462
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
vbwProfiler.vbwExecuteLine 4463 'B

        '//Submerged
'vbwLine 4464:        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancesubmerged Then
        ElseIf vbwProfiler.vbwExecuteLine(4464) Or TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancesubmerged Then
vbwProfiler.vbwExecuteLine 4465
            Select Case element.Datatype
'vbwLine 4466:                Case TrackedDrivetrain, LegDrivetrain, FlexibodyDrivetrain, PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
                Case IIf(vbwProfiler.vbwExecuteLine(4466), VBWPROFILER_EMPTY, _
        TrackedDrivetrain), LegDrivetrain, FlexibodyDrivetrain, PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine

vbwProfiler.vbwExecuteLine 4467
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
vbwProfiler.vbwExecuteLine 4468 'B

        '//Aerial Performance
'vbwLine 4469:        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceair Then
        ElseIf vbwProfiler.vbwExecuteLine(4469) Or TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceair Then
vbwProfiler.vbwExecuteLine 4470
            Select Case element.Datatype
'vbwLine 4471:                Case ContraGravGenerator, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine
                Case IIf(vbwProfiler.vbwExecuteLine(4471), VBWPROFILER_EMPTY, _
        ContraGravGenerator), CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine

vbwProfiler.vbwExecuteLine 4472
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
vbwProfiler.vbwExecuteLine 4473 'B
        '//Mag-Lev Performance
'vbwLine 4474:        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancemaglev Then
        ElseIf vbwProfiler.vbwExecuteLine(4474) Or TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancemaglev Then
vbwProfiler.vbwExecuteLine 4475
            Select Case element.Datatype
'vbwLine 4476:                Case ContraGravGenerator, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine, MagLevLifter
                Case IIf(vbwProfiler.vbwExecuteLine(4476), VBWPROFILER_EMPTY, _
        ContraGravGenerator), CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine, MagLevLifter

vbwProfiler.vbwExecuteLine 4477
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
vbwProfiler.vbwExecuteLine 4478 'B
        '//Hovercraft Performance
'vbwLine 4479:        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancehover Then
        ElseIf vbwProfiler.vbwExecuteLine(4479) Or TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancehover Then
vbwProfiler.vbwExecuteLine 4480
            Select Case element.Datatype
                'note Ramjet's are not allowed since max speed for Hovercrafts is 300mph
'vbwLine 4481:                Case ContraGravGenerator, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine
                Case IIf(vbwProfiler.vbwExecuteLine(4481), VBWPROFILER_EMPTY, _
        ContraGravGenerator), CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, DuctedFan, AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness, AerialSail, AerialSailForeAftRig, Turbojet, Turbofan, TurboRamjet, Hyperfan, FusionAirRam, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine

vbwProfiler.vbwExecuteLine 4482
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
vbwProfiler.vbwExecuteLine 4483 'B
        '//Space Performance
'vbwLine 4484:        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancespace Then
        ElseIf vbwProfiler.vbwExecuteLine(4484) Or TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancespace Then
vbwProfiler.vbwExecuteLine 4485
            Select Case element.Datatype
'vbwLine 4486:                Case TeleportationDrive, Hyperdrive, JumpDrive, WarpDrive, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine, lightSail
                Case IIf(vbwProfiler.vbwExecuteLine(4486), VBWPROFILER_EMPTY, _
        TeleportationDrive), Hyperdrive, JumpDrive, WarpDrive, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine, lightSail

vbwProfiler.vbwExecuteLine 4487
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
vbwProfiler.vbwExecuteLine 4488 'B
        End If
vbwProfiler.vbwExecuteLine 4489 'B
vbwProfiler.vbwExecuteLine 4490
    Next

   'check mark existing
vbwProfiler.vbwExecuteLine 4491
   Call CheckMarkExisting

vbwProfiler.vbwProcOut 226
vbwProfiler.vbwExecuteLine 4492
End Sub


Private Sub CheckMarkExisting()
vbwProfiler.vbwProcIn 227

 'Check Mark any items that are already added
 Dim arrKeys() As String
 Dim i As Long
 Dim j As Long

vbwProfiler.vbwExecuteLine 4493
    On Error GoTo err

vbwProfiler.vbwExecuteLine 4494
    frmDesigner.lstPropulsionSystems.TAG = CHECKLIST_STATE_RESTORE

vbwProfiler.vbwExecuteLine 4495
    If m_lngCheckListType = WEAPON_CHECKLIST Then
vbwProfiler.vbwExecuteLine 4496
        arrKeys = m_oCurrentVeh.WeaponProfiles(m_sCurrent).getcurrentkeys
    Else
vbwProfiler.vbwExecuteLine 4497 'B
vbwProfiler.vbwExecuteLine 4498
        arrKeys = m_oCurrentVeh.PerformanceProfiles(m_sCurrent).getcurrentkeys
    End If
vbwProfiler.vbwExecuteLine 4499 'B

vbwProfiler.vbwExecuteLine 4500
    If arrKeys(1) = "" Then
    Else
vbwProfiler.vbwExecuteLine 4501 'B
vbwProfiler.vbwExecuteLine 4502
        For i = 1 To UBound(arrKeys)
            ' if the keys inside the performance profile match any on the list, checkmark them
            ' since they are already added
vbwProfiler.vbwExecuteLine 4503
            For j = 0 To UBound(m_sMappedKeys)
vbwProfiler.vbwExecuteLine 4504
                If arrKeys(i) = m_sMappedKeys(j) Then
vbwProfiler.vbwExecuteLine 4505
                    frmDesigner.lstPropulsionSystems.Selected(j) = True
vbwProfiler.vbwExecuteLine 4506
                    Exit For
                End If
vbwProfiler.vbwExecuteLine 4507 'B
            'frmDesigner.lstPropulsionSystems.AddItem m_oCurrentVeh.Components(arrKeys(i)).customDescription, arrKeys(i)
vbwProfiler.vbwExecuteLine 4508
            Next
vbwProfiler.vbwExecuteLine 4509
        Next
    End If
vbwProfiler.vbwExecuteLine 4510 'B
vbwProfiler.vbwExecuteLine 4511
    frmDesigner.lstPropulsionSystems.TAG = ""
vbwProfiler.vbwProcOut 227
vbwProfiler.vbwExecuteLine 4512
    Exit Sub
err:
vbwProfiler.vbwExecuteLine 4513
    frmDesigner.lstPropulsionSystems.TAG = ""
vbwProfiler.vbwProcOut 227
vbwProfiler.vbwExecuteLine 4514
End Sub
Private Sub AddCheckListKey(ByRef sKey As String, ByRef sDescription As String)
vbwProfiler.vbwProcIn 228
    Dim Count As Long

vbwProfiler.vbwExecuteLine 4515
    frmDesigner.lstPropulsionSystems.AddItem sDescription
vbwProfiler.vbwExecuteLine 4516
    p_CheckListKeys = mAddKey(p_CheckListKeys, sKey)

vbwProfiler.vbwExecuteLine 4517
    Count = UBound(m_sMappedKeys)
vbwProfiler.vbwExecuteLine 4518
    If (Count = 0) And (m_sMappedKeys(0) = "") Then
vbwProfiler.vbwExecuteLine 4519
        Count = 0
    Else
vbwProfiler.vbwExecuteLine 4520 'B
vbwProfiler.vbwExecuteLine 4521
        Count = Count + 1
    End If
vbwProfiler.vbwExecuteLine 4522 'B

vbwProfiler.vbwExecuteLine 4523
    ReDim Preserve m_sMappedKeys(Count)

vbwProfiler.vbwExecuteLine 4524
    m_sMappedKeys(Count) = sKey

vbwProfiler.vbwProcOut 228
vbwProfiler.vbwExecuteLine 4525
End Sub
Public Sub UpdatePerformanceStats()
    ' when a user adds/removes propulsion/drivetrain components from a profile, the stats
    ' need to be adjusted in real time.
vbwProfiler.vbwProcIn 229
    Dim o As Object

    ' re-calc vehicle performance figures
    'note: to optimize, i would only update figures for those profiles
    'which the user changed
vbwProfiler.vbwExecuteLine 4526
    For Each o In m_oCurrentVeh.PerformanceProfiles
vbwProfiler.vbwExecuteLine 4527
        o.CalcPerformance
vbwProfiler.vbwExecuteLine 4528
    Next


vbwProfiler.vbwProcOut 229
vbwProfiler.vbwExecuteLine 4529
End Sub

Public Sub PropulsionSelect(sCurrentProfile As String, iIndex As Long)
    'add the item to the keychain of the Current Performance Profile (see clsPerformanceXXXXXX)
vbwProfiler.vbwProcIn 230
vbwProfiler.vbwExecuteLine 4530
    If m_lngCheckListType = WEAPON_CHECKLIST Then
vbwProfiler.vbwExecuteLine 4531
        m_oCurrentVeh.WeaponProfiles(m_sCurrent).AddKey (p_CheckListKeys(iIndex + 1))
    Else
vbwProfiler.vbwExecuteLine 4532 'B
vbwProfiler.vbwExecuteLine 4533
        m_oCurrentVeh.PerformanceProfiles(m_sCurrent).AddKey (p_CheckListKeys(iIndex + 1))
    End If
vbwProfiler.vbwExecuteLine 4534 'B
vbwProfiler.vbwProcOut 230
vbwProfiler.vbwExecuteLine 4535
End Sub

Public Sub PropulsionDeSelect(sCurrentProfile As String, iIndex As Long)
    'remove the item from the keychain of the Performance Profile
vbwProfiler.vbwProcIn 231
vbwProfiler.vbwExecuteLine 4536
    If m_lngCheckListType = WEAPON_CHECKLIST Then
vbwProfiler.vbwExecuteLine 4537
        m_oCurrentVeh.WeaponProfiles(m_sCurrent).removekey (p_CheckListKeys(iIndex + 1))
    Else
vbwProfiler.vbwExecuteLine 4538 'B
vbwProfiler.vbwExecuteLine 4539
        m_oCurrentVeh.PerformanceProfiles(m_sCurrent).removekey (p_CheckListKeys(iIndex + 1))
    End If
vbwProfiler.vbwExecuteLine 4540 'B
vbwProfiler.vbwProcOut 231
vbwProfiler.vbwExecuteLine 4541
End Sub

Public Sub DeletePerformanceProfile()
vbwProfiler.vbwProcIn 232

Dim sItem As String
vbwProfiler.vbwExecuteLine 4542
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 4543
    sItem = frmDesigner.treeVehicle.SelectedItem.Text

vbwProfiler.vbwExecuteLine 4544
    If MsgBox("Are you sure you want to delete the profile '" & sItem & "' ?", vbYesNo) = vbYes Then
vbwProfiler.vbwExecuteLine 4545
        m_oCurrentVeh.PerformanceProfiles.Remove sItem ' remove the item from the collection
vbwProfiler.vbwExecuteLine 4546
        frmDesigner.treeVehicle.Nodes.Remove sItem
vbwProfiler.vbwExecuteLine 4547
        p_bChangedFlag = True  ' JAW 2000.05.07
    End If
vbwProfiler.vbwExecuteLine 4548 'B
vbwProfiler.vbwProcOut 232
vbwProfiler.vbwExecuteLine 4549
    Exit Sub
errorhandler:
vbwProfiler.vbwProcOut 232
vbwProfiler.vbwExecuteLine 4550
    Exit Sub
vbwProfiler.vbwProcOut 232
vbwProfiler.vbwExecuteLine 4551
End Sub

