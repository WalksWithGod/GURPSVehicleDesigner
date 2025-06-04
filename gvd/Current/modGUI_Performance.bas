Attribute VB_Name = "modGUI_Performance"
Option Explicit

Private p_CheckListKeys() As String
Private m_sMappedKeys() As String ' since our listbox only holds indexes, we map this array by index to determine the key of the listitem
Private m_sCurrent As String
Private m_lngCheckListType As Long

Public Sub PopulateWeaponLinkCheckList(ByVal sCurrent As String)
    Dim element As Object
    
    
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
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
        
                AddCheckListKey element.Key, element.CustomDescription
                
                'check mark existing
                Call CheckMarkExisting
        End Select
    Next
End Sub

Public Sub PopulateCheckList(ByVal sCurrent As String)
    'clear the list
    frmDesigner.lstPropulsionSystems.Clear
    
    m_sCurrent = m_oCurrentVeh.ActiveCheckList
    m_lngCheckListType = m_oCurrentVeh.ActiveCheckListType
    
    ReDim p_CheckListKeys(1)
    ReDim m_sMappedKeys(0)
    
     Debug.Assert m_oCurrentVeh.ActiveCheckListType > 0
    If m_oCurrentVeh.ActiveCheckListType = PERFORMANCE_CHECKLIST Then
        Call PopulatePerformanceCheckList(sCurrent)
    Else
        Call PopulateWeaponLinkCheckList(sCurrent)
    End If
    
End Sub
Public Sub PopulatePerformanceCheckList(ByVal sCurrent As String)

   
'todo: all this crap needs to be converted to check bitflag for propulsion system capabilities
'fill the global propulsion system list based on Type of Performance Profile
    Dim sPType As String 'performance type
    Dim element As Object
    Dim sMType As String 'holds classname for the Ground MotiveAssembly
    
    
    ' cycle through all objects in the Vehicle (NOTE: to optimize, a creating a keychain in the AddObject
    ' routine to track propulsion systems would be helpful) and come up with list of propulsion systems that
    ' are relevant to the performance profile type
    For Each element In m_oCurrentVeh.Components
        '//Wheels - only display elements which can be used with Wheeled Performance
        If TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsPerformanceWheel Then
            Select Case element.Datatype
                Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, DuctedFan, _
                AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, _
                WhiffletreeHarness, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, _
                Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, _
                LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                OptimizedFusion, AntimatterThermal, AntimatterPion, _
                StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
                
                Call AddCheckListKey(element.Key, element.CustomDescription)
                
            End Select
        '//Skids
        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceskid Then
            Select Case element.Datatype
               Case DuctedFan, AerialPropeller, RopeHarness, _
                YokeandPoleHarness, ShaftandCollarHarness, _
                WhiffletreeHarness, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, _
                Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, _
                LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                OptimizedFusion, AntimatterThermal, AntimatterPion, _
                StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
                
                Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
        '//Tracks, Halftracks, Skitracks
        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceTrack Then
            Select Case element.Datatype
                Case TrackedDrivetrain
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
            
        '//Legs
        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceLeg Then
            Select Case element.Datatype
                Case LegDrivetrain
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
            
        '//Flexibody
        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceflex Then
            Select Case element.Datatype
                Case FlexibodyDrivetrain
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
            
        '//WATER
        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancewater Then
            Select Case element.Datatype
                Case TrackedDrivetrain, LegDrivetrain, FlexibodyDrivetrain, _
                    WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, RowingPositions, _
                    PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, _
                    Hydrojet, MHDTunnel, DuctedFan, AerialPropeller, RopeHarness, _
                    YokeandPoleHarness, ShaftandCollarHarness, _
                    WhiffletreeHarness, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, _
                    Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, _
                    LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                    OptimizedFusion, AntimatterThermal, AntimatterPion, _
                    StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
                    
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
            
        '//Submerged
        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancesubmerged Then
            Select Case element.Datatype
                Case TrackedDrivetrain, LegDrivetrain, FlexibodyDrivetrain, _
                    PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, _
                    Hydrojet, MHDTunnel, RopeHarness, _
                    YokeandPoleHarness, ShaftandCollarHarness, _
                    WhiffletreeHarness, _
                    LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                    OptimizedFusion, AntimatterThermal, AntimatterPion, _
                    StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
                    
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
            
        '//Aerial Performance
        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformanceair Then
            Select Case element.Datatype
                Case ContraGravGenerator, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, _
                    OrnithopterDrivetrain, DuctedFan, _
                    AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, _
                    WhiffletreeHarness, AerialSail, AerialSailForeAftRig, _
                    Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, _
                    LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                    OptimizedFusion, AntimatterThermal, AntimatterPion, _
                    StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine
                    
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
        '//Mag-Lev Performance
        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancemaglev Then
            Select Case element.Datatype
                Case ContraGravGenerator, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, _
                    OrnithopterDrivetrain, DuctedFan, _
                    AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, _
                    WhiffletreeHarness, AerialSail, AerialSailForeAftRig, _
                    Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, _
                    LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                    OptimizedFusion, AntimatterThermal, AntimatterPion, _
                    StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine, MagLevLifter
                    
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
        '//Hovercraft Performance
        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancehover Then
            Select Case element.Datatype
                'note Ramjet's are not allowed since max speed for Hovercrafts is 300mph
                Case ContraGravGenerator, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, _
                    OrnithopterDrivetrain, DuctedFan, _
                    AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, _
                    WhiffletreeHarness, AerialSail, AerialSailForeAftRig, _
                    Turbojet, Turbofan, TurboRamjet, Hyperfan, FusionAirRam, _
                    LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                    OptimizedFusion, AntimatterThermal, AntimatterPion, _
                    StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine
                    
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
        '//Space Performance
        ElseIf TypeOf m_oCurrentVeh.PerformanceProfiles(m_sCurrent) Is clsperformancespace Then
            Select Case element.Datatype
                Case TeleportationDrive, Hyperdrive, JumpDrive, WarpDrive, _
                LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                    OptimizedFusion, AntimatterThermal, AntimatterPion, _
                    StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine, lightSail
                    
                    Call AddCheckListKey(element.Key, element.CustomDescription)
            End Select
        End If
    Next
       
   'check mark existing
   Call CheckMarkExisting
   
End Sub


Private Sub CheckMarkExisting()

 'Check Mark any items that are already added
 Dim arrKeys() As String
 Dim i As Long
 Dim j As Long
 
    On Error GoTo err
    
    frmDesigner.lstPropulsionSystems.TAG = CHECKLIST_STATE_RESTORE
    
    If m_lngCheckListType = WEAPON_CHECKLIST Then
        arrKeys = m_oCurrentVeh.WeaponProfiles(m_sCurrent).getcurrentkeys
    Else
        arrKeys = m_oCurrentVeh.PerformanceProfiles(m_sCurrent).getcurrentkeys
    End If
    
    If arrKeys(1) = "" Then
    Else
        For i = 1 To UBound(arrKeys)
            ' if the keys inside the performance profile match any on the list, checkmark them
            ' since they are already added
            For j = 0 To UBound(m_sMappedKeys)
                If arrKeys(i) = m_sMappedKeys(j) Then
                    frmDesigner.lstPropulsionSystems.Selected(j) = True
                    Exit For
                End If
            'frmDesigner.lstPropulsionSystems.AddItem m_oCurrentVeh.Components(arrKeys(i)).customDescription, arrKeys(i)
            Next
        Next
    End If
    frmDesigner.lstPropulsionSystems.TAG = ""
    Exit Sub
err:
    frmDesigner.lstPropulsionSystems.TAG = ""
End Sub
Private Sub AddCheckListKey(ByRef sKey As String, ByRef sDescription As String)
    Dim Count As Long
    
    frmDesigner.lstPropulsionSystems.AddItem sDescription
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
    ' when a user adds/removes propulsion/drivetrain components from a profile, the stats
    ' need to be adjusted in real time.
    Dim o As Object
    
    ' re-calc vehicle performance figures
    'note: to optimize, i would only update figures for those profiles
    'which the user changed
    For Each o In m_oCurrentVeh.PerformanceProfiles
        o.CalcPerformance
    Next
        
        
End Sub

Public Sub PropulsionSelect(sCurrentProfile As String, iIndex As Long)
    'add the item to the keychain of the Current Performance Profile (see clsPerformanceXXXXXX)
    If m_lngCheckListType = WEAPON_CHECKLIST Then
        m_oCurrentVeh.WeaponProfiles(m_sCurrent).AddKey (p_CheckListKeys(iIndex + 1))
    Else
        m_oCurrentVeh.PerformanceProfiles(m_sCurrent).AddKey (p_CheckListKeys(iIndex + 1))
    End If
End Sub

Public Sub PropulsionDeSelect(sCurrentProfile As String, iIndex As Long)
    'remove the item from the keychain of the Performance Profile
    If m_lngCheckListType = WEAPON_CHECKLIST Then
        m_oCurrentVeh.WeaponProfiles(m_sCurrent).removekey (p_CheckListKeys(iIndex + 1))
    Else
        m_oCurrentVeh.PerformanceProfiles(m_sCurrent).removekey (p_CheckListKeys(iIndex + 1))
    End If
End Sub

Public Sub DeletePerformanceProfile()

Dim sItem As String
    On Error GoTo errorhandler
    
    sItem = frmDesigner.treeVehicle.SelectedItem.Text
    
    If MsgBox("Are you sure you want to delete the profile '" & sItem & "' ?", vbYesNo) = vbYes Then
        m_oCurrentVeh.PerformanceProfiles.Remove sItem ' remove the item from the collection
        frmDesigner.treeVehicle.Nodes.Remove sItem
        p_bChangedFlag = True  ' JAW 2000.05.07
    End If
    Exit Sub
errorhandler:
    Exit Sub
End Sub
