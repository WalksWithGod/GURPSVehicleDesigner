Attribute VB_Name = "modTextOutput"
Option Explicit
Private sBreak As String
Private sLineBreak As String
Private bSlimline As Boolean

Private Const CREATED_WITH = "Created with GURPS Vehicle Designer 2.0"
Private Const GVD_URL = "http://www.makosoft.com/gvd"

    
Public Function createGURPSText(ByVal sType As String) As String
    Dim sOutput As String
    Dim sTemp As String
    Dim sTagline As String
    
    #If DEBUG_MODE Then
        MsgBox "modTextOutput:createGURPSText() - Function not available in Debug Mode."
        Exit Function
    #End If
    
    'jaw 2000.06.25
    'reformed to select case to allow for additional exports to be easily added
    Select Case sType
        Case "Text"
            sBreak = Chr(13) + Chr(10) + Chr(13) + Chr(10)
            sTagline = CREATED_WITH + Chr(13) + Chr(10) + GVD_URL
            sLineBreak = Chr(13) + Chr(10)
        Case "Text Slim"
            sBreak = Chr(13) + Chr(10) '+ Chr(13) + Chr(10)
            sTagline = CREATED_WITH + Chr(13) + Chr(10) + GVD_URL + Chr(13) + Chr(10)
            sLineBreak = Chr(13) + Chr(10)
            bSlimline = True
        Case "Class HTML"
            sBreak = "<BR> <BR>" & vbNewLine & vbNewLine
            sTagline = CREATED_WITH & "<BR>" & GVD_URL & vbCrLf & "</body></html>"
            sLineBreak = "<BR>"
        Case "New HTML"
            sBreak = "<BR> <BR>" & vbNewLine & vbNewLine
            sTagline = CREATED_WITH & "<BR>" & GVD_URL & vbCrLf & "</body></html>"
            sLineBreak = "<BR>"
        Case Else
            Exit Function
    End Select
    On Error Resume Next
    'get header, vehicle name, copyright info and description
    sOutput = GetHeaderOutput(sType)
    sOutput = sOutput + "Subassemblies and Body Features: " + GetSubassemblyOutput + GetBodyFeatures + sBreak
    sTemp = GetCustomComponentsOutput
    If sTemp <> "" Then sOutput = sOutput + "Custom Components: " + sTemp + sBreak
    sTemp = GetPropulsionOutput
    If sTemp <> "" Then sOutput = sOutput + "Propulsion: " + sTemp + sBreak
    sTemp = GetAerostaticLiftOutput
    If sTemp <> "" Then sOutput = sOutput + "Aerostatic Lift: " + sTemp + sBreak
    sTemp = GetWeaponryOutput
    If sTemp <> "" Then sOutput = sOutput + "Weaponry: " + sTemp + sBreak
    sTemp = GetWeaponLinksOutput
    If sTemp <> "" Then sOutput = sOutput + "Weapon Links: " + sTemp + sBreak
    sTemp = GetWeaponAccessoriesOutput
    If sTemp <> "" Then sOutput = sOutput + "Weapon Accessories: " + sTemp + sBreak
    sTemp = GetCommunicationsOutput
    If sTemp <> "" Then sOutput = sOutput + "Communications: " + sTemp + sBreak
    sTemp = GetSensorsOutput
    If sTemp <> "" Then sOutput = sOutput + "Sensors: " + sTemp + sBreak
    sTemp = GetAudioVisualOutput
    If sTemp <> "" Then sOutput = sOutput + "Audio/Visual: " + sTemp + sBreak
    sTemp = GetNavigationOutput
    If sTemp <> "" Then sOutput = sOutput + "Navigation: " + sTemp + sBreak
    sTemp = GetTargetingOutput
    If sTemp <> "" Then sOutput = sOutput + "Targeting: " + sTemp + sBreak
    sTemp = GetECMOutput
    If sTemp <> "" Then sOutput = sOutput + "ECM: " + sTemp + sBreak
    sTemp = GetComputersOutput
    If sTemp <> "" Then sOutput = sOutput + "Computers: " + sTemp + sBreak
    sTemp = GetSoftwareOutput
    If sTemp <> "" Then sOutput = sOutput + "Software: " + sTemp + sBreak
    sTemp = GetMiscellaneousOutput
    If sTemp <> "" Then sOutput = sOutput + "Miscellaneous: " + sTemp + sBreak
    sTemp = GetVehicleControlsOutput
    If sTemp <> "" Then sOutput = sOutput + "Vehicle Controls: " + sTemp + sBreak
    sTemp = GetNeuralInterfaceSystemOutput
    If sTemp <> "" Then sOutput = sOutput + "Neural Interfaces: " + sTemp + sBreak
    sTemp = GetCrewStationsOutput
    If sTemp <> "" Then sOutput = sOutput + "Crew Stations: " + sTemp + sBreak
    sTemp = GetOccupancyOutput
    If sTemp <> "" Then sOutput = sOutput + "Occupancy: " + sTemp + sBreak
    sTemp = GetAccomodationsOutput
    If sTemp <> "" Then sOutput = sOutput + "Accommodations: " + sTemp + sBreak
    sTemp = GetEnvironmentalSystemsOutput
    If sTemp <> "" Then sOutput = sOutput + "Environmental Systems: " + sTemp + sBreak
    sTemp = GetSafetySystemsOutput
    If sTemp <> "" Then sOutput = sOutput + "Safety Systems: " + sTemp + sBreak
    sTemp = GetPowerSystemsOutPut
    If sTemp <> "" Then sOutput = sOutput + "Power Systems: " + sTemp + sBreak
    sTemp = GetFuelOutput
    If sTemp <> "" Then sOutput = sOutput + "Fuel: " + sTemp + sBreak
    sTemp = GetSpaceOutput
    If sTemp <> "" Then sOutput = sOutput + sTemp + sBreak
    sTemp = GetSurfaceAreaOutput
    If sTemp <> "" Then sOutput = sOutput + "Surface Area: " + sTemp + sBreak
    sTemp = GetStructureOutput
    If sTemp <> "" Then sOutput = sOutput + "Structure: " + sTemp + sBreak
    sTemp = GetHitPointsOutput
    If sTemp <> "" Then sOutput = sOutput + "Hit Points: " + sTemp + sBreak
    sTemp = GetStructuralOptionsOutput
    If sTemp <> "" Then sOutput = sOutput + sTemp + sBreak
    sTemp = GetArmorOutput
    If sTemp <> "" Then sOutput = sOutput + sTemp + sBreak
    sTemp = GetSurfaceFeaturesOutput
    If sTemp <> "" Then sOutput = sOutput + sTemp + sBreak
    sTemp = GetDefensiveSurfaceFeatures
    If sTemp <> "" Then sOutput = sOutput + sTemp + sBreak
    sTemp = GetOtherSurfaceFeatures
    If sTemp <> "" Then sOutput = sOutput + sTemp + sBreak
    sTemp = GetTopDeckSurfaceFeatures
    If sTemp <> "" Then sOutput = sOutput + sTemp + sBreak
    sTemp = GetWeaponBaysAndHardpoints
    If sTemp <> "" Then sOutput = sOutput + sTemp + sBreak
    sTemp = GetVisionAndDetailsOutput
    If sTemp <> "" Then sOutput = sOutput + sTemp + sBreak
    sTemp = GetStatisticsOutput
    If sTemp <> "" Then sOutput = sOutput + "Statistics: " + sTemp + sBreak
    sTemp = GetPerformanceOutput
    If sTemp <> "" Then sOutput = sOutput + sTemp + sBreak
    sTemp = GetDetailedWeaponStats
    If sTemp <> "" Then sOutput = sOutput + sTemp + vbNewLine
    
    '//add our tag line
    sOutput = sOutput + sTagline
    'jaw 2000.06.25
    If bSlimline Then
        sOutput = RemoveParenthetical(sOutput)
    End If
    createGURPSText = sOutput
End Function

Private Function GetHeaderOutput(ByVal sType As String) As String
    Dim sOutput As String
    Dim sTemp As String
'    With m_oCurrentVeh.Description
'        sTemp = .Header
'        If sTemp <> "" Then sOutput = sTemp + sBreak
'        sTemp = .NickName
'        If sTemp <> "" Then sOutput = sOutput + "Name: " + sTemp + sLineBreak
'        sTemp = .ClassName
'        If sTemp <> "" Then sOutput = sOutput + "Class: " & sTemp + sLineBreak
'        sTemp = .category
'        If sTemp <> "" Then
'            If .subcategory <> "" Then
'                sOutput = sOutput + "Category: " & sTemp & "  Subcategory: " & .subcategory & sLineBreak
'            Else
'                sOutput = sOutput + "Category: " & sTemp & sLineBreak
'            End If
'        End If
'
'        sTemp = .CopyrightDate
'        If sTemp <> "" Then sOutput = sOutput + "Copyright (c) " + sTemp + sLineBreak
'        sTemp = .author
'        If sTemp <> "" Then
'            sOutput = sOutput + "by " + sTemp
'
'            sTemp = .email
'            If sTemp <> "" Then
'                sOutput = sOutput + " " + "<" + sTemp + ">" + sLineBreak
'            Else
'                sOutput = sOutput + sLineBreak
'            End If
'       End If
'        sTemp = .url
'        If sTemp <> "" Then sOutput = sOutput + "http://" + sTemp + sLineBreak
'
'        sTemp = .VehicleDescription
'        If sTemp <> "" Then sOutput = sOutput + sLineBreak + sTemp + sBreak
'
'    End With
'
'    'JAW 2000.06.25
'    'change header to include doc head tags for HTML
'    Select Case sType
'        Case "Class HTML", "New HTML"
'            GetHeaderOutput = "<html><head><title>" & m_oCurrentVeh.Description.Header _
'                & ", " & m_oCurrentVeh.Description.ClassName & "-class " & _
'                m_oCurrentVeh.Description.category & " " & m_oCurrentVeh.Description.subcategory _
'                & "</title></head><body>" & vbCrLf & sOutput
'        Case Else
'            GetHeaderOutput = sOutput
'    End Select
End Function

Private Function CHOPCHOP(ByVal s As String) As String
    Dim GooGooGaga As Long
    Dim ScoobyDoo As Boolean
    Dim tempbyte() As Byte
    Dim bFlag As Boolean
    Dim i As Long
    On Error GoTo errorhandler
    '//this routine mangles the Print Output if the program is not registered
    Randomize
    tempbyte = ChopCheck
    If (IsEmpty(tempbyte) = False) And (UBound(tempbyte) - LBound(tempbyte) = UBound(gsRegNum) - LBound(gsRegNum)) Then
        For i = 1 To UBound(gsRegNum)
            If tempbyte(i) = gsRegNum(i) Then
                bFlag = True
            Else
                bFlag = False
                Exit For
            End If
        Next
        
        If bFlag Then
            CHOPCHOP = s
            Exit Function
        ElseIf ScoobyDoo Then
            GoTo ScoobySnack
        Else
        
        End If
    End If
    '//mangle time
    If Len(s) <= 1 Then
        CHOPCHOP = ""
    Else
        For i = 1 To Len(s)
            Mid(s, i, 1) = Chr(Int((255 - 0 + 1) * Rnd))
        Next
    End If
    
    '//add some fake code that never gets executed
    If GooGooGaga = -74439050 Then
        GooGooGaga = 85858859
    End If
    CHOPCHOP = s
    If ScoobyDoo Then GoTo ScoobySnack
    Exit Function
ScoobySnack:
        'note that this never gets called because ScoobyDoo always evaluates to False
        Resume Next
errorhandler:
End Function

Public Function ChopCheck() As Byte()
    Dim tempbyte() As Byte
    Dim i As Long
    Dim j As Single
    Dim sTName As String
    Dim lngtotal As Single
    Dim sRegNumber As String
    
    ReDim tempbyte(1)

#If DEBUG_MODE Then
    MsgBox "modTextOutput:ChopCheck() -- Function not available in debug mode."
    Exit Function
#End If
    On Error GoTo errorhandler
    '//one of the local reg number checkers.  There will be several of these so
    ' so that a hacker will have to do some serious code hacking to disable all
    ' of them
    
    'here's the reg key formula
    '1- the user's reg name and key are accepted into a byte array with each
    '   letter being actually the ascii code for that letter.  Total them up
    For i = 1 To UBound(gsRegName)
        lngtotal = lngtotal + gsRegName(i)
        'at the same time total the ascii value for every even valued ascii code
        If gsRegName(i) Mod 2 = 0 Then
            lngtotal = lngtotal + gsRegName(i)
        End If
    Next
    '2 - the RegID is actually just a modifier to prevent two people having the same
    '    name winding up with the same ID.  This ID is unique and alone can be used
    '   to identify a user.  Multiply this to the total
    lngtotal = lngtotal * gsRegID
    '3- take the ascii value of the typename of the Body and multiply that to it
    sTName = TypeName(m_oCurrentVeh.Body) '(BODY_KEY))
    For i = 1 To Len(sTName)
        lngtotal = lngtotal * Asc(Mid(sTName, i, 1))
    Next
    '6- take a random seed to generate the seeded random number and multiply that
    Rnd -1
    Randomize 9921988
    lngtotal = lngtotal * Rnd()
    '8- return this as a byte array that we can compare with our current one
    'how do we split this up into seperate bytes? well we know our ascii values
    'must be between 48-57, 65-90 and 97-122
    'well, we can generate a random reg code based on each number in the string
    'representation using the random seed of each number
    For i = 1 To Len(Str(lngtotal))
        j = Rnd()
        If j <= 0.33 Then
            ReDim Preserve tempbyte(i)
            Rnd -1
            Randomize Asc(Mid(Str(lngtotal), i, 1))
            tempbyte(i) = Int((57 - 48 + 1) * Rnd + 48)
            sRegNumber = sRegNumber & Chr(tempbyte(i))
        ElseIf j <= 0.66 Then
            ReDim Preserve tempbyte(i)
            Rnd -1
            Randomize Asc(Mid(Str(lngtotal), i, 1))
            tempbyte(i) = Int((90 - 65 + 1) * Rnd + 65)
            sRegNumber = sRegNumber & Chr(tempbyte(i))
        Else
            ReDim Preserve tempbyte(i)
            Rnd -1
            Randomize Asc(Mid(Str(lngtotal), i, 1))
            tempbyte(i) = Int((122 - 97 + 1) * Rnd + 97)
            sRegNumber = sRegNumber & Chr(tempbyte(i))
        End If
    Next
    
    ChopCheck = tempbyte
    Exit Function
errorhandler:
        ReDim tempbyte(1)
        ChopCheck = tempbyte
End Function

Private Function GetSubassemblyOutput() As String
    Dim sOutput As String
    Dim element As Object
    On Error GoTo err

'todo: fix
'    For Each element In m_oCurrentVeh.Components
'        Select Case element.Datatype
'
'            Case Wheel, Skid, Track, Hydrofoil, Hovercraft, _
'                Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, _
'                Wing, Mast, Superstructure, Turret, Popturret, _
'                OpenMount, Gasbag, Pod, SolarPanel, equipmentPod
'
'                sOutput = sOutput + element.PrintOutput + " "
'
'        End Select
'    Next
    
    GetSubassemblyOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetSubassemblyOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetCustomComponentsOutput() As String
     Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
'    For Each element In m_oCurrentVeh.Components
'        If TypeOf element Is clsSimpleCustom Then
'           sOutput = sOutput + element.PrintOutput + " "
'        End If
'    Next
'
'    GetCustomComponentsOutput = sOutput
'Exit Function
err:
    Debug.Print "modTextOutput:GetCustomComponentsOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetPropulsionOutput() As String
    Dim sOutput As String
    Dim element As Object
    
'    On Error GoTo err
'    For Each element In m_oCurrentVeh.Components
'        Select Case element.Datatype
'
'            Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, _
'                FlexibodyDrivetrain, TrackedDrivetrain, LegDrivetrain, _
'                CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, _
'                OrnithopterDrivetrain, AerialPropeller, DuctedFan, PaddleWheel, _
'                ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, _
'                MHDTunnel, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, _
'                WhiffletreeHarness, MagLevLifter, Turbojet, Turbofan, Ramjet, _
'                TurboRamjet, Hyperfan, FusionAirRam, StandardThruster, _
'                SuperThruster, MegaThruster, LiquidFuelRocket, MOXRocket, _
'                IonDrive, FissionRocket, FusionRocket, OptimizedFusion, _
'                AntimatterThermal, AntimatterPion, RowingPositions, ForeandAftRig, _
'                SquareRig, FullRig, AerialSail, AerialSailForeAftRig, lightSail, SolidRocketEngine, _
'                OrionEngine, TeleportationDrive, Hyperdrive, JumpDrive, _
'                WarpDrive, QuantumConveyor, SubQuantumConveyor, _
'                TwoQuantumConveyor
'
'                sOutput = sOutput + element.PrintOutput + " "
'
'        End Select
'    Next
'
'    GetPropulsionOutput = sOutput
'Exit Function
err:
    Debug.Print "modTextOutput:GetPropulsionOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetAerostaticLiftOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
'    For Each element In m_oCurrentVeh.Components
'        Select Case element.Datatype
'
'            Case ContraGravGenerator, HotAir, Hydrogen, Helium
'
'                sOutput = sOutput + element.PrintOutput + " "
'
'        End Select
'    Next
    
    GetAerostaticLiftOutput = sOutput
Exit Function
err:
    Debug.Print "modTextOutput:GetAerostaticLiftOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetWeaponryOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case StoneThrower, BoltThrower, RepeatingBoltThrower, _
                MuzzleLoader, BreechLoader, ManualRepeater, Revolver, _
                MechanicalGatling, SlowAutoloader, FastAutoloader, _
                lightAutomatic, HeavyAutomatic, ElectricGatling, BlueGreenLaser, _
                RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, _
                ChargedParticleBeam, NeutralParticleBeam, _
                Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, _
                FusionBeam, GravityBeam, AntiparticleBeam, Graser, _
                Disintegrator, Displacer, BeamedPowerTransmitter, _
                MilitaryParalysisBeam, EnergyDrill, IronBomb, _
                RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, _
                ProximityMine, PressureTriggerMine, CommandTriggerMine, _
                SmartTriggerMine, UnGuidedMissile, UnGuidedTorpedo, _
                GuidedMissile, GuidedTorpedo, FlameThrower, WaterCannon, _
                DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, _
                ManualRepeaterLauncher, SlowAutoLoaderLauncher, _
                FastAutoLoaderLauncher, RevolverLauncher, _
                lightAutomaticLauncher, HeavyAutomaticLauncher
    
                sOutput = sOutput + element.PrintOutput + " "
        
        End Select
    Next
    
    For Each element In m_oCurrentVeh.Components
        If TypeOf element Is clsweaponAmmunition Then
            sOutput = sOutput + element.PrintOutput + " "
            
        End If
    Next
    
    GetWeaponryOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetWeaponryOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetWeaponAccessoriesOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case PartialStabilizationGear, FullStabilizationGear, _
                UniversalMount, CasemateMount, DoorMount, Cyberslave, _
                AntiBlastMagazine
                
                sOutput = sOutput + element.PrintOutput + " "
        
        End Select
    Next
    
    GetWeaponAccessoriesOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetWeaponAccessoriesOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetWeaponLinksOutput() As String
    Dim sOutput As String
    Dim element As Object
    Dim sKeyArray() As String
    Dim i As Long
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        If TypeOf element Is clsWeaponLink Then
            '//append the weapons that are in the link
                sKeyArray = element.getcurrentkeys
                If sKeyArray(1) = "" Then
                Else
                    sOutput = sOutput + element.Key & " controls "
                    For i = 1 To UBound(sKeyArray)
                        sOutput = sOutput + m_oCurrentVeh.Components(sKeyArray(i)).Description & ", "
                    Next
                    '//delete the last "," and replace it with "."
                    sOutput = Left(sOutput, Len(sOutput) - 2)
                    sOutput = sOutput + ".  "
        
                End If
        
        End If
    Next
    
    GetWeaponLinksOutput = sOutput
Exit Function
err:
    Debug.Print "modTextOutput:GetWeaponLinksOutput -- Error #" & err.Number & " " & err.Description
    
End Function
Private Function GetCommunicationsOutput() As String
    Dim sOutput As String
    Dim element As Object
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case RadioDirectionFinder, RadioCommunicator, _
            TightBeamRadio, VLFRadio, CellularPhone, _
            CellularPhonewithRadio, RadioJammer, ElfReceiver, _
            LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator
            
                sOutput = sOutput + element.PrintOutput + " "
        
        End Select
    Next
    
    GetCommunicationsOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetCommunicationsOutput -- Error #" & err.Number & " " & err.Description
    
End Function
Private Function GetSensorsOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case Headlight, Searchlight, InfraredSearchlight, _
            AstronomicalInstruments, Telescope, lightAmplification, _
            LowlightTV, ExtendableSensorPeriscope, Radar, Ladar, _
            NavigationalRadar, AntiCollisionRadar, AESA, LowResImagingRadar, _
            HiResImagingRadar, ActiveSonar, PassiveSonar, PassiveInfrared, _
            Thermograph, PassiveRadar, PESA, Geophone, MAD, MultiScanner, _
            ChemScanner, RadScanner, BioScanner, GravScanner, _
            RangingSoundDetector, SurveillanceSoundDetector, _
            MeteorologicalInstruments, LowResPlanetarySurveyArray, _
            MedResPlanetarySurveyArray, HighResPlanetarySurveyArray
                
                sOutput = sOutput + element.PrintOutput + " "
        
        End Select
    Next
    
    GetSensorsOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetSensorsOutput -- Error #" & err.Number & " " & err.Description
    
End Function
Private Function GetAudioVisualOutput() As String
    Dim sOutput As String
    Dim element As Object
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case SoundSystem, FlightRecorder, VehicleCamera, DigitalVehicleCamera, _
                ReconCamera, DigitalReconCamera

                sOutput = sOutput + element.PrintOutput + " "
        
        End Select
    Next
    
    GetAudioVisualOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetAudioVisualOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetNavigationOutput() As String
    Dim sOutput As String
    Dim element As Object
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case NavigationInstruments, AutoPilot, IFF, Transponder, INS, _
                GPS, MilitaryGPS, TFR
                
                sOutput = sOutput + element.PrintOutput + " "
        
        End Select
    Next
    
    GetNavigationOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetNavigationOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetTargetingOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case ImprovedOpticalBombSight, AdvancedOpticalBombSight, _
                OpticalBombSight, FireDirectionCenter, HUDWAC, PupilHUDWAC, _
                LaserRangeFinder, LaserDesignator, LaserSpotTracker
               
                sOutput = sOutput + element.PrintOutput + " "
        
        End Select
    Next
    
    GetTargetingOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetTargetingOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetECMOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case RadarDetector, LaserSensor, LaserRadarDetector, AreaRadarJammer, _
                DeceptiveRadarJammer, InfraredJammer, ChaffDecoyDischarger, _
                SmokeDecoyDischarger, FlareDecoyDischarger, SonarDecoyDischarger, _
                HotSmokeDecoyDischarger, PrismDecoyDischarger, _
                BlackOutGasDecoyDischarger, RadarReflector, BlipEnhancer, TEMPEST

                sOutput = sOutput + element.PrintOutput + " "
        
        End Select
    Next
    
    GetECMOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetECMOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetComputersOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case MacroFrame, MainFrame, MicroFrame, MiniComputer, SmallComputer, _
                Terminal
                
               sOutput = sOutput + element.PrintOutput + " "
        
        End Select
    Next
    
    GetComputersOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetComputersOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetSoftwareOutput() As String
     Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        If TypeOf element Is clssoftware Then
            sOutput = sOutput + element.PrintOutput + " "
        End If
    Next
    
    
GetSoftwareOutput = sOutput
Exit Function
err:
    Debug.Print "modTextOutput:GetSoftwareOutput -- Error #" & err.Number & " " & err.Description
    
End Function


Private Function GetNeuralInterfaceSystemOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err

     For Each element In m_oCurrentVeh.Components
        If TypeOf element Is clsneuralinterfacesystem Then
            sOutput = sOutput + element.PrintOutput + " "
        End If
    Next
    
    GetNeuralInterfaceSystemOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetNeuralInterfaceSystemOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetMiscellaneousOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case ArmMotor, FireExtinguisherSystem, FullFireSuppressionSystem, _
                CompactFireSuppressionSystem, BilgePump, CompleteWorkshop, _
                MechanicWorkshop, EngineeringWorkshop, ElectronicsWorkshop, _
                ArmouryWorkshop, CompleteMiniWorkshop, ScienceLab, _
                MiniMechanicWorkshop, MiniElectronicsWorkshop, _
                MiniEngineeringWorkshop, MiniArmouryWorkshop
                
                sOutput = sOutput + element.PrintOutput + " "
    
            Case ExtendableLadder, Crane, Winch, PowerShovel, WreckingCrane, ForkLift, _
                VehicularBridge, LaunchCatapult, SkyHook, Bore, SuperBore, EnergyDrill, _
                TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet, _
                OperatingRoom, StretcherPallet, EmergencySupportUnit, _
                EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable, _
                Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, _
                MovieScreenandProjectorSmall, HoloventureZone, CargoRamp, _
                Airlock, MembraneAirlock, Forcelock, PassageTube, ArmoredPassageTube, _
                TeleportProjector, BrigsandRestraints, BurglarAlarm, _
                HighSecurityAlarm, MutableLicensePlate, OilSprayer, PaintSprayer, _
                SmokeScreen, SpikeDropper, VehicleBay, HangerBay, DryDock, SpaceDock, _
                ExternalCradle, ArrestorHook, VehicularParachute, RefuellingProbe, _
                RefuellingDrogue, FuelElectrolysisSystem, HydrogenFuelScoop, _
                AtmosphereProcessor, NuclearDamper, SmallRealityStabilizer, _
                MediumRealityStabilizer, HeavyRealityStabilizer, ModularSocket, Module

                sOutput = sOutput + element.PrintOutput + " "
        
        End Select
    Next
    
    GetMiscellaneousOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetMiscellaneousOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetVehicleControlsOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case PrimitiveManeuverControl, ElectronicDivingControl, _
                ComputerizedDivingControl, MechanicalManeuverControl, _
                ElectronicManeuverControl, ComputerizedManeuverControl, _
                MechanicalDivingControl
    
                sOutput = sOutput + element.PrintOutput + " "
        
        End Select
    Next
    
    GetVehicleControlsOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetVehicleControlsOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetCrewStationsOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case CrampedCrewStation, NormalCrewStation, RoomyCrewStation, _
                CycleCrewStation, HarnessCrewStation
    
                
                sOutput = sOutput + element.PrintOutput + " "
        
        End Select
    Next
    
    GetCrewStationsOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetCrewStationsOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetOccupancyOutput() As String
   
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    With m_oCurrentVeh.crew
         If .numshifts > 1 Then sOutput = NumericToString(.numshifts) & " shifts. "
        If .numcaptains > 0 Then sOutput = sOutput & NumericToString(.numcaptains) & " captains. "
        If .NumOfficers > 0 Then sOutput = sOutput & NumericToString(.NumOfficers) & " officers. "
        If .NumCrewStationOperators > 0 Then sOutput = sOutput & NumericToString(.NumCrewStationOperators) & " crew station operators. "
        If .NumWeaponLoaders > 0 Then sOutput = sOutput & NumericToString(.NumWeaponLoaders) & " weapon loaders. "
        If .NumRowers > 0 Then sOutput = sOutput & NumericToString(.NumRowers) & " rowers. "
        If .NumSailors > 0 Then sOutput = sOutput & NumericToString(.NumSailors) & " sailors. "
        If .NumRiggers > 0 Then sOutput = sOutput & NumericToString(.NumRiggers) & " sail riggers. "
        If .NumFuelStokers > 0 Then sOutput = sOutput & NumericToString(.NumFuelStokers) & " fuel stokers. "
        If .NumMechanics > 0 Then sOutput = sOutput & NumericToString(.NumMechanics) & " mechanics. "
        If .NumServiceCrewmen > 0 Then sOutput = sOutput & NumericToString(.NumServiceCrewmen) & " service crewmen. "
        If .NumMedics > 0 Then sOutput = sOutput & NumericToString(.NumMedics) & " medics. "
        If .NumScientists > 0 Then sOutput = sOutput & NumericToString(.NumScientists) & " scientists. "
        If .NumAuxiliaryVehicleCrew > 0 Then sOutput = sOutput & NumericToString(.NumAuxiliaryVehicleCrew) & " auxiliary vehicle crewmen. "
        If .NumStewards > 0 Then sOutput = sOutput & NumericToString(.NumStewards) & " stewards. "
        If .NumLuxury > 0 Then sOutput = sOutput & NumericToString(.NumLuxury) & " luxury class passengers. "
        If .NumFirstClass > 0 Then sOutput = sOutput & NumericToString(.NumFirstClass) & " first class passengers. "
        If .NumSecondClass > 0 Then sOutput = sOutput & NumericToString(.NumSecondClass) & " second class passengers. "
        If .NumSteerage > 0 Then sOutput = sOutput & NumericToString(.NumSteerage) & " steerage passengers. "
        
        ' append whether its short or long Occupancy
        sOutput = .Occupancy & ". " & sOutput
   End With
                            
    GetOccupancyOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetOccupancyOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetAccomodationsOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case CrampedSeat, NormalSeat, RoomySeat, CrampedStandingRoom, _
                NormalStandingRoom, RoomyStandingRoom, CycleSeat, Hammock, _
                Bunk, Cabin, LuxuryCabin, Suite, LuxurySuite, SmallGalley
    
                
                sOutput = sOutput + element.PrintOutput + " "
        
        End Select
    Next
    
    GetAccomodationsOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetAccomodationsOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetEnvironmentalSystemsOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case TotalLifeSystem, ArtificialGravityUnit, EnvironmentalControl, _
                NBCKit, LimitedLifeSystem, FullLifeSystem

                sOutput = sOutput + element.PrintOutput + " "
        End Select
    Next
    
    'append the provisions to it
    For Each element In m_oCurrentVeh.Components
        If element.Datatype = Provisions Then
            sOutput = sOutput + element.PrintOutput + " "
        End If
    Next
    GetEnvironmentalSystemsOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:EnvironmentalSystemsOutput -- Error #" & err.Number & " " & err.Description
End Function

Private Function GetSafetySystemsOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            Case EjectionSeat, CrewEscapeCapsule, Airbag, CrashWeb, _
                WombTank, GravityWeb, GravCompensator
                
                sOutput = sOutput + element.PrintOutput + " "
        End Select
    Next
    GetSafetySystemsOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetSafetySystemsOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Public Function GetPowerSystemsOutPut() As String
    Dim oProfile As clsProfilePower
    Dim oGroup As clsSupplyConsumeGroup
    Dim iGroupCount As Long
    Dim i As Long
    Dim iSupplierCount As Long
    Dim iConsumerCount As Long
    Dim sTemp As String
    Dim j As Long
    
    On Error GoTo err
    For Each oProfile In m_oCurrentVeh.Profiles
        sTemp = sTemp & "Profile " & oProfile.Key & vbNewLine
        iGroupCount = oProfile.groupcount
        If iGroupCount > 0 Then
            For i = 1 To iGroupCount
                Set oGroup = oProfile.Group(i)
                ' get the suppliers
                sTemp = sTemp & " Suppliers " & vbNewLine
                iSupplierCount = oGroup.SupplierCount
                For j = 1 To iSupplierCount
                    sTemp = sTemp & m_oCurrentVeh.Components(oGroup.Supplier(j)).Description
                Next
                ' get the consumers
                sTemp = sTemp & " Consumers " & vbNewLine
                iConsumerCount = oGroup.ConsumerCount
                For j = 1 To iConsumerCount
                    sTemp = sTemp & m_oCurrentVeh.Components(oGroup.consumer(j)).Description
                Next
                sTemp = sTemp
            Next
       End If
        
        GetPowerSystemsOutPut = sTemp
        ' get the keys for each supplier in each group
        
        
        ' get the keys for all consumers attached to each group
        
        
    Next ' next profile
    Set oGroup = Nothing
    Set oProfile = Nothing
    Exit Function
err:
    Debug.Print "modTextOutput:GetPowerSystemsOutput -- Error #" & err.Number & " " & err.Description
    
End Function
'Private Function GetPowerSystemsOutPut() As String
'  MPJ 05/27/02 ENTIRE FUNCTION OBSOLETE and NON FUNCTIONAL WITH NEW POWER SYSTEM PROFILES
'    Dim sOutput As String
'    Dim element As Object
'    Dim sKeyArray() As String
'    Dim i As Long
'    Dim sngPowerRemaining As Single
'
'    'todo: this is simplified because we can get this info directly from the
'    ' m_SC_Groups in each profile
'
'    ' todo: however, we do need seperate writes ups for each profile
'
'    For Each element In m_oCurrentVeh.Components
'        Select Case element.Datatype
'
'            Case MuscleEngine, GasolineEngine, HPGasolineEngine, _
'                TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, _
'                SuperHPGasolineEngine, StandardDieselEngine, _
'                TurboStandardDieselEngine, MarineDieselEngine, _
'                HPDieselEngine, TurboHPDieselEngine, CeramicEngine, _
'                TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, _
'                TurboHPCeramicEngine, SuperHPCeramicEngine, _
'                HydrogenCombustionEngine, EarlySteamEngine, _
'                ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine
'
'                sOutput = sOutput + element.PrintOutput + " "
'
'                '//append the systems that it powers
'                sKeyArray = element.GetCurrentConsumptionSystemKeys
'                If sKeyArray(1) = "" Then
'                Else
'                    sOutput = sOutput + "Powers the "
'                    For i = 1 To UBound(sKeyArray)
'                        sOutput = sOutput + m_oCurrentVeh.Components(sKeyArray(i)).Description & ", "
'                    Next
'                    '//delete the last "," and replace it with "."
'                    sOutput = Left(sOutput, Len(sOutput) - 2)
'                    sngPowerRemaining = element.Output - element.PowerConsumed
'                    If sngPowerRemaining > 0 Then
'                        sOutput = sOutput + " with " & sngPowerRemaining & " kW in reserve."
'                    Else
'                        sOutput = sOutput + ".  "
'                    End If
'                End If
'
'            Case StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, _
'                StandardMHDTurbine, HPMHDTurbine, FuelCell, FissionReactor, _
'                RTGReactor, NPU, FusionReactor, AntimatterReactor, _
'                TotalConversionPowerPlant, CosmicPowerPlant, Soulburner, _
'                ElementalFurnace, ManaEngine, Carnivore, Herbivore, Omnivore, _
'                Vampire, ClockWork, LeadAcidBattery, AdvancedBattery, _
'                Flywheel, RechargeablePowerCell, PowerCell, Snorkel, _
'                ElectricContactPower, LaserBeamedPowerReceiver, _
'                MaserBeamedPowerReceiver, SolarCellArray
'
'                sOutput = sOutput + element.PrintOutput + " "
'
'                '//append the systems that it powers
'                sKeyArray = element.GetCurrentConsumptionSystemKeys
'                If sKeyArray(1) = "" Then
'                Else
'                    sOutput = sOutput + "Powers the "
'                    For i = 1 To UBound(sKeyArray)
'                        sOutput = sOutput + m_oCurrentVeh.Components(sKeyArray(i)).Description & ", "
'                    Next
'                    '//delete the last "," and replace it with "."
'                    sOutput = Left(sOutput, Len(sOutput) - 2)
'                    sngPowerRemaining = element.Output - element.PowerConsumed
'                    If sngPowerRemaining > 0 Then
'                        sOutput = sOutput + " with " & sngPowerRemaining & " kW in reserve."
'                    Else
'                        sOutput = sOutput + ".  "
'                    End If
'                End If
'        End Select
'
'    Next
'
'
'
'    GetPowerSystemsOutPut = sOutput
'End Function

Private Function GetFuelOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    'TODO: Is this where endurance gets spit out?
    '
' 'find the endurance of the engine
'    mvarEndurance = 0 'reset the variable
'    If mvarFuelStorageKeyChain(1) <> "" Then
'        For i = 1 To UBound(mvarFuelStorageKeyChain)
'            mvarEndurance = mvarEndurance + Veh.Components(mvarFuelStorageKeyChain(i)).capacity
'        Next
'        mvarEndurance = mvarEndurance / mvarFuelConsumption
'    Else
'        mvarEndurance = 0
'    End If

    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            Case AntiMatterBay, CoalBunker, WoodBunker, StandardTank, _
                lightTank, UltralightTank, StandardSelfSealingTank, _
                lightSelfSealingTank, UltralightSelfSealingTank
                
                sOutput = sOutput + element.PrintOutput + " "
        End Select
    Next
    GetFuelOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetFuelOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetSurfaceAreaOutput() As String
    Dim sOutput As String
    Dim element As Object
    Dim totalsurfacearea As Single
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            Case Body, Wheel, Skid, Track, Hydrofoil, Hovercraft, _
                Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, _
                Wing, Mast, Superstructure, Turret, Popturret, _
                OpenMount, Gasbag, Pod
                
                totalsurfacearea = totalsurfacearea + element.SurfaceArea
                sOutput = sOutput + element.abbrev + " " + Format(element.SurfaceArea, Settings.FormatString) + ". "
        End Select
    Next
    GetSurfaceAreaOutput = sOutput + "total " + Format(totalsurfacearea, Settings.FormatString) + "."
    Exit Function
err:
    Debug.Print "modTextOutput:GetSurfaceAreaOutput -- Error #" & err.Number & " " & err.Description
End Function

Private Function GetStructureOutput() As String
    Dim sOutput As String
    Dim element As Object
    Dim sBodyStruct As String
    Dim tOutput As String
    
On Error GoTo err
    'get the structure of the body first
    
    With m_oCurrentVeh.Components(BODY_KEY)
        sBodyStruct = element.Description + " - " + .FrameStrength + " frame" + " with " + .Materials + " materials. "
    End With
    
    sOutput = sBodyStruct
    
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            'note Open Mount, Mast and Gasbag are not included here
            Case Body, Wheel, Skid, Track, Hydrofoil, Hovercraft, _
                Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, _
                Wing, Superstructure, Turret, Popturret, _
                Pod
                
                With element
                    tOutput = element.abbrev + " - " + .FrameStrength + " frame" + " with " + .Materials + " materials. "
                End With
                
                'only print this if its different than the Body's structure
                If tOutput <> sBodyStruct Then
                    sOutput = sOutput + " " + tOutput
                End If
        End Select
    Next
    
    GetStructureOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetStructureOutput -- Error #" & err.Number & " " & err.Description
    Resume Next
End Function
Private Function GetHitPointsOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            Case Body, Wheel, Skid, Track, Hydrofoil, Hovercraft, _
                Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, _
                Wing, Mast, Superstructure, Turret, Popturret, _
                OpenMount, Gasbag, Pod, equipmentPod, SolarPanel
        
                sOutput = sOutput + element.abbrev + " " + Format(element.HitPoints) + ", "
        End Select
    Next
    GetHitPointsOutput = Left(sOutput, Len(sOutput) - 2) + "."
    Exit Function
err:
    Debug.Print "modTextOutput:GetHitPointsOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetArmorOutput() As String
    Dim sOutput As String
    Dim element As Object
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            Case ArmorComplexFacing, ArmorBasicFacing, _
                 ArmorOpenFrame, ArmorGunShield, ArmorLocation, _
                 ArmorComponent, ArmorOverall, ArmorWheelGuard
        
                 sOutput = sOutput + m_oCurrentVeh.Components(element.LogicalParent).CustomDescription + " armor: " + element.PrintOutput + vbNewLine
        End Select
    Next
    GetArmorOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetArmorOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetStatisticsOutput() As String
    Dim sTemp As String
    
    On Error GoTo err
    With m_oCurrentVeh.Stats
        sTemp = "Empty weight " + Format(.EmptyWeight, Settings.FormatString) + " lbs., "
        If .UsualInternalPayload <> 0 Then sTemp = sTemp + "Internal payload " + Format(.UsualInternalPayload, Settings.FormatString) + " lbs., "
        sTemp = sTemp + "Loaded weight " + Format(.LoadedWeight, Settings.FormatString) + " lbs., "
        If .SubmergedWeight <> 0 Then sTemp = sTemp + "Submerged weight " + Format(.SubmergedWeight, Settings.FormatString) + " lbs., "
        sTemp = sTemp + "Volume " + Format(.TotalVolume, Settings.FormatString) + " cf. "
        sTemp = sTemp + "Size modifier " + Format(.SizeModifier) + ". "
        sTemp = sTemp + "Cost $" + Format(.TotalPrice, Settings.FormatString) + ". "
        sTemp = sTemp + "HT " + Format(.StructuralHealth)
    End With
    GetStatisticsOutput = sTemp
    Exit Function
err:
    Debug.Print "modTextOutput:GetStatisticsOutput -- Error #" & err.Number & " " & err.Description
    
End Function

Private Function GetSpaceOutput() As String
   'access, empty and cargo space
    Dim sAccessOutput As String
    Dim sEmptyOutput As String
    Dim sCargoOutput As String
    Dim element As Object
    Dim sOutput As String
    
    On Error GoTo err
    
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case Body, Wheel, Skid, Track, Hydrofoil, Hovercraft, _
                Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, _
                Wing, Mast, Superstructure, Turret, Popturret, _
                OpenMount, Gasbag, Pod, SolarPanel, equipmentPod
                
                If element.EmptySpace <> 0 Then
                    sEmptyOutput = sEmptyOutput + element.abbrev + " " + Format(element.EmptySpace, Settings.FormatString) + " cf, "
                End If
                If element.AccessSpace <> 0 Then
                    sAccessOutput = sAccessOutput + element.abbrev + " " + Format(element.AccessSpace, Settings.FormatString) + " cf, "
                End If
            Case Cargo
                sCargoOutput = sCargoOutput + element.PrintOutput + " "
        
        End Select
    Next
    sCargoOutput = Left(sCargoOutput, Len(sCargoOutput) - 1)
    If sCargoOutput <> "" Then
       sOutput = "Space: " + sCargoOutput
    End If
    If sAccessOutput <> "" Then
        sAccessOutput = Left(sAccessOutput, Len(sAccessOutput) - 2)
        sAccessOutput = "(" + sAccessOutput + ")"
        If sOutput = "" Then sOutput = "Space: "
        sOutput = sOutput + " Access space " + sAccessOutput + "."
    End If
    If sEmptyOutput <> "" Then
        sEmptyOutput = Left(sEmptyOutput, Len(sEmptyOutput) - 2)
        sEmptyOutput = "(" + sEmptyOutput + ")"
        If sOutput = "" Then sOutput = "Space: "
        sOutput = sOutput + " Empty space " + sEmptyOutput + "."
    End If
    GetSpaceOutput = sOutput
       'this will error if the component doesnt have Access or Emtpyspace properties.
        'so will just resume past the error
        Exit Function
err:
    Debug.Print "modTextOutput:GetSpaceOutput -- Error #" & err.Number & " " & err.Description
    Resume Next
End Function

Private Function GetStructuralOptionsOutput() As String
    
Dim element As Object
Dim sOutput As String
Dim bControlledInstability As Boolean
Dim bImprovedSuspension As Boolean

On Error GoTo err
    '//get the structural options that are stored in the body
    With m_oCurrentVeh.Options
        If .RollStabilizers Then
        
        End If
        If m_oCurrentVeh.surface.Submersible = True Then
        
        End If
    End With
    '//now get the rest of the structural options from the various other subsassemblies
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            Case Body, Superstructure, Popturret, Turret
                If element.Compartmentalization <> "none" Then
                    'compartmentalization
                    sOutput = sOutput & " " & element.Compartmentalization & " compartmentalization in " & element.abbrev & "."
                End If
            Case Wing
                'folding wings or rotors
                If element.Folding Then
                    sOutput = sOutput & " Folding wings."
                End If
                If element.VariableSweep <> "none" Then
                    sOutput = sOutput & element.VariableSweep & " variable sweep wings."
                End If
                'controlled instability
                If (element.ControlledInstability) And (bControlledInstability = False) Then
                    sOutput = sOutput & " Controlled instability."
                    bControlledInstability = True
                End If
            Case TTRotor, AutogyroRotor, MMRotor, CARotor
                'folding wings or rotors
                If element.Folding Then
                    sOutput = sOutput & " Folding rotors."
                End If
                'controlled instability
                If (element.ControlledInstability) And (bControlledInstability = False) Then
                    sOutput = sOutput & " Controlled instability."
                    bControlledInstability = True
                End If
            Case Track, Skid, Leg
                If (element.ImprovedSuspension) And (bImprovedSuspension = False) Then
                    sOutput = sOutput & " Improved Suspension."
                    bImprovedSuspension = True
                End If
            
            Case Wheel
                If (element.ImprovedSuspension) And (bImprovedSuspension = False) Then
                    sOutput = sOutput & " Improved Suspension."
                    bImprovedSuspension = True
                End If
                If element.ImprovedBrakes Then
                    sOutput = sOutput & " Improved Brakes."
                End If
                If element.AllwheelSteering Then
                    sOutput = sOutput & " All wheel steering."
                End If
                If element.Smartwheels Then
                    sOutput = sOutput & " Smart wheels."
                End If
        End Select
    Next
    
    If sOutput <> "" Then sOutput = "Structural Options: " + sOutput
    
    GetStructuralOptionsOutput = sOutput
    Exit Function
err:
        Debug.Print "modTextOutput:GetStructuralOptionsOutput --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
        Resume Next
End Function

Private Function GetBodyFeatures() As String
Dim sOutput As String
Dim element As Object
Dim sSlope As String
Dim sOutput2 As String

'On Error GoTo err
'    With m_oCurrentVeh.surface
'        If .FloatationHull Then
'            sOutput = sOutput & " Floatation hull " ' todo: fix (rating " & m_oCurrentVeh.stats.FloatationRating & " lbs)."
'        End If
'        If .SubmarineLines Then
'            sOutput = sOutput & " Submarine lines."
'        End If
'        If .HydrodynamicLines <> "none" Then
'            sOutput = sOutput & " " & .HydrodynamicLines & " hydrodynamic lines."
'        End If
'        If .Catamaran Then
'            sOutput = sOutput & " Catamaran."
'        End If
'        If .Trimaran Then
'            sOutput = sOutput & " Trimaran."
'        End If
'        If .StreamLining <> "none" Then
'            sOutput = sOutput & " " & .StreamLining & " streamlining."
'        End If
'    End With
'    '//to get the slope i must check the Body, Superstructure and Turrets and Popturrets
'    For Each element In Vehicle
'        Select Case element.Datatype
'            Case Body, Superstructure, Turret, Popturret
'                sSlope = ""
'                If element.slopef <> "none" Then
'                    If sSlope <> "" Then
'                        sSlope = sSlope & ", front " & element.slopef
'                    Else
'                        sSlope = sSlope & "front " & element.slopef
'                    End If
'                End If
'                If element.slopeb <> "none" Then
'                    If sSlope <> "" Then
'                        sSlope = sSlope & ", back " & element.slopeb
'                    Else
'                        sSlope = sSlope & "back " & element.slopeb
'                    End If
'                End If
'                If element.slopel <> "none" Then
'                    If sSlope <> "" Then
'                        sSlope = sSlope & ", left " & element.slopel
'                    Else
'                        sSlope = sSlope & "left " & element.slopel
'                    End If
'                End If
'                If element.SlopeR <> "none" Then
'                    If sSlope <> "" Then
'                        sSlope = sSlope & ", right " & element.SlopeR
'                    Else
'                        sSlope = sSlope & "right " & element.SlopeR
'                    End If
'                End If
'                If sSlope <> "" Then
'                    sOutput2 = sOutput2 + " Slope on " & element.abbrev & ": " & sSlope & "."
'                End If
'        End Select
'    Next
'
'    sOutput = sOutput + sOutput2
'    If sOutput <> "" Then sOutput = sOutput
'    GetBodyFeatures = sOutput
'    Exit Function
'err:
'    Debug.Print "modTextOutput:GetBodyFeatures --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
End Function

Private Function GetSurfaceFeaturesOutput() As String
Dim sOutput As String
Dim element As Object

On Error GoTo err
With m_oCurrentVeh.surface
    If .Sealed Then sOutput = sOutput + "sealed. "
    
    If .Sealed = False Then
        If .WaterProof Then
            sOutput = sOutput + "waterproofed. "
        End If
    End If
    
    'concealment and stealth
    If .Camouflage Then sOutput = sOutput + "camouflage. "
    If .infraredcloaking <> "none" Then sOutput = sOutput + .infraredcloaking + " infrared cloaking. "
    If .EmissionCloaking <> "none" Then sOutput = sOutput + .EmissionCloaking + " emission cloaking. "
    If .SoundBaffling <> "none" Then sOutput = sOutput + .SoundBaffling + " sound baffling. "
    If .stealth <> "none" Then sOutput = sOutput + .stealth + " stealth. "
    If .LiquidCrystal Then sOutput = sOutput + "liquid crystal skin. "
    If .Chameleon <> "none" Then sOutput = sOutput + .Chameleon + " chameleon system. "
    If .PsiShielding Then sOutput = sOutput + "Psi Shielding. "
    If m_oCurrentVeh.Components(BODY_KEY).liftingbody Then sOutput = sOutput + "Lifting Body. "
    If m_oCurrentVeh.Components(BODY_KEY).FlexibodyOption Then sOutput = sOutput + "Flexibody. "
        
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            Case ForceScreen, DeflectorField, VariableForceScreen
                sOutput = sOutput + element.PrintOutput
            
        End Select
    Next
    If .bMagicLevitation Then sOutput = sOutput + "Magic Levitation. "
    If .bAntigravityCoating Then sOutput = sOutput + "Antigravity Coating. "
    If .bSuperScienceCoating Then sOutput = sOutput + "Super Science Coating. "
End With

If sOutput <> "" Then sOutput = "Surface Features: " + sOutput
GetSurfaceFeaturesOutput = sOutput
Exit Function
err:
    Debug.Print "modTextOutput:GetSufaceFeaturesOutput --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
End Function

Private Function GetOtherSurfaceFeatures() As String
    Dim sOutput As String
    Dim element As Object
    Dim sDefensive As String
    Dim sOutput2 As String
    
    On Error GoTo err
    With m_oCurrentVeh.Options
    'other
        If .Convertible <> "none" Then sOutput = sOutput + .Convertible + ". "
        If .Ram Then sOutput = sOutput + "Ram. "
        If .Bulldozer Then sOutput = sOutput + "Bulldozer. "
        If .Plow Then sOutput = sOutput + "Plow. "
        If .Hitch Then sOutput = sOutput + "Hitch. "
        If .Pin <> "none" Then sOutput = sOutput + .Pin + " pin. "
            
        For Each element In m_oCurrentVeh.Components
            If TypeOf element Is clsWheel Then
                If element.Wheelblades <> "none" Then sOutput = sOutput + element.Wheelblades & " wheelblades. "
                If element.snowtires Then sOutput = sOutput + "Snow tires. "
                If element.racingtires Then sOutput = sOutput + "Racing tires. "
                If element.PunctureResistant Then sOutput = sOutput + "Puncture resistant tires. "
            End If
        Next
    End With
    If sOutput <> "" Then sOutput = "Other Surface Features: " + sOutput
    GetOtherSurfaceFeatures = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetOtherSurfaceFeatures --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
End Function
Private Function GetDefensiveSurfaceFeatures() As String
Dim element As Object
Dim sDefensive As String
Dim sOutput2 As String

    On Error GoTo err
    'defensive surface features found in Armor classes
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case ArmorComplexFacing, ArmorBasicFacing, _
                 ArmorGunShield, ArmorLocation, _
                 ArmorOverall, ArmorWheelGuard
        
                 sDefensive = ""
                If element.rap Then
                    If sDefensive <> "" Then
                        sDefensive = sDefensive & ", reactive armor"
                    Else
                        sDefensive = sDefensive & "reactive armor"
                    End If
                End If
                If element.electrified Then
                    If sDefensive <> "" Then
                        sDefensive = sDefensive & ", electrified"
                    Else
                        sDefensive = sDefensive & "electrified"
                    End If
                End If
                If element.thermal Then
                    If sDefensive <> "" Then
                        sDefensive = sDefensive & ", thermal superconductor armor"
                    Else
                        sDefensive = sDefensive & "thermal superconductor armor"
                    End If
                End If
                If element.radiation Then
                    If sDefensive <> "" Then
                        sDefensive = sDefensive & ", radiation shielding"
                    Else
                        sDefensive = sDefensive & "radiation shielding"
                    End If
                End If
                If element.coating <> "none" Then
                    If sDefensive <> "" Then
                        sDefensive = sDefensive & ", " & element.coating & " coating"
                    Else
                        sDefensive = sDefensive & element.coating & " coating"
                    End If
                End If
                If sDefensive <> "" Then
                    sOutput2 = sOutput2 + " On " & m_oCurrentVeh.Components(element.LogicalParent).Description & ": " & sDefensive & "."
                End If
        End Select
    Next
    
    If sOutput2 <> "" Then sOutput2 = "Defensive Surface Features: " + sOutput2
    GetDefensiveSurfaceFeatures = sOutput2
    Exit Function
err:
    Debug.Print "modTextOutput:GetDefensiveSurfaceFeatures --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
End Function

Private Function GetTopDeckSurfaceFeatures() As String
Dim element As Object
Dim sTopDeck As String
Dim sOutput2 As String
Dim sDeckType As String

    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
        
            Case Body, Superstructure
            If element.TopDeck Then
                 sTopDeck = ""
                If element.covereddeckarea <> 0 Then
                    If sTopDeck <> "" Then
                        sTopDeck = sTopDeck & ", " & Format(element.covereddeckarea, Settings.FormatString) & "sq ft covered"
                    Else
                        sTopDeck = sTopDeck & Format(element.covereddeckarea, Settings.FormatString) & "sq ft covered"
                    End If
                End If
                If element.FlightDeckArea <> 0 Then
                    'get the decktype
                    If element.flightdeckoption = "none" Then
                        sDeckType = "flight deck"
                    Else
                        sDeckType = element.flightdeckoption
                    End If
                    If sTopDeck <> "" Then
                        sTopDeck = sTopDeck & ", " & Format(element.FlightDeckArea, Settings.FormatString) & " sq ft " & sDeckType & " with a length of " & Format(element.flightdecklength, Settings.FormatString) & " ft"
                    Else
                        sTopDeck = sTopDeck & Format(element.FlightDeckArea, Settings.FormatString) & "sq ft " & sDeckType & " with a length of " & Format(element.flightdecklength, Settings.FormatString) & " ft"
                    End If
                End If
                
                If sTopDeck <> "" Then
                    sOutput2 = sOutput2 + " On " & element.Description & ": " & sTopDeck & "."
                End If
        End If
        End Select
    Next
    
    If sOutput2 <> "" Then sOutput2 = "Top Deck: " + sOutput2
    GetTopDeckSurfaceFeatures = sOutput2
    Exit Function
err:
    Debug.Print "modTextOutput:GetTopDeckSurfaceFeatures --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
End Function

Private Function GetWeaponBaysAndHardpoints() As String
    Dim sOutput As String
    Dim sOutput2 As String
    Dim element As Object
    Dim sngLoad As Single
    Dim sngLoad2 As Single
    
    On Error GoTo err
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            Case WeaponBay
                sOutput = sOutput + element.abbrev + " "
                sngLoad = sngLoad + element.loadcapacity
            Case HardPoint
                sOutput2 = sOutput2 + element.abbrev + " "
                sngLoad2 = sngLoad2 + element.loadcapacity
        End Select
    Next
    
    If sOutput <> "" Then
        sOutput = "Weapon bays: " & sOutput & "Total weapon bay load " & Format(sngLoad, Settings.FormatString) & " lbs."
    
    End If
    If sOutput2 <> "" Then
        sOutput2 = "Hardpoints: " & sOutput2 & "Total hardpoint load " & Format(sngLoad2, Settings.FormatString) & " lbs."
    End If
    If (sOutput <> "") And (sOutput2 <> "") Then
        sOutput = sOutput & vbNewLine & vbNewLine & sOutput2
    Else
        sOutput = sOutput & sOutput2
    End If
    GetWeaponBaysAndHardpoints = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetWeaponBaysAndHardpoints --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
End Function

Private Function GetVisionAndDetailsOutput() As String
    'user entered vision and details
    Dim sOutput As String

    On Error GoTo err
    With m_oCurrentVeh.Description
        If .Details <> "" Then
            sOutput = "Details: " + .Details
        End If
        If .Vision <> "" Then
            sOutput = sOutput + vbNewLine + "Vision: " + .Vision
        End If
    End With
    
    GetVisionAndDetailsOutput = sOutput
    Exit Function
err:
    Debug.Print "modTextOutput:GetVisionDetailsOutput --  Error #" & err.Number & " " & err.Description
End Function
Private Function GetPerformanceOutput() As String
Dim element As Object
Dim sOutput As String

On Error GoTo err
For Each element In m_oCurrentVeh.PerformanceProfiles
    If element.Datatype = PERFORMANCEPROFILE Then
    
        With element
            Select Case .PerformanceType
                Case "Air"
                    sOutput = sOutput + .Key + ": "
                    sOutput = sOutput + "Stall Speed " & Format(.aStallSpeed, "standard") & " mph."
                    sOutput = sOutput + " Aerial motive thrust " & Format(.aMotiveThrust, "standard") & " lbs."
                    sOutput = sOutput + " Aerodynamic drag " & .aDrag & "."
                    sOutput = sOutput + " Top speed " & Format(.aTopSpeed, "standard") & " mph."
                    sOutput = sOutput + " aAccel " & Format(.aAcceleration, "standard") & " mph/s."
                    sOutput = sOutput + " aMR " & .aManeuverability & "."
                    sOutput = sOutput + " aSR " & .aStability & "."
                    sOutput = sOutput + " aDecel " & Format(.aDeceleration, "standard") & " mph/s."
                    
                Case "Ground"
                    sOutput = sOutput + .Key + ": "
                    sOutput = sOutput + " Speed " & Format(.gTopSpeed, "standard") & " mph."
                    sOutput = sOutput + " gAccel " & Format(.gAcceleration, "standard") & " mph/s."
                    sOutput = sOutput + " gDecel " & Format(.gDeceleration, "standard") & " mph/s."
                    sOutput = sOutput + " gSR " & .gStability & "."
                    sOutput = sOutput + " gMR " & .gManeuverability & "."
                    sOutput = sOutput + " " & Format(.gPressureDescription, "standard") & " ground pressure."
                    sOutput = sOutput + " Off road speed " & Format(.gOffRoad, "standard") & " mph/s."
                    
                
                Case "Hovercraft"
                    sOutput = sOutput + .Key + ": "
                    sOutput = sOutput + " Hover Altitude " & .hHoverAltitude & " feet."
                    sOutput = sOutput + " Thrust " & Format(.hMotiveThrust, "standard") & " lbs."
                    sOutput = sOutput + " Speed " & Format(.hTopSpeed, "standard") & " mph."
                    sOutput = sOutput + " Drag " & .hDrag & "."
                    sOutput = sOutput + " hAccel " & Format(.hAcceleration, "standard") & " mph/s."
                    sOutput = sOutput + " hSR " & .hstability & "."
                    sOutput = sOutput + " hMR " & .hmaneuverability & " g."
                    sOutput = sOutput + " hDecel " & Format(.hDeceleration, "standard") & " mph/s."
                    
                
                Case "Mag-Lev"
                    sOutput = sOutput + .Key + ": "
                    sOutput = sOutput + " Thrust " & Format(.mlMotiveThrust, "standard") & " lbs."
                    sOutput = sOutput + " Speed " & Format(.mlTopSpeed, "standard") & " mph."
                    sOutput = sOutput + " Stall Speed " & Format(.mlStallSpeed, "standard") & " mph."
                    sOutput = sOutput + " mDrag " & .mlDrag & "."
                    sOutput = sOutput + " mAccel " & Format(.mlAcceleration, "standard") & " mph/s."
                    sOutput = sOutput + " mSR " & .mlStability & "."
                    sOutput = sOutput + " mMR " & .mlManeuverability & "."
                    sOutput = sOutput + " mDecel " & Format(.mlDeceleration, "standard") & " mph/s."
                
                Case "Water"
                    sOutput = sOutput + .Key + ": "
                    sOutput = sOutput + " Hydrodynamic drag " & Format(.wHydroDrag, "standard") & "."
                    sOutput = sOutput + " Aquatic motive thrust " & Format(.wTotalAquaticThrust, "standard") & " lbs."
                    sOutput = sOutput + " Speed " & Format(.wTopSpeed, "standard") & " mph."
                    sOutput = sOutput + " Hydrofoil Speed " & Format(.wHydrofoilSpeed, "standard") & " mph."
                    sOutput = sOutput + " Planing Speed " & Format(.wPlaningSpeed, "standard") & " mph."
                    sOutput = sOutput + " wAccel " & Format(.wAcceleration, "standard") & " mph/s."
                    sOutput = sOutput + " wMR " & .wManeuverability & "."
                    sOutput = sOutput + " wSR  " & .wStability & "."
                    sOutput = sOutput + " wDecel " & Format(.wDeceleration, "standard") & " mph/s."
                    If .wIDeceleration > 0 Then
                        sOutput = sOutput + " Incr wDecel " & Format(.wIDeceleration, "standard") & " mph/s."
                    End If
                    sOutput = sOutput + " wDraft " & Format(.wDraft, "standard") & " feet."
                
                Case "Submerged"
                    sOutput = sOutput + .Key + ": "
                    sOutput = sOutput + "suThrust " & Format(.suTotalAquaticThrust, "standard") & " lbs."
                    sOutput = sOutput + " suDrag " & .suHydroDrag & "."
                    sOutput = sOutput + " suSpeed " & Format(.suTopSpeed, "standard") & " mph."
                    sOutput = sOutput + " suAccel " & Format(.suAcceleration, "standard") & " mph/s."
                    sOutput = sOutput + " suDecel " & Format(.suDeceleration, "standard") & " mph/s."
                    If .suIDeceleration > 0 Then
                        sOutput = sOutput + " Incr suDecel " & Format(.suIDeceleration, "standard") & " mph/s."
                    End If
                    sOutput = sOutput + " suSR " & .suStability & "."
                    sOutput = sOutput + " suMR " & .suManeuverability & "."
                    sOutput = sOutput + " Draft " & Format(.suDraft, "standard") & " feet."
                    If .suCrushDepth = -1 Then
                        sOutput = sOutput & "No Crush Depth"
                    Else
                        sOutput = sOutput + " Crush Depth " & Format(.suCrushDepth, "standard") & " yards."
                     End If
                     
               Case "Space"
                    sOutput = sOutput + .Key + ": "
                    sOutput = sOutput + " Thrust " & Format(.sMotiveThrust, "standard") & " lbs."
                    sOutput = sOutput + " sAccel " & Format(.sAccelerationG, "standard") & " g."
                    sOutput = sOutput + " sAccel " & Format(.sAccelerationMPH, "standard") & " mph/s."
                    sOutput = sOutput + " Turn Around " & Format(.sTurnAroundTime, "standard") & " secs."
                    sOutput = sOutput + " sMR " & Format(.sManeuverability, "standard") & "."
                    sOutput = sOutput + " Hyper " & Format(.sHyperSpeed, "standard") & " parsecs per day."
                    sOutput = sOutput + " Warp " & Format(.sWarpSpeed, "standard") & " parsecs per day."
                    If .sJumpDriveable Then
                        sOutput = sOutput + " Has jump capabilities."
                    End If
                    If .sTeleportationDriveable Then
                        sOutput = sOutput + " Has teleportation drive capabilities."
                    End If
            End Select
            sOutput = sOutput + vbNewLine + vbNewLine
        End With
    End If
Next

GetPerformanceOutput = sOutput
Exit Function
err:
    Debug.Print "modTextOutput:GetPerformanceOutput --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
End Function

Private Function GetDetailedWeaponStats() As String
Dim element As Object
Dim sOutput As String
Dim aHeader, bHeader, cHeader As Boolean
Dim gun() As Variant
Dim i, j, k, l As Long
Dim iLength As Long
Dim iOldLength As Long
Dim iPropID As Long


On Error GoTo err

i = 1
j = 1
k = 1
l = 1
    '//guns and artillery
    bHeader = False
    ReDim gun(1 To 15, 1)
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            Case StoneThrower, BoltThrower, RepeatingBoltThrower, _
                MuzzleLoader, BreechLoader, ManualRepeater, Revolver, _
                MechanicalGatling, SlowAutoloader, FastAutoloader, _
                lightAutomatic, HeavyAutomatic, ElectricGatling
                
                
                If bHeader = False Then
                    'we havent printed our header yet so do it now
                    gun(1, 1) = "Name"
                    gun(2, 1) = "Malf"
                    gun(3, 1) = "Type"
                    gun(4, 1) = "Damage"
                    gun(5, 1) = "SS"
                    gun(6, 1) = "Acc"
                    gun(7, 1) = "1/2D"
                    gun(8, 1) = "Max"
                    gun(9, 1) = "RoF"
                    gun(10, 1) = "Weight"
                    gun(11, 1) = "Cost"
                    gun(12, 1) = "WPS"
                    gun(13, 1) = "VPS"
                    gun(14, 1) = "CPS"
                    gun(15, 1) = "Ldrs."
                    bHeader = True
                End If
                i = i + 1
                ReDim Preserve gun(1 To 15, i)
                gun(1, i) = element.CustomDescription
                gun(2, i) = element.Malfunction
                gun(3, i) = element.TypeDamage1
                gun(4, i) = element.Damage1
                gun(5, i) = element.SnapShot
                gun(6, i) = element.Accuracy
                gun(7, i) = element.halfDamage
                gun(8, i) = element.MaxRange
                gun(9, i) = element.sRoF
                gun(10, i) = element.Weight
                gun(11, i) = element.Cost
                gun(12, i) = element.WPS
                gun(13, i) = element.VPS
                gun(14, i) = element.CPS
                gun(15, i) = element.Loaders
        End Select
    Next
    '//now we must pad each row item with spaces so that they are all the same length
    If gun(1, 1) <> "" Then
        For iPropID = 1 To 15
            iLength = 0
            iOldLength = 0
            For j = 1 To i
                iLength = Maximum(iLength, Len(gun(iPropID, j)))
            Next
            iLength = iLength + 1 '//we need 1 space seperation
            For j = 1 To i
                iOldLength = Len(gun(iPropID, j))
                For k = 1 To iLength - iOldLength
                    gun(iPropID, j) = gun(iPropID, j) & " "
                Next
            Next
        Next
        '//finally we can output it all
        For j = 1 To i
            sOutput = sOutput + sLineBreak
            sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j) & gun(8, j) & gun(9, j) & gun(10, j) & gun(11, j) & gun(12, j) & gun(13, j) & gun(14, j) & gun(15, j)
        Next
        sOutput = sOutput + sLineBreak
        gun(1, 1) = ""
    End If
    '////////////////////////////////////////////////////////////////////
    '//Beam Weapons
    bHeader = False
    i = 1
    ReDim gun(1 To 12, 1)
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            Case BlueGreenLaser, _
                RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, _
                ChargedParticleBeam, NeutralParticleBeam, _
                Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, _
                FusionBeam, GravityBeam, AntiparticleBeam, Graser, _
                Disintegrator, Displacer, BeamedPowerTransmitter, _
                MilitaryParalysisBeam, EnergyDrill
                
                If bHeader = False Then
                    'we havent printed our header yet so do it now
                    gun(1, 1) = "Name"
                    gun(2, 1) = "Malf"
                    gun(3, 1) = "Type"
                    gun(4, 1) = "Damage"
                    gun(5, 1) = "SS"
                    gun(6, 1) = "Acc"
                    gun(7, 1) = "1/2D"
                    gun(8, 1) = "Max"
                    gun(9, 1) = "RoF"
                    gun(10, 1) = "Weight"
                    gun(11, 1) = "Cost"
                    gun(12, 1) = "Power"
                    bHeader = True
                End If
                i = i + 1
                ReDim Preserve gun(1 To 12, i)
                gun(1, i) = element.CustomDescription
                gun(2, i) = element.Malfunction
                gun(3, i) = element.TypeDamage
                gun(4, i) = element.Damage
                gun(5, i) = element.SnapShot
                gun(6, i) = element.Accuracy
                gun(7, i) = element.halfDamage
                gun(8, i) = element.MaxRange
                gun(9, i) = element.rof
                gun(10, i) = element.Weight
                gun(11, i) = element.Cost
                gun(12, i) = element.PowerReqt
        End Select
    Next
    '//now we must pad each row item with spaces so that they are all the same length
    If gun(1, 1) <> "" Then
        For iPropID = 1 To 12
            iLength = 0
            iOldLength = 0
            For j = 1 To i
                iLength = Maximum(iLength, Len(gun(iPropID, j)))
            Next
            iLength = iLength + 1 '//we need 1 space seperation
            For j = 1 To i
                iOldLength = Len(gun(iPropID, j))
                For k = 1 To iLength - iOldLength
                    gun(iPropID, j) = gun(iPropID, j) & " "
                Next
            Next
        Next
        '//finally we can output it all
        For j = 1 To i
            sOutput = sOutput + sLineBreak
            sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j) & gun(8, j) & gun(9, j) & gun(10, j) & gun(11, j) & gun(12, j)
        Next
        sOutput = sOutput + sLineBreak
        gun(1, 1) = ""
    End If
    '////////////////////////////////////////////////////////////////
    '//Bombs, missiles and torps
    bHeader = False
    i = 1
    ReDim gun(1 To 13, 1)
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            Case IronBomb, _
                RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, _
                ProximityMine, PressureTriggerMine, CommandTriggerMine, _
                SmartTriggerMine, UnGuidedMissile, UnGuidedTorpedo, _
                GuidedMissile, GuidedTorpedo
                
                If bHeader = False Then
                    'we havent printed our header yet so do it now
                    gun(1, 1) = "Name"
                    gun(2, 1) = "Malf"
                    gun(3, 1) = "Guid"
                    gun(4, 1) = "Type"
                    gun(5, 1) = "Damage"
                    gun(6, 1) = "Spd"
                    gun(7, 1) = "End"
                    gun(8, 1) = "Max"
                    gun(9, 1) = "Min"
                    gun(10, 1) = "Skill"
                    gun(11, 1) = "WPS"
                    gun(12, 1) = "VPS"
                    gun(13, 1) = "CPS"
                    
                    bHeader = True
                End If
                i = i + 1
                ReDim Preserve gun(1 To 13, i)
                gun(1, i) = element.CustomDescription
                gun(2, i) = element.Malfunction
                gun(3, i) = element.GuidanceSystem
                gun(4, i) = element.TypeDamage1
                gun(5, i) = element.Damage1
                gun(6, i) = element.Speed
                gun(7, i) = element.Endurance
                gun(8, i) = element.MaxRange
                gun(9, i) = element.MinRange
                gun(10, i) = element.Skill
                gun(11, i) = element.Weight / element.Quantity
                gun(12, i) = element.Volume
                gun(13, i) = element.Cost / element.Quantity
        End Select
    Next
    '//now we must pad each row item with spaces so that they are all the same length
    If gun(1, 1) <> "" Then
        For iPropID = 1 To 13
            iLength = 0
            iOldLength = 0
            For j = 1 To i
                iLength = Maximum(iLength, Len(gun(iPropID, j)))
            Next
            iLength = iLength + 1 '//we need 1 space seperation
            For j = 1 To i
                iOldLength = Len(gun(iPropID, j))
                For k = 1 To iLength - iOldLength
                    gun(iPropID, j) = gun(iPropID, j) & " "
                Next
            Next
        Next
        '//finally we can output it all
        For j = 1 To i
            sOutput = sOutput + sLineBreak
            sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j) & gun(8, j) & gun(9, j) & gun(10, j) & gun(11, j) & gun(12, j) & gun(13, j)
        Next
        sOutput = sOutput + sLineBreak
        gun(1, 1) = ""
     End If
     '/////////////////////////////////////////////////////////////
     '//Liquid projectors
     bHeader = False
     i = 1
     ReDim gun(1 To 13, 1)
     For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            Case FlameThrower, WaterCannon
                
                If bHeader = False Then
                    'we havent printed our header yet so do it now
                    gun(1, 1) = "Name"
                    gun(2, 1) = "Malf"
                    gun(3, 1) = "Type"
                    gun(4, 1) = "Damage"
                    gun(5, 1) = "SS"
                    gun(6, 1) = "Acc"
                    gun(7, 1) = "1/2D"
                    gun(8, 1) = "Max"
                    gun(9, 1) = "RoF"
                    gun(10, 1) = "Weight"
                    gun(11, 1) = "Cost"
                    gun(12, 1) = "WPS"
                    gun(13, 1) = "CPS"
                    bHeader = True
                End If
                i = i + 1
                ReDim Preserve gun(1 To 13, i)
                gun(1, i) = element.CustomDescription
                gun(2, i) = element.Malfunction
                gun(3, i) = element.TypeDamage
                gun(4, i) = element.Damage
                gun(5, i) = element.SnapShot
                gun(6, i) = element.Accuracy
                gun(7, i) = element.halfDamage
                gun(8, i) = element.MaxRange
                gun(9, i) = element.rof
                gun(10, i) = element.Weight
                gun(11, i) = element.Cost
                gun(12, i) = element.WPS
                gun(13, i) = element.CPS
        End Select
    Next
    '//now we must pad each row item with spaces so that they are all the same length
    If gun(1, 1) <> "" Then
        For iPropID = 1 To 13
            iLength = 0
            iOldLength = 0
            For j = 1 To i
                iLength = Maximum(iLength, Len(gun(iPropID, j)))
            Next
            iLength = iLength + 1 '//we need 1 space seperation
            For j = 1 To i
                iOldLength = Len(gun(iPropID, j))
                For k = 1 To iLength - iOldLength
                    gun(iPropID, j) = gun(iPropID, j) & " "
                Next
            Next
        Next
        '//finally we can output it all
        For j = 1 To i
            sOutput = sOutput + sLineBreak
            sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j) & gun(8, j) & gun(9, j) & gun(10, j) & gun(11, j) & gun(12, j) & gun(13, j)
        Next
        sOutput = sOutput + sLineBreak
        gun(1, 1) = ""
    End If
    '///////////////////////////////////////////////////////////////
    '//Launchers
    bHeader = False
    i = 1
    ReDim gun(1 To 7, 1)
    For Each element In m_oCurrentVeh.Components
        Select Case element.Datatype
            Case DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, _
                ManualRepeaterLauncher, SlowAutoLoaderLauncher, _
                FastAutoLoaderLauncher, RevolverLauncher, _
                lightAutomaticLauncher, HeavyAutomaticLauncher
            
                If bHeader = False Then
                    'we havent printed our header yet so do it now
                    gun(1, 1) = "Name"
                    gun(2, 1) = "SS"
                    gun(3, 1) = "RoF"
                    gun(4, 1) = "Weight"
                    gun(5, 1) = "Cost"
                    gun(6, 1) = "Ldrs."
                    gun(7, 1) = "Rating"
                    bHeader = True
                End If
                i = i + 1
                ReDim Preserve gun(1 To 7, i)
                gun(1, i) = element.CustomDescription
                gun(2, i) = element.SnapShot
                gun(3, i) = element.rof
                gun(4, i) = element.Weight
                gun(5, i) = element.Cost
                gun(6, i) = element.Loaders
                gun(7, i) = element.MaxLoad
        End Select
    Next
    '//now we must pad each row item with spaces so that they are all the same length
    If gun(1, 1) <> "" Then
        For iPropID = 1 To 7
            iLength = 0
            iOldLength = 0
            For j = 1 To i
                iLength = Maximum(iLength, Len(gun(iPropID, j)))
            Next
            iLength = iLength + 1 '//we need 1 space seperation
            For j = 1 To i
                iOldLength = Len(gun(iPropID, j))
                For k = 1 To iLength - iOldLength
                    gun(iPropID, j) = gun(iPropID, j) & " "
                Next
            Next
        Next
        '//finally we can output it all
        For j = 1 To i
            sOutput = sOutput + sLineBreak
            sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j)
        Next
        sOutput = sOutput + sLineBreak
        gun(1, 1) = ""
    End If
    '//add a new line and send then return the entire output value
    sOutput = sOutput + vbNewLine
    GetDetailedWeaponStats = sOutput
   Exit Function
err:
    Debug.Print "modTextOutput:GetDetailedWeaponStats --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
End Function

Private Function NumericToString(ByVal nNumber As Variant) As String
'//this function accepts a number and if that number is between
'1 and 10 it will convert them to "One" and "Ten" for instance. If
'the number is greater than 10 it will just return the number formatted
'as a string
Dim retval As String

'NOTE: This is currently only set up to handle longs and not decimals
If nNumber >= 1 And nNumber <= 10 Then
    If nNumber = 1 Then
        retval = "One"
        'retval = "" 'if its a 1 we'll jsut leave it blank since its assumed to be 1 unless noted
    ElseIf nNumber = 2 Then
        retval = "Two"
    ElseIf nNumber = 3 Then
        retval = "Three"
    ElseIf nNumber = 4 Then
        retval = "Four"
    ElseIf nNumber = 5 Then
        retval = "Five"
    ElseIf nNumber = 6 Then
        retval = "Six"
    ElseIf nNumber = 7 Then
        retval = "Seven"
    ElseIf nNumber = 8 Then
        retval = "Eight"
    ElseIf nNumber = 9 Then
        retval = "Nine"
    ElseIf nNumber = 10 Then
        retval = "Ten"
    End If
Else
    retval = "(" + Format(nNumber) + ")"
End If

NumericToString = retval

End Function

Private Function RemoveParenthetical(ByVal strIn As String)
    'JAW 2000.060.26
    Dim varSplit As Variant
    Dim i As Integer
    Dim strTemp As String
    Dim strLocation As String
    
    varSplit = Split(strIn, "(")
    For i = 1 To UBound(varSplit)
        strTemp = varSplit(i - 1)
        If InStr(1, strTemp, ")") Then
'                If Left(strTemp, 1) = "$" Then
'                strLocation = ""
'            Else
'                strLocation = Left(strTemp, 2)
'            End If
'        varSplit(i - 1) = "[" & strLocation & "]" & Split(strTemp, ")")(1)
            varSplit(i - 1) = Split(strTemp, ")")(1)
        End If
    Next i
    For i = 1 To UBound(varSplit)
        RemoveParenthetical = RemoveParenthetical & varSplit(i - 1)
    Next i

End Function
