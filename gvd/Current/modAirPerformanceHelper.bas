Attribute VB_Name = "modAirPerformanceHelper"
Option Explicit

Function CalcAMotiveThrust(ByRef KeyChain As Variant, _
                        ByRef sngReservedVectoredThrust As Single, _
                        ByVal bTreatTiltRotorsAsPropellers As Boolean, _
                        ByVal sngTiltRotorForwardThrust As Single, _
                        ByVal sngPercentThrust As Single, _
                        ByVal bAfterBurnersOn As Boolean, _
                        ByVal lngPerformanceType As Long, _
                        ByVal sngTotalRamJetThrust As Single, _
                        ByVal sngTotalRamJetThrustAB As Single, _
                        ByVal sngTotalTurboRamJetThrust As Single, _
                        ByVal sngTotalTurboramJetThrustAB As Single) As Single
'JAW 2000.05.28
'modified this sub to exclude vectored thrust directly, since that is accounted for
'by the CalcExcessVecThrust sub added in at the end.
'JAW 2000.06.13
'reformed sub to use only vectored thrust not allocated to lift
On Error Resume Next
Dim i As Integer
Dim TempThrust As Single
Dim dType As String
Dim sKey As String
Dim VectoredThrust As Single
Dim bVTStandIn As Boolean

    If KeyChain(1) = "" Then Exit Function 'exit if there are no propulsion systems in the keychain
    
    'cycle through each item in the array and total the Motive Thrust
    For i = 1 To UBound(KeyChain)
        sKey = KeyChain(i)
        dType = Veh.Components(sKey).Datatype
        
        'JAW 2000.06.05
        bVTStandIn = False
On Error Resume Next
'On Error GoTo ErrHandler
        bVTStandIn = Veh.Components(sKey).VectoredThrust
ErrHandler:
        'do nothing
'On Error Resume Next
        TempThrust = 0
        Select Case dType
            Case CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain
                If bTreatTiltRotorsAsPropellers Then
                    If Veh.Components(sKey).TiltRotor Then
                        '//add in the title rotor forward thrust
                        TempThrust = sngTiltRotorForwardThrust * sngPercentThrust
                    Else
                        TempThrust = Veh.Components(sKey).MotiveThrust * sngPercentThrust
                    End If
                Else
                    TempThrust = Veh.Components(sKey).MotiveThrust * sngPercentThrust
                End If
            Case OrnithopterDrivetrain, _
                 AerialPropeller, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, _
                 WhiffletreeHarness, AerialSail, AerialSailForeAftRig
                
                TempThrust = Veh.Components(sKey).MotiveThrust * sngPercentThrust
            Case Turbojet, Turbofan, Hyperfan, FusionAirRam
                'use Afterburner Thrust if this option is enabled by the user
                If Veh.Components(sKey).LiftEngine Then
                    'dont add lift engine thrust
                ElseIf bAfterBurnersOn Then
                    'determine if this engine has afterburners or not
                    If Veh.Components(sKey).Afterburner Then
                        TempThrust = Veh.Components(sKey).ABThrust 'PPP SHould not be using AB unless user has it checked
                    Else
                        TempThrust = Veh.Components(sKey).MotiveThrust * sngPercentThrust
                    End If
                Else 'use normal engine thrust without afterburner
                    TempThrust = TempThrust + Veh.Components(sKey).MotiveThrust * sngPercentThrust
                End If
            Case Ramjet
                'store these values for use later in the TopSpeed calculations since
                'Ramjets only work if Topspeed is at least 375mph
                sngTotalRamJetThrust = sngTotalRamJetThrust + Veh.Components(sKey).MotiveThrust * sngPercentThrust
                sngTotalRamJetThrustAB = sngTotalRamJetThrustAB + Veh.Components(sKey).ABThrust
                
            Case TurboRamjet
                'store these values for use later since they add .2 x their thrust
                'if the speed is greater than 375mph
                sngTotalTurboRamJetThrust = sngTotalTurboRamJetThrust + Veh.Components(sKey).MotiveThrust * sngPercentThrust
                sngTotalTurboramJetThrustAB = sngTotalTurboramJetThrustAB + Veh.Components(sKey).ABThrust
                
                'but in the meantime, they just add their default thrust
                'use Afterburner Thrust if this option is enabled by the user
                If Veh.Components(sKey).LiftEngine Then
                
                ElseIf bAfterBurnersOn Then
                    'determine if this engine has afterburners or not
                    If Veh.Components(sKey).Afterburner Then
                        TempThrust = Veh.Components(sKey).ABThrust
                    Else
                        TempThrust = Veh.Components(sKey).MotiveThrust * sngPercentThrust
                    End If
                Else 'use normal engine thrust without afterburner
                    TempThrust = Veh.Components(sKey).MotiveThrust * sngPercentThrust
                End If
            
            Case DuctedFan '//dont add lift engine OR hoverfan lift
                If Veh.Components(sKey).LiftEngine Then
                'if engine uses vectored thrust we must total the thrust
                'seperatetly.  If its determined that the vehicle does not have
                'wings, Rotors or a lifting body, then we must omit
                ' the vectored thrust required to lift the vehicle off the ground
                ElseIf lngPerformanceType = PERFORMANCEHOVER Then '//hovercrafts we already calculate the amount of vectored thrust we need to remove in the CalcHHoveraltitude function
                    TempThrust = Veh.Components(sKey).MotiveThrust * sngPercentThrust
                Else
                    TempThrust = Veh.Components(sKey).MotiveThrust * sngPercentThrust
                End If
                
            Case LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                 OptimizedFusion, AntimatterThermal, AntimatterPion, _
                 StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine
                 'do not included thrust if they are lift engines
                If Veh.Components(sKey).LiftEngine Then
                'if engine uses vectored thrust we must total the thrust
                'seperatetly.  If its determined that the vehicle does not have
                'wings, Rotors or a lifting body, then we must omit
                ' the vectored thrust required to lift the vehicle off the ground
                ElseIf lngPerformanceType = PERFORMANCEHOVER Then '//hovercrafts we already calculate the amount of vectored thrust we need to remove in the CalcHHoveraltitude function
                    TempThrust = Veh.Components(sKey).MotiveThrust * sngPercentThrust
                Else
                    TempThrust = Veh.Components(sKey).MotiveThrust * sngPercentThrust
                End If
        End Select

        If bVTStandIn Then
            CalcAMotiveThrust = CalcAMotiveThrust + TempThrust '* ThrustContribution(sKey))
        Else
            CalcAMotiveThrust = CalcAMotiveThrust + TempThrust
        End If
        
    Next
    
    '//now determine if we have wings, rotors or lifting body which will allow us to use all our our thrust
    '//else we cant use the reserved thrust
    If (VehicleHasWings) Or (VehicleHasRotors) Or (Veh.Components(BODY_KEY).LiftingBody) Then
        sngReservedVectoredThrust = 0
    End If
        
    'subtract the amount of vectored thrust we are using
    CalcAMotiveThrust = CalcAMotiveThrust - sngReservedVectoredThrust
    If CalcAMotiveThrust < 0 Then CalcAMotiveThrust = 0
    
End Function

Function GetUseableVectoredThrust(ByVal sngThrustNeeded As Single, _
                                  ByRef KeyChain As Variant, _
                                  ByVal sngPercentThrust As Single) As Single
    
    Dim sngTemp As Single
    Dim sKey As String
    Dim dType As Long
    Dim i As Long
    
    If KeyChain(1) = "" Then Exit Function 'exit if there are no propulsion systems in the keychain
    
    'cycle through each item in the array and total the Motive Thrust
    For i = 1 To UBound(KeyChain)
        sKey = KeyChain(i)
        dType = Veh.Components(sKey).Datatype
        Select Case dType
            Case DuctedFan, _
                LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                OptimizedFusion, AntimatterThermal, AntimatterPion, _
                StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, _
                TurboRamjet, Ramjet, Turbojet, Turbofan, Hyperfan, FusionAirRam
                If Veh.Components(sKey).VectoredThrust Then
                    sngTemp = sngTemp + Veh.Components(sKey).MotiveThrust * sngPercentThrust
                    If sngTemp > sngThrustNeeded Then
                        sngTemp = sngThrustNeeded
                        Exit For
                    End If
                End If
        End Select
    Next
    GetUseableVectoredThrust = sngTemp

End Function

        
Function CalcADrag(ByVal bPopTurretsExtended As Boolean, _
                   ByVal bWheelsSkidsExtended As Boolean, _
                   ByVal bResponsive As Boolean) As Single
                   
Dim element As Object
Dim dType As Integer
Dim Sa As Single 'surface area of entire vehicle
Dim R As Single 'surface area of anything retractable
Dim Sl As Long 'streamlining modifier
Dim D As Long
Dim sStreamlining As String

'get Sa
Sa = Veh.Stats.totalSurfaceArea

'get surface area of anything retractable on the vehicle
For Each element In Veh.Components
    If TypeOf element Is clsPopTurret Then
        If bPopTurretsExtended = False Then
            R = R + element.SurfaceArea
        End If
    ElseIf TypeOf element Is clsWheel Then
        If element.RetractLocation <> "none" Then
            If bWheelsSkidsExtended = False Then
                R = R + element.SurfaceArea
            End If
        End If
    ElseIf TypeOf element Is clsSkid Then
        If element.RetractLocation <> "none" Then
            If bWheelsSkidsExtended = False Then
                R = R + element.SurfaceArea
            End If
        End If
    End If
Next
    
'get streamlining modifier
With Veh.Surface
    sStreamlining = .StreamLining
    If sStreamlining = "none" Then
        Sl = 1
    ElseIf sStreamlining = "fair" Then
        Sl = 2
    ElseIf sStreamlining = "good" Then
        Sl = 3
    ElseIf sStreamlining = "very good" Then
        Sl = 5
    ElseIf sStreamlining = "superior" Then
        Sl = 10
    ElseIf sStreamlining = "excellent" Then
        Sl = 20
    ElseIf sStreamlining = "radical" Then
        Sl = 40
    End If
    'add 20% for responsive structure
    If bResponsive Then Sl = Sl * 1.2
End With

'get D
'add 5 for each loaded hardpoint but only if calcuating with hardpoints loaded
D = D + (5 * Veh.Stats.TotalHardPointConnections)
    
'add total surface area of any vehicle on the deck or on external cradles

'add 20 if the vehicle is worn as a harness.
For Each element In Veh.Components
    If TypeOf element Is clsCrewStation Then
        If element.Datatype = HarnessCrewStation Then
            D = D + 20
            Exit For
        End If
    End If
Next

For Each element In Veh.Components
    dType = element.Datatype
    
    If (TypeOf element Is clsAccommodation) Or (TypeOf element Is clsCrewStation) Then
        'check for exposed seat or standingroom
        If dType = CycleSeat Then
            D = D + (15 * element.Quantity)
        ElseIf element.Exposed Then
            D = D + (10 * element.Quantity)
        End If
    ElseIf dType = HardPoint Then ' add 5 for each loaded hardpoint
        D = D + (5 * element.Quantity)
    End If
Next

'compute drag
CalcADrag = ((Sa - R) / Sl) + D

End Function


Function CalcAAcceleration(ByVal Thrust As Single, ByVal Weight As Double) As Single
Dim TempAcceleration As Single


If Weight = 0 Then 'check for divide by zero
    TempAcceleration = 0
Else ' do calculation
    TempAcceleration = (Thrust / Weight) * 20
End If

CalcAAcceleration = Round(TempAcceleration, 0)
End Function

Function CalcAManeuverability(ByVal sngStallSpeed As Single, ByVal bResponsive As Boolean, ByVal Weight As Double) As Single
Dim TempM As Single
Dim TempM2 As Single

'this is the max safe G's the vehicle can pull in flight
If sngStallSpeed = 0 Then
    If (bResponsive) And (VehiclehasElectORCompcontrols) Then
        TempM = (gVehicleTL + 2 - Veh.Stats.SizeModifier) / 2
    ElseIf (bResponsive) Or (VehiclehasElectORCompcontrols) Then
        TempM = (gVehicleTL + 1 - Veh.Stats.SizeModifier) / 2
    Else
        TempM = (gVehicleTL - Veh.Stats.SizeModifier) / 2
    End If
    If TempM <= 0 Then TempM = 0.125
End If

If (Veh.Components(BODY_KEY).LiftingBody) And (VehicleHasWings = False) Then
    If VehiclehasElectORCompcontrols Then
        TempM2 = 0.25
    Else
        TempM2 = 0.125
    End If
    If bResponsive Then TempM = TempM * 2
ElseIf (VehicleHasWings) Or (VehicleHasRotors) Then
    Dim Whp As Long 'wing hit points
    Dim Rhp As Long 'Rotor hit points
    Dim Lwt As Single
    Dim RotorTL As Integer
    Dim WingTL As Integer
    Dim TempTL As Integer
    
    Whp = GetTotalHitPoints("clsWing")
    Rhp = GetTotalHitPoints("clsRotor")
    Lwt = Weight
    WingTL = GetLowestTL("clsWing")
    RotorTL = GetLowestTL("clsRotor")
    TempTL = MinimumNonZero(WingTL, RotorTL)
    
    'tech level modifiers
    If bResponsive Then TempTL = TempTL + 1
    If AllWingsAreHighAgility Then TempTL = TempTL + 1
    If AllWingsAreVariableSweep Then TempTL = TempTL + 1
    If VehicleHasCompControls Then TempTL = TempTL + 1
    If AllWingsRotorsControlledInstability Then TempTL = TempTL + 2
    If VehicleHasMMRRotors Then TempTL = TempTL - 1
    
    'check for divide by zero
    If Lwt = 0 Then
        TempM2 = 0
    Else
        TempM2 = ((Whp + Rhp) / Lwt) * TempTL * 30
    End If
End If

'use the higher result
TempM = Maximum(TempM, TempM2)

'round to nearest whole or .5
TempM = TempM / 0.5

TempM = Fix(TempM) * 0.5

If TempM = 0 Then TempM = 0.5
CalcAManeuverability = TempM

End Function

Function CalcAStability() As Single
Dim TempSR

'first get the initial SR based on vehicle volume
Select Case Veh.Stats.TotalVolume
    Case Is <= 99
        TempSR = 2
    Case Is <= 999
        TempSR = 3
    Case Is <= 9999
        TempSR = 4
    Case Is <= 99999
        TempSR = 5
    Case Is >= 100000
        TempSR = 6
End Select
'//maneuver controls modifier
If VehiclehasElectORCompcontrols Then TempSR = TempSR + 1
'//tech level modifier
If gVehicleTL <= 6 Then
    TempSR = TempSR - 1
ElseIf gVehicleTL >= 8 Then
    TempSR = TempSR + 1
End If

If VehicleHasCoaxialRotors Then '//vehicles with coax rotors can ignore the NoWing/StubWing modifier
    
ElseIf (VehicleHasWings = False) Or (VehicleHasOnlyStubWings) Then
   TempSR = TempSR - 1
End If
If Veh.Components(BODY_KEY).LiftingBody Then TempSR = TempSR - 1
If VehicleHasBipeorTripWings Then TempSR = TempSR - 1
If AllWingsRotorsControlledInstability Then TempSR = TempSR - 1
If (Veh.Surface.Stealth = "radical") And (VehicleHasWings) Then TempSR = TempSR - 1

If TempSR < 1 Then TempSR = 1

CalcAStability = TempSR

End Function


Function CalcATopSpeed(ByVal Drag As Single, _
                       ByVal Thrust As Single, _
                       ByVal bAfterBurnersOn As Boolean, _
                       ByVal sngTotalRamJetThrust As Single, _
                       ByVal sngTotalTurboRamJetThrust As Single, _
                       ByVal sngTotalRamJetThrustAB As Single, _
                       ByVal sngTotalTurboramJetThrustAB As Single) As Single
                       
Dim TempSpeed As Single

If Drag = 0 Or Thrust < 0 Then 'check for divide by zero or sqrt of a negative
    CalcATopSpeed = 0
    Exit Function
End If

TempSpeed = Sqr(7500 * (Thrust / Drag)) 'do the main calculation

'adjust speed if using Ramjets or TurboRamJets
If TempSpeed >= 375 Then
    If bAfterBurnersOn Then
        Thrust = Thrust + sngTotalRamJetThrustAB
        Thrust = Thrust + (sngTotalTurboramJetThrustAB * 0.2)
    Else
        Thrust = Thrust + sngTotalRamJetThrust
        Thrust = Thrust + (sngTotalTurboRamJetThrust * 0.2)
    End If
    TempSpeed = Sqr(7500 * (Thrust / Drag)) 'do the updated calculation
End If


CalcATopSpeed = TempSpeed
End Function

Function CalcAMaxSpeed(ByVal TempSpeed As Single, _
                       ByRef KeyChain As Variant, _
                       ByVal bTreatTiltRotorsAsPropellers As Boolean, _
                       ByRef sDesignCheck As String) As Single
Dim Sl As String
Dim sDC As String ' design check
Dim bSpeed As Double
Dim bOriginal As Double

'reduce vehicle's top speed to its restricted maximum speed

' Check to make sure we dont exceed speed of slowest animal(if applicable)
bOriginal = TempSpeed
bSpeed = TempSpeed
TempSpeed = MinimumNonZero(TempSpeed, GetSlowestAnimalSpeed(KeyChain))
If bSpeed <> TempSpeed Then
    sDC = "your speed is limited by a harnessed animal's top speed."
    bSpeed = TempSpeed
End If
    
'DR
If GetLowestDR < 5 Then
    TempSpeed = Minimum(TempSpeed, 600)
    If bSpeed <> TempSpeed Then
        sDC = "all armor DR must be greater than 5 to exceed 600mph."
        bSpeed = TempSpeed
    End If
ElseIf GetLowestDR < 20 Then
    TempSpeed = Minimum(TempSpeed, 2000)
    If bSpeed <> TempSpeed Then
        sDC = "all armor DR must be greater than 20 to exceed 2000mph."
        bSpeed = TempSpeed
    End If
End If

'streamlining
Sl = Veh.Surface.StreamLining
If (Sl = "none") Or (Sl = "fair") Or (Sl = "good") Then
    TempSpeed = Minimum(TempSpeed, 600)
    If bSpeed <> TempSpeed Then
        sDC = "you need better than good streamlining to exceed 600mph."
        bSpeed = TempSpeed
    End If
ElseIf Sl = "very good" Then
    If Not Veh.Components(BODY_KEY).LiftingBody Then
        TempSpeed = Minimum(TempSpeed, 740)
        If bSpeed <> TempSpeed Then
            sDC = "you need better than Very good streamlining (or very good + lifting body) to exceed 740mph."
            bSpeed = TempSpeed
        End If
    End If
End If

'aerial propellers
If PropulsionSystemExistsOnKeychain(AerialPropeller, KeyChain) Then
    TempSpeed = Minimum(TempSpeed, 600)
    If bSpeed <> TempSpeed Then
        sDC = "your aerial Propellers limit top speed to 600mph."
        bSpeed = TempSpeed
    End If
End If

'flarecraft wings
If VehicleHasWings Then
    If VehicleHasFlarecraftWings Then
        TempSpeed = Minimum(TempSpeed, 400)
        If bSpeed <> TempSpeed Then
            sDC = "Your flarecraft wings are limiting top speed to 400mph."
            bSpeed = TempSpeed
        End If
    End If
End If

'Rotors
If (bTreatTiltRotorsAsPropellers) And (VehicleHasRotors) And (VehicleHasNonTiltRotors = False) Then
    TempSpeed = Minimum(TempSpeed, 600)
    If bSpeed <> TempSpeed Then
        sDC = "Your tilt rotors acting as propellers are limiting top speed to 600mph." & vbNewLine
        bSpeed = TempSpeed
    End If
Else
    If VehicleHasRotors Then
        If GetLowestTL("clsRotor") <= 6 Then
            TempSpeed = Minimum(TempSpeed, 150)
            If bSpeed <> TempSpeed Then
                sDC = "Your TL6- rotors are limiting top speed to 150mph." & vbNewLine
                bSpeed = TempSpeed
            End If
        Else
            TempSpeed = Minimum(TempSpeed, 300)
            If bSpeed <> TempSpeed Then
                sDC = "Your TL7+ rotors are limiting top speed to 300mph." & vbNewLine
                bSpeed = TempSpeed
            End If
        End If
    End If
End If

'aerial sails
If (PropulsionSystemExistsOnKeychain(AerialSail, KeyChain)) Then
    TempSpeed = Minimum(TempSpeed, 100)
    If bSpeed <> TempSpeed Then
        sDC = "Your aerial sail is  limiting your top speed to 100mph." & vbNewLine
        bSpeed = TempSpeed
    End If
ElseIf PropulsionSystemExistsOnKeychain(AerialSailForeAftRig, KeyChain) Then 'uses keychain propulsion systems only
    TempSpeed = Minimum(TempSpeed, 100)
    If bSpeed <> TempSpeed Then
        sDC = "Your aerial sail is limiting your top speed to 100mph." & vbNewLine
    End If
End If

If bOriginal = TempSpeed Then
    sDesignCheck = ""
Else
    sDesignCheck = sDC
End If

CalcAMaxSpeed = TempSpeed
End Function

Function PropulsionSystemExistsOnKeychain(ByVal Datatype As Integer, _
                                          ByRef KeyChain As Variant) As Boolean
Dim i As Long
Dim sKey As String
Dim dType As Integer

If KeyChain(1) = "" Then Exit Function 'exit if there are no propulsion systems in the keychain

For i = 1 To UBound(KeyChain)
sKey = KeyChain(i)
dType = Veh.Components(sKey).Datatype

    If dType = Datatype Then
        PropulsionSystemExistsOnKeychain = True
        Exit Function
    End If
Next

PropulsionSystemExistsOnKeychain = False
End Function


