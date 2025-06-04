VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerformanceWater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvartotalramjetthrust As Single
Private mvarTotalTurboRamJetThrust As Single
Private mvarTotalRamJetThrustAB As Single
Private mvarTotalTurboramJetThrustAB As Single

Private mvarwAcceleration As Single
Private mvarwDeceleration As Single
Private mvarwIDeceleration As Single 'the increased deceleration
Private mvarwDraft As Single
Private mvarwHydroDrag As Single
Private mvarwManeuverability As Single
Private mvarwStability As Single
Private mvarwTopSpeed As Single
Private mvarwPlaningSpeed As Single 'note: planing speed is in addition to topspeed (see page 131 top right column)
Private mvarwHydrofoilSpeed As Single 'note:hydrofoil speed is in addition to topspeed(see page 131 top right column
Private mvarwTotalAquaticThrust As Single

Private mvarKey As String
Private mvarParent As String
Private mvarDatatype As Long
Private mvarDescription As String

Private mbResponsive As Boolean
Private mvarTiltRotorForwardThrust As Single
Private mvarTreatTiltRotorsAsPropellers As Boolean
Private mvarAfterBurnersOn As Boolean
Private mvarHardPointsOn As Boolean
Private mvarWheelsSkidsExtended As Boolean
Private mvarPopTurretsExtended As Boolean
        
Private mvarPercentThrust As Single
Private mvarPercentCrewWeight As Single
Private mvarPercentFuelWeight As Single
Private mvarPercentCargoWeight As Single
Private mvarPercentHardpointWeight As Single
Private mvarPercentProvisionWeight As Single
Private mvarPercentAmmunitionWeight As Single
Private mvarPercentAuxVehicleWeight As Single
Private m_VWeight As Double
Private m_VMass As Double
Private g_sDC As String
Private mvarZZInit As Byte

Option Base 1 'make sure the Keychain array starts at 1 and not 0

Private mvarKeyChain As Variant
'JAW 2000.06.12
'mvarKeyChain is an array of PropulsionKeys representing engine components used in a performance profile.

'JAW 2000.06.15
'Short message to be displayed if something is wrong with mode of transportation.
Private mvarAdvisory As String

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

Public Property Let Datatype(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataType = 5
    mvarDatatype = vdata
End Property


Public Property Get Datatype() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataType
    Datatype = mvarDatatype
End Property


Public Property Let PercentThrust(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wTopSpeed = 5
    If vdata > 100 Then vdata = 100
    If vdata < 0 Then vdata = 0
    mvarPercentThrust = vdata / 100
    
End Property


Public Property Get PercentThrust() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wTopSpeed
    PercentThrust = mvarPercentThrust * 100
End Property

Public Property Let PercentCrewWeight(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wTopSpeed = 5
    If vdata > 100 Then vdata = 100
    If vdata < 0 Then vdata = 0
    mvarPercentCrewWeight = vdata / 100
End Property


Public Property Get PercentCrewWeight() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wTopSpeed
    PercentCrewWeight = mvarPercentCrewWeight * 100
End Property
Public Property Let PercentFuelWeight(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wTopSpeed = 5
    If vdata > 100 Then vdata = 100
    If vdata < 0 Then vdata = 0
    mvarPercentFuelWeight = vdata / 100
End Property


Public Property Get PercentFuelWeight() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wTopSpeed
    PercentFuelWeight = mvarPercentFuelWeight * 100
End Property
Public Property Let PercentCargoWeight(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wTopSpeed = 5
    If vdata > 100 Then vdata = 100
    If vdata < 0 Then vdata = 0
    mvarPercentCargoWeight = vdata / 100
End Property


Public Property Get PercentCargoWeight() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wTopSpeed
    PercentCargoWeight = mvarPercentCargoWeight * 100
End Property
Public Property Let PercentHardpointWeight(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wTopSpeed = 5
    If vdata > 100 Then vdata = 100
    If vdata < 0 Then vdata = 0
    mvarPercentHardpointWeight = vdata / 100
End Property


Public Property Get PercentHardpointWeight() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wTopSpeed
    PercentHardpointWeight = mvarPercentHardpointWeight * 100
End Property
Public Property Let PercentProvisionWeight(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wTopSpeed = 5
    If vdata > 100 Then vdata = 100
    If vdata < 0 Then vdata = 0
    mvarPercentProvisionWeight = vdata / 100
End Property


Public Property Get PercentProvisionWeight() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wTopSpeed
    PercentProvisionWeight = mvarPercentProvisionWeight * 100
End Property
Public Property Let PercentAmmunitionWeight(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wTopSpeed = 5
    If vdata > 100 Then vdata = 100
    If vdata < 0 Then vdata = 0
    mvarPercentAmmunitionWeight = vdata / 100
End Property


Public Property Get PercentAmmunitionWeight() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wTopSpeed
    PercentAmmunitionWeight = mvarPercentAmmunitionWeight * 100
End Property


Public Property Let PercentAuxVehicleWeight(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PercentAuxVehicleWeight = 5
    If vdata > 100 Then vdata = 100
    If vdata < 0 Then vdata = 0
    mvarPercentAuxVehicleWeight = vdata / 100
End Property


Public Property Get PercentAuxVehicleWeight() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PercentAuxVehicleWeight
    PercentAuxVehicleWeight = mvarPercentAuxVehicleWeight * 100
End Property

Public Property Let KeyChain(ByVal vdata As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.KeyChain = 5
    mvarKeyChain = vdata
End Property

Public Property Get KeyChain() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.KeyChain
    KeyChain = mvarKeyChain
End Property

Public Property Let Key(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vdata
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
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



Public Function GetCurrentKeys() As String()
GetCurrentKeys = VariantArrayToStringArray(mvarKeyChain)
End Function

Public Sub AddKey(PropulsionKey As String)
    mvarKeyChain = mAddKey(mvarKeyChain, PropulsionKey)
End Sub

Public Sub RemoveKey(PropulsionKey As String)
    mvarKeyChain = mRemoveKey(mvarKeyChain, PropulsionKey)
End Sub

        
Public Property Let Advisory(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Advisory = "5"
    mvarAdvisory = vdata
End Property


Public Property Get Advisory() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Advisory
    Advisory = mvarAdvisory
End Property

Public Property Let HardPointsOn(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HardPointsOn = 5
    mvarHardPointsOn = vdata
End Property


Public Property Get HardPointsOn() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HardPointsOn
    HardPointsOn = mvarHardPointsOn
End Property



Public Property Let WheelsSkidsExtended(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WheelsSkidsExtended = 5
    mvarWheelsSkidsExtended = vdata
End Property


Public Property Get WheelsSkidsExtended() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WheelsSkidsExtended
    WheelsSkidsExtended = mvarWheelsSkidsExtended
End Property

Public Property Let PopTurretsExtended(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PopTurretsExtended = 5
    mvarPopTurretsExtended = vdata
End Property


Public Property Get PopTurretsExtended() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HardPointsOn
    PopTurretsExtended = mvarPopTurretsExtended
End Property

Public Property Let AfterburnersOn(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AfterburnersOn = 5
    mvarAfterBurnersOn = vdata
End Property

Public Property Get AfterburnersOn() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AfterburnersOn
    AfterburnersOn = mvarAfterBurnersOn
End Property


Public Property Let TreatTiltRotorsAsPropellers(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AfterburnersOn = 5
    mvarTreatTiltRotorsAsPropellers = vdata
End Property

Public Property Get TreatTiltRotorsAsPropellers() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AfterburnersOn
    TreatTiltRotorsAsPropellers = mvarTreatTiltRotorsAsPropellers
End Property

Public Property Let TotalRamJetThrust(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalRamJetThrust = 5
    mvartotalramjetthrust = vdata
End Property

Public Property Get TotalRamJetThrust() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalRamJetThrust
    TotalRamJetThrust = mvartotalramjetthrust
End Property

Public Property Let TotalTurboRamJetThrust(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalTurboRamJetThrust = 5
    mvarTotalTurboRamJetThrust = vdata
End Property

Public Property Get TotalTurboRamJetThrust() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalTurboRamJetThrust
    TotalTurboRamJetThrust = mvarTotalTurboRamJetThrust
End Property

Public Property Let TotalRamJetThrustAB(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalRamJetThrustAB = 5
    mvarTotalRamJetThrustAB = vdata
End Property

Public Property Get TotalRamJetThrustAB() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalRamJetThrustAB
    TotalRamJetThrustAB = mvarTotalRamJetThrustAB
End Property

Public Property Let TotalTurboRamJetThrustAB(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalTurboRamJetThrustAB = 5
    mvarTotalTurboramJetThrustAB = vdata
End Property

Public Property Get TotalTurboRamJetThrustAB() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalTurboRamJetThrustAB
    TotalTurboRamJetThrustAB = mvarTotalTurboramJetThrustAB
End Property


Public Property Let wTopSpeed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wTopSpeed = 5
    mvarwTopSpeed = vdata
End Property


Public Property Get wTopSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wTopSpeed
    wTopSpeed = mvarwTopSpeed
End Property

Public Property Let wHydrofoilSpeed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wHydrofoilSpeed = 5
    mvarwHydrofoilSpeed = vdata
End Property


Public Property Get wHydrofoilSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wHydrofoilSpeed
    wHydrofoilSpeed = mvarwHydrofoilSpeed
End Property

Public Property Let wPlaningSpeed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wPlaningSpeed = 5
    mvarwPlaningSpeed = vdata
End Property


Public Property Get wPlaningSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wPlaningSpeed
    wPlaningSpeed = mvarwPlaningSpeed
End Property


Public Property Let wStability(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wStability = 5
    mvarwStability = vdata
End Property

Public Property Get wStability() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wStability
    wStability = mvarwStability
End Property

Public Property Let wManeuverability(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wManeuverability = 5
    mvarwManeuverability = vdata
End Property


Public Property Get wManeuverability() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wManeuverability
    wManeuverability = mvarwManeuverability
End Property

Public Property Let wHydroDrag(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wHydroDrag = 5
    mvarwHydroDrag = vdata
End Property

Public Property Get wHydroDrag() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wHydroDrag
    wHydroDrag = mvarwHydroDrag
End Property

Public Property Let wDraft(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wDraft = 5
    mvarwDraft = vdata
End Property

Public Property Get wDraft() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wDraft
    wDraft = mvarwDraft
End Property

Public Property Let wDeceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wDeceleration = 5
    mvarwDeceleration = vdata
End Property

Public Property Get wDeceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wDeceleration
    wDeceleration = mvarwDeceleration
End Property
Public Property Let wIDeceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wIDeceleration = 5
    mvarwIDeceleration = vdata
End Property

Public Property Get wIDeceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wIDeceleration
    wIDeceleration = mvarwIDeceleration
End Property
Public Property Let wAcceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wAcceleration = 5
    mvarwAcceleration = vdata
End Property

Public Property Get wAcceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wAcceleration
    wAcceleration = mvarwAcceleration
End Property

Public Property Let wTotalAquaticThrust(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wTotalAquaticThrust = 5
    mvarwTotalAquaticThrust = vdata
End Property

Public Property Get wTotalAquaticThrust() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wTotalAquaticThrust
    wTotalAquaticThrust = mvarwTotalAquaticThrust
End Property


Public Sub CalcPerformance()
    mvarAdvisory = ""

    'determine if vehicle has responsive structure
    mbResponsive = VehicleHasResponsiveStruct

    Call GetVehicleWeight(PERFORMANCEWATER, _
                            mvarPercentAuxVehicleWeight, _
                            mvarPercentCargoWeight, _
                            mvarPercentAmmunitionWeight, _
                            mvarPercentHardpointWeight, _
                            mvarPercentFuelWeight, _
                            mvarPercentProvisionWeight, _
                            m_VWeight, m_VMass)
                            
    If mvarTreatTiltRotorsAsPropellers Then
        mvarTiltRotorForwardThrust = GetTiltRotorForwardThrust(mvarPercentThrust)
    End If


    mvarwHydroDrag = CalcHydroDrag
    mvarwTopSpeed = CalcWaterSpeed
    mvarwAcceleration = CalcWaterAcceleration(mvarwTotalAquaticThrust, m_VWeight)
    Call CalcWaterMRandSR(mvarwStability, mvarwManeuverability, mvarKeyChain, mbResponsive)
    Call CalcWaterDeceleration(mvarwManeuverability, mvarwAcceleration, mvarwDeceleration, mvarwIDeceleration)  'this does both Deceleration and Increased Deceleration
    mvarwDraft = CalcDraft
    
End Sub

Function CalcHydroDrag() As Single
    Dim Hl As Integer
    Dim Templift As Single
    Dim MinWeight As Single
    Dim TempWeight As Single
    Dim TempDrag As Single
    
    Templift = CalcTotalContragravLift
    Hl = GetHl
    With Veh.surface
        If .CataTrimaran = Trimaran Then 'todo: need constant for this
            Hl = Hl + (Hl * 0.1)
        ElseIf .CataTrimaran = Catamaran Then
            Hl = Hl + (Hl * 0.2)
        End If
        TempWeight = m_VWeight
    End With
        
    MinWeight = TempWeight * 0.1
    TempWeight = TempWeight - Templift
    If TempWeight < MinWeight Then TempWeight = MinWeight
    TempDrag = ((TempWeight ^ (1 / 3)) ^ 2) / Hl
    CalcHydroDrag = Round(TempDrag, 0)

End Function


Function CalcWaterSpeed() As Single
Dim TempSpeed As Single
Dim dType As String
Dim sKey As String
Dim i As Long
Dim TotalMotivePower As Single
On Error Resume Next

If mvarKeyChain(1) = "" Then Exit Function 'if no propulsion systems, exit the function

For i = 1 To UBound(mvarKeyChain)
sKey = mvarKeyChain(i)
dType = Veh.Components(sKey).Datatype
    ' add motive powers for all other types of thrust components
    Select Case dType
        Case TrackedDrivetrain, LegDrivetrain, WheeledDrivetrain, AllWheelDriveWheeledDrivetrain
            TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotivePower * 2 * mvarPercentThrust)
        Case FlexibodyDrivetrain
            TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotivePower * 5 * mvarPercentThrust)
        Case "" ' debug Need to check if putting this hack in is ok
            'do nothing
        Case RowingPositions, _
                PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, _
                Hydrojet, MHDTunnel, DuctedFan, AerialPropeller, RopeHarness, _
                YokeandPoleHarness, ShaftandCollarHarness, _
                WhiffletreeHarness, ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig, _
                LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                OptimizedFusion, AntimatterThermal, AntimatterPion, _
                StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
            TotalMotivePower = TotalMotivePower + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
        Case Turbojet, Turbofan, Hyperfan, FusionAirRam
                'use Afterburner Thrust if this option is enabled by the user
                If mvarAfterBurnersOn Then
                    'determine if this engine has afterburners or not
                    If Veh.Components(sKey).Afterburner Then
                        TotalMotivePower = TotalMotivePower + Veh.Components(sKey).ABThrust
                    Else
                        TotalMotivePower = TotalMotivePower + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
                    End If
                Else 'use normal engine thrust without afterburner
                    TotalMotivePower = TotalMotivePower + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
                End If
        Case Ramjet
            'store these values for use later in the TopSpeed calculations since
            'Ramjets only work if Topspeed is at least 375mph
            mvartotalramjetthrust = mvartotalramjetthrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
            mvarTotalRamJetThrustAB = mvarTotalRamJetThrustAB + Veh.Components(sKey).ABThrust
            
        Case TurboRamjet
            'store these values for use later since they add .2 x their thrust
            'if the speed is greater than 375mph
            mvarTotalTurboRamJetThrust = mvarTotalTurboRamJetThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
            mvarTotalTurboramJetThrustAB = mvarTotalTurboramJetThrustAB + Veh.Components(sKey).ABThrust
            
            'but in the meantime, they just add their default thrust
            'use Afterburner Thrust if this option is enabled by the user
            If mvarAfterBurnersOn Then
                'determine if this engine has afterburners or not
                If Veh.Components(sKey).Afterburner Then
                    TotalMotivePower = TotalMotivePower + Veh.Components(sKey).ABThrust
                Else
                        TotalMotivePower = TotalMotivePower + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
                End If
            Else 'use normal engine thrust without afterburner
                TotalMotivePower = TotalMotivePower + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
            End If
    End Select
    Next
    '//add in the title rotor forward thrust
    TotalMotivePower = TotalMotivePower + mvarTiltRotorForwardThrust
        
wTotalAquaticThrust = TotalMotivePower 'save the totalaquaticthrust
If mvarwHydroDrag = 0 Then
    TempSpeed = 0 '//can we have a speed of 0 here?
Else
    TempSpeed = ((TotalMotivePower / mvarwHydroDrag) ^ (1 / 3)) * 6
End If
' add in the streamlining effects
With Veh.surface
If TempSpeed > 50 Then
    If .StreamLining = "none" Then
        If TempSpeed > 150 Then TempSpeed = 150
    ElseIf .StreamLining = "fair" Then
        TempSpeed = TempSpeed + (TempSpeed * 0.05)
        If TempSpeed > 150 Then TempSpeed = 150
    ElseIf .StreamLining = "good" Then
        TempSpeed = TempSpeed + (TempSpeed * 0.1)
        If TempSpeed > 150 Then TempSpeed = 150
    Else
        TempSpeed = TempSpeed + (TempSpeed * 0.1)
    End If
End If
End With

'adjust speed if using Ramjets or TurboRamJets
If TempSpeed >= 375 Then
    If mvarAfterBurnersOn Then
        TotalMotivePower = TotalMotivePower + mvarTotalRamJetThrustAB
        TotalMotivePower = TotalMotivePower + (mvarTotalTurboramJetThrustAB * 0.2)
    Else
        TotalMotivePower = TotalMotivePower + mvartotalramjetthrust
        TotalMotivePower = TotalMotivePower + (mvarTotalTurboRamJetThrust * 0.2)
    End If
    TempSpeed = ((TotalMotivePower / wHydroDrag) ^ (1 / 3)) * 6 'Note: do i need to readd streamlining mods here?
End If

' Check to make sure we dont exceed speed of slowest animal(if applicable)
TempSpeed = MinimumNonZero(TempSpeed, GetSlowestAnimalSpeed(mvarKeyChain))

' Do the final rounding
If TempSpeed >= 20 Then
    ' round to nearest 5mph
    TempSpeed = Round(TempSpeed / 5, 0) * 5
Else ' round to nearest whole number
    TempSpeed = Round(TempSpeed, 0)
End If

'TODO: check if hydrofoil and planing top speed calculations are done BEFORE or AFTER
' Streamlining effects adjustments below (see page 131 top right column) If they are "before" then
' i need to move Planing and Hydrofoil speed calcs above the Streamlinging effects section
'TODO: check if the errata "hydrofoil modifier" it speaks of is indeed PlaningSpeed * 1.5

'Get the Planing and Hydrofoil Top Speed
Dim NeededHydrofoilSpeed As Single
Dim NeededPlaningSpeed As Single
Dim TempPlaningSpeed As Single
Dim TempHydrofoilSpeed As Single
Dim element As Object

NeededHydrofoilSpeed = 20 + (Veh.Stats / 100)
NeededPlaningSpeed = ((GetHl * 5) + 5)
If mvarwTotalAquaticThrust >= NeededPlaningSpeed / 100 * m_VWeight Then
    TempPlaningSpeed = TempSpeed * 2
End If
For Each element In Veh
    If TypeOf element Is clsHydrofoil Then
        If TempPlaningSpeed > 0 Then
            TempHydrofoilSpeed = TempPlaningSpeed * 1.5
        Else
            TempHydrofoilSpeed = TempSpeed * 1.5
        End If
        Exit For 'exit this loop after we find a hydrofoil
    End If
Next
wPlaningSpeed = TempPlaningSpeed 'save the calculations
wHydrofoilSpeed = TempHydrofoilSpeed 'save the calculations
'return the function's value
CalcWaterSpeed = TempSpeed
End Function





Function CalcDraft() As Single

    Dim Hl As Single
    Dim TempDraft As Single
    Dim TempWeight As Single
    Dim sLines As String

    TempWeight = m_VWeight
    sLines = Veh.surface.HydrodynamicLines
    
    Select Case sLines
        Case "none"
            Hl = 1
        Case "submarine"
            Hl = 2
        Case "mediocre"
            Hl = 1.1
        Case "average"
            Hl = 1.2
        Case "fine"
            Hl = 1.3
        Case "very fine"
            Hl = 1.4
    End Select

    TempDraft = ((TempWeight ^ (1 / 3)) / 15) * Hl
    CalcDraft = Round(TempDraft, 1) 'round to one decimal place

End Function



Private Sub Class_Initialize()
    ReDim mvarKeyChain(1)
    
    
    mvarAfterBurnersOn = True
    'Private mvarHardPointsOn As Boolean
    'JAW 2000.06.20
    'value of 1 means 100%
    mvarPercentThrust = 1
    mvarPercentCrewWeight = 1
    mvarPercentFuelWeight = 1
    mvarPercentCargoWeight = 1
    mvarPercentHardpointWeight = 1
    mvarPercentProvisionWeight = 1
    mvarPercentAmmunitionWeight = 1
    mvarPercentAuxVehicleWeight = 1
End Sub




