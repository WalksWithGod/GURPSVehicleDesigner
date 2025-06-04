VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerformanceWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'local variable(s) to hold property value(s)
Private mvartotalramjetthrust As Single
Private mvarTotalTurboRamJetThrust As Single
Private mvarTotalRamJetThrustAB As Single
Private mvarTotalTurboramJetThrustAB As Single


Private mvargSpeedFactor As Long
Private mvargAcceleration As Single
Private mvargDeceleration As Single
Private mvargManeuverability As Single
Private mvargOffRoad As Single
Private mvargPressure As Single
Private mvargPressureDescription As String
Private mvargStability As Single
Private mvargTopSpeed As Single

Private mvarKey As String
Private mvarParent As String
Private mvarDatatype As Long
Private mvarDescription As String
Private mvarMotiveAssemblyKey As String

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

Public Property Let MotiveAssemblyKey(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MotiveAssemblyKey = 5
    mvarMotiveAssemblyKey = vdata
End Property


Public Property Get MotiveAssemblyKey() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MotiveAssemblyKey
    MotiveAssemblyKey = mvarMotiveAssemblyKey
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



Public Property Let gSpeedFactor(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gSpeedFactor = 5

    mvargSpeedFactor = vdata
End Property

Public Property Get gSpeedFactor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gSpeedFactor
    gSpeedFactor = mvargSpeedFactor
End Property




Public Property Let gTopSpeed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gTopSpeed = 5
    mvargTopSpeed = vdata
End Property

Public Property Get gTopSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gTopSpeed
    gTopSpeed = mvargTopSpeed
End Property

Public Property Let gStability(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gStability = 5
    mvargStability = vdata
End Property

Public Property Get gStability() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gStability
    gStability = mvargStability
End Property

Public Property Let gPressure(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gPressure = 5
    mvargPressure = vdata
End Property


Public Property Get gPressure() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gPressure
    gPressure = mvargPressure
End Property


Public Property Let gPressureDescription(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gPressureDescription = 5
    mvargPressureDescription = vdata
End Property


Public Property Get gPressureDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gPressureDescription
    gPressureDescription = mvargPressureDescription
End Property

Public Property Let gOffRoad(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gOffRoad = 5
    mvargOffRoad = vdata
End Property


Public Property Get gOffRoad() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gOffRoad
    gOffRoad = mvargOffRoad
End Property

Public Property Let gManeuverability(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gManeuverability = 5
    mvargManeuverability = vdata
End Property


Public Property Get gManeuverability() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gManeuverability
    gManeuverability = mvargManeuverability
End Property



Public Property Let gDeceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gDeceleration = 5
    mvargDeceleration = vdata
End Property


Public Property Get gDeceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gDeceleration
    gDeceleration = mvargDeceleration
End Property



Public Property Let gAcceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gAcceleration = 5
    mvargAcceleration = vdata
End Property


Public Property Get gAcceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gAcceleration
    gAcceleration = mvargAcceleration
End Property

Public Sub CalcPerformance()

    mvarAdvisory = ""
   
    mvarMotiveAssemblyKey = GetMotiveAssemblyKey(mvarDatatype)
    
    If mvarMotiveAssemblyKey = "" Then
        InfoPrint 0, "No Wheel Subassembly detected on vehicle.  Cannot caclulate Wheeled Ground performance."
        Exit Sub
    Else
        If Not TypeOf Veh.Components(mvarMotiveAssemblyKey) Is clsWheel Then
            InfoPrint 0, "No Wheel Subassembly detected on vehicle.  Cannot caclulate Wheeled Ground performance."
            Exit Sub
        End If
    End If
    
    'determine if vehicle has responsive structure
    mbResponsive = VehicleHasResponsiveStruct

    Call GetVehicleWeight(PERFORMANCEWHEEL, _
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

    mvargSpeedFactor = CalcGroundSpeedFactor 'this one cant be an mvar unless i move
    mvargTopSpeed = CalcGroundSpeed
    mvargAcceleration = CalcGroundAcceleration(mvargSpeedFactor, mvargTopSpeed)
    CalcGGroundDeceleration
    CalcGSRandMR
    CalcGPressureandOffRoadSpeed
    
End Sub

Function CalcGroundSpeed() As Single
Dim TempSpeed As Single
Dim dType As String
Dim TotalMotivePower As Single
Dim i As Long
Dim sKey As String

On Error GoTo errorhandler

'if there are no propulsion systems on the Keychain, exit the function
If mvarKeyChain(1) = "" Then Exit Function


For i = 1 To UBound(mvarKeyChain)
    sKey = mvarKeyChain(i)
    dType = Veh.Components(sKey).Datatype
    ' add motive powers for all other types of thrust components
    Select Case dType
        Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, _
            RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness
            
            TotalMotivePower = TotalMotivePower + Veh.Components(sKey).MotivePower * mvarPercentThrust
        Case Turbojet, Turbofan, Hyperfan, FusionAirRam
            'use Afterburner Thrust if this option is enabled by the user
            If mvarAfterBurnersOn Then
                'determine if this engine has afterburners or not
                If Veh.Components(sKey).Afterburner Then
                    TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).ABThrust / 4)
                Else
                    TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotiveThrust / 4 * mvarPercentThrust)
                End If
            Else 'use normal engine thrust without afterburner
                TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotiveThrust / 4 * mvarPercentThrust)
            End If
        Case Ramjet
            'store these values for use later in the TopSpeed calculations since
            'Ramjets only work if Topspeed is at least 375mph
            mvartotalramjetthrust = mvartotalramjetthrust + (Veh.Components(sKey).MotiveThrust / 4 * mvarPercentThrust)
            mvarTotalRamJetThrustAB = mvarTotalRamJetThrustAB + (Veh.Components(sKey).ABThrust / 4)
            
        Case TurboRamjet
            'store these values for use later since they add .2 x their thrust
            'if the speed is greater than 375mph
            mvarTotalTurboRamJetThrust = mvarTotalTurboRamJetThrust + (Veh.Components(sKey).MotiveThrust / 4 * mvarPercentThrust)
            mvarTotalTurboramJetThrustAB = mvarTotalTurboramJetThrustAB + (Veh.Components(sKey).ABThrust / 4)
            
            'but in the meantime, they just add their default thrust
            'use Afterburner Thrust if this option is enabled by the user
            If mvarAfterBurnersOn Then
                'determine if this engine has afterburners or not
                If Veh.Components(sKey).Afterburner Then
                    TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).ABThrust / 4)
                Else
                    TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotiveThrust / 4 * mvarPercentThrust)
                End If
            Else 'use normal engine thrust without afterburner
                TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotiveThrust / 4 * mvarPercentThrust)
            End If
        Case Else
            TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotiveThrust / 4 * mvarPercentThrust)
        End Select
Next
'//add in the title rotor forward thrust
TotalMotivePower = TotalMotivePower + mvarTiltRotorForwardThrust

'check for divide by zero
If Veh.Stats.LoadedMass = 0 Then
    TempSpeed = 0
Else
    TempSpeed = Sqr(TotalMotivePower / m_VMass)
End If
TempSpeed = TempSpeed * gSpeedFactor

'adjust speed for Wheels/Skids if using Ramjets or TurboRamJets

If TempSpeed >= 375 Then
    If mvarAfterBurnersOn Then
        TotalMotivePower = TotalMotivePower + mvarTotalRamJetThrustAB
        TotalMotivePower = TotalMotivePower + (mvarTotalTurboramJetThrustAB * 0.2)
    Else
        TotalMotivePower = TotalMotivePower + mvartotalramjetthrust
        TotalMotivePower = TotalMotivePower + (mvarTotalTurboRamJetThrust * 0.2)
    End If
    TempSpeed = Sqr(TotalMotivePower / m_VMass)
    TempSpeed = TempSpeed * gSpeedFactor
End If

 
' add in the streamlining effects
With Veh.surface
If TempSpeed > 50 Then
    If .StreamLining = "none" Then
        If TempSpeed > 600 Then TempSpeed = 600
    ElseIf .StreamLining = "fair" Then
        TempSpeed = TempSpeed + (TempSpeed * 0.05)
        If TempSpeed > 600 Then TempSpeed = 600
    Else
        TempSpeed = TempSpeed + (TempSpeed * 0.1)
    End If
End If
End With

' Check to make sure we dont exceed speed of slowest animal(if applicable)
TempSpeed = MinimumNonZero(TempSpeed, GetSlowestAnimalSpeed(mvarKeyChain))

' Do the final rounding
If TempSpeed >= 20 Then
    ' round to nearest 5mph
    CalcGroundSpeed = Round(TempSpeed / 5, 0) * 5
Else ' round to nearest whole number
    CalcGroundSpeed = Round(TempSpeed, 0)
End If
Exit Function

errorhandler:
    Debug.Print "clsPerformanceWheel.CalcGroundSpeed - Error # " & err.Number & " " & err.Description
    
    If err.Number = 9 Then  'subscript out of range check for the Keychain if it hasnt been intialized yet
        Exit Function
    End If
End Function

Function CalcGroundSpeedFactor() As Long
Dim TempKey As String
Dim Bonus As Integer
Dim RailBonus As Boolean
Dim legcount As Integer
Dim ImpvdSuspension As Boolean
Dim legarray() As String
Dim element

TempKey = MotiveAssemblyKey

    
    If TempKey = "" Then
        CalcGroundSpeedFactor = 2
        Advisory = Advisory & "NO LANDING GEAR CHOSEN. "
        Exit Function
    End If

    'find the improved suspension bonus
    If Veh.Components(TempKey).ImprovedSuspension Then
        ImpvdSuspension = True
        Bonus = 2
    End If

    'Determine if there is a railway bonus
    If Veh.Components(TempKey).SubType = "railway" Then
        If gVehicleTL >= 5 Then
            RailBonus = True
        Else
            RailBonus = False
        End If
    End If

    ' get the Final total speed factor
        If gVehicleTL <= 4 Then
            CalcGroundSpeedFactor = 8 + Bonus
            Exit Function
        ElseIf gVehicleTL = 5 Then
            If RailBonus Then
                CalcGroundSpeedFactor = 2 * (12 + Bonus)
                Exit Function
            Else
                CalcGroundSpeedFactor = 12 + Bonus
                Exit Function
            End If
        ElseIf gVehicleTL >= 6 Then
            If RailBonus Then
                CalcGroundSpeedFactor = 2 * (16 + Bonus)
                Exit Function
            Else
                CalcGroundSpeedFactor = 16 + Bonus
                Exit Function
            End If
        End If
    
        
End Function


Function CalcGroundAcceleration(ByVal Speedfactor As Long, ByVal TopSpeed As Single) As Single
Dim TempAcceleration As Long

On Error GoTo errorhandler

    If Speedfactor > 0 Then
        TempAcceleration = ((TopSpeed / Speedfactor) * 0.8)
    Else
        Debug.Print "clsPerformanceWheel:CalcGroundAcceleration -- Speedfactor = 0: Divide By Zero."
    End If
' Do the rounding
If TempAcceleration > 5 Then
    mvargAcceleration = Round(TempAcceleration / 5, 0) * 5 'to nearest 5mph
Else
    mvargAcceleration = Round(TempAcceleration, 0) 'to nearest 1mph
End If

CalcGroundAcceleration = TempAcceleration
Exit Function
errorhandler:
    Exit Function
End Function

Function CalcGGroundDeceleration()


'//////////////////////////////////////////////////////
'now get the Ground Deceleration

    'check for improved breaks with smartwheels
    If (Veh.Components(mvarMotiveAssemblyKey).ImprovedBrakes) And (Veh.Components(mvarMotiveAssemblyKey).Smartwheels) Then
        mvargDeceleration = 20
    'check for either or
    ElseIf (Veh.Components(mvarMotiveAssemblyKey).ImprovedBrakes) Or (Veh.Components(mvarMotiveAssemblyKey).Smartwheels) Then
        mvargDeceleration = 15
    Else ' neither smartwheels or improved found
        mvargDeceleration = 10
    End If

End Function

Sub CalcGSRandMR()
'///////////////////////////////////////////////////////////
'now get the GroundStability and Manuever Ratings
Dim MotiveSystem As Integer
Dim TempMR As Single
Dim TempSR As Single
Dim TempVolume As Single
Dim VehicleWornAsHarness As Boolean
Dim ImpvdSuspension As Boolean
Dim legarray() As String
Dim legcount As Long
Dim element As Object 'used to finding exceptions for sails, harnessed animals, etc.

TempVolume = Veh.Components(BODY_KEY).Volume

    If Veh.Components(mvarMotiveAssemblyKey).ImprovedSuspension Then
        ImpvdSuspension = True
    End If


' Determine which Motive System to use
'1 = skid
'2 = Wheels 1
'3 = Wheels 2
'4 = Wheels 3
'5 = wheels 4-7
'6 = Wheels 8+
'7 = tracks
'8 = skitracks
'9 = halftracks
'10 = 2 legs
'11 = 3 legs
'12 = 4+ legs
'13 = Flexibody

    With Veh.Components(mvarMotiveAssemblyKey)
        If .Quantity = 1 Then
            MotiveSystem = 2
        ElseIf .Quantity = 2 Then
            MotiveSystem = 3
        ElseIf .Quantity = 3 Then
            MotiveSystem = 4
        ElseIf .Quantity <= 6 Then
            MotiveSystem = 5
        Else: MotiveSystem = 6
        End If
    End With

' get the actual values from the table

If TempVolume <= 30 Then
    TempMR = GroundStabMatrix(MotiveSystem).M1
    TempSR = GroundStabMatrix(MotiveSystem).S1
ElseIf TempVolume <= 100 Then
    TempMR = GroundStabMatrix(MotiveSystem).M2
    TempSR = GroundStabMatrix(MotiveSystem).S2
ElseIf TempVolume <= 300 Then
    TempMR = GroundStabMatrix(MotiveSystem).M3
    TempSR = GroundStabMatrix(MotiveSystem).S3
ElseIf TempVolume <= 3000 Then
    TempMR = GroundStabMatrix(MotiveSystem).M4
    TempSR = GroundStabMatrix(MotiveSystem).S4
Else
    TempMR = GroundStabMatrix(MotiveSystem).M5
    TempSR = GroundStabMatrix(MotiveSystem).S5
End If
'//////////////////////////////////////////////////////
' apply the gMR and gSR modifiers
If ImpvdSuspension Then
    TempSR = TempSR + 1 'increase the gSR
    If TempMR = 0.125 Then
        TempMR = 0.25
    Else
        TempMR = TempMR + 0.25
    End If
End If

'add mods for Electronic or Computer controls
If VehiclehasElectORCompcontrols Then
    If TempMR = 0.125 Then
        TempMR = 0.25
    Else
        TempMR = TempMR + 0.25
    End If
End If

'add modifier for responsive structure.
If mbResponsive Then
    If TempMR = 0.125 Then
       TempMR = 0.25
    Else
        TempMR = TempMR + 0.25
    End If
End If
'add modifiers for smartwheels and all wheel steering
With Veh.Components(mvarMotiveAssemblyKey)

    If .Smartwheels Then
        If TempMR = 0.125 Then
            TempMR = 0.25
        Else
            TempMR = TempMR + 0.25
        End If
    End If
    If .AllwheelSteering Then
        If TempMR = 0.125 Then
            TempMR = 0.25
        Else
            TempMR = TempMR + 0.25
        End If
    End If
    If .SubType = "small" Then
        TempSR = TempSR + 1
        If TempMR > 0.5 Then TempMR = 0.5 ' vehicle with small wheels cant exceed .5 MR
    ElseIf .SubType = "railway" Then
        If TempMR > 0.5 Then TempMR = 0.5 ' vehicle with railway wheels cant exceed .5 MR
    End If
End With

' conduct final exception checks for harnessed animals, sails or non-folded wings and Rotors
For Each element In Veh.Components
    If TypeOf element Is clsSail Then
        If TempMR > 0.5 Then TempMR = 0.5 'vehicle with sails cant exceed .5 MR
    ElseIf TypeOf element Is clsHarness Then
        If TempMR > 0.5 Then TempMR = 0.5 'vehicle with harnessed animals cant exceed .5MR
    ElseIf TypeOf element Is clsWing Then
        If element.Folding <> True Then
            If TempMR > 0.5 Then TempMR = 0.5 'vehicles with non folded wings limited to .5MR
        End If
    ElseIf TypeOf element Is clsRotor Then
        If element.Folding <> True Then
            If TempMR > 0.5 Then TempMR = 0.5 'vehicles with non folded Rotors limited to .5MR
        End If
    ElseIf TypeOf element Is clsCrewStation Then
        If element.Datatype = HarnessCrewStation Then
            VehicleWornAsHarness = True
        End If
    End If
Next
'vehicles worn as harness are SR - 1
If VehicleWornAsHarness Then
    TempSR = TempSR - 1
End If

'Save the final gMR and gSR results
gManeuverability = TempMR
gStability = TempSR

End Sub

Sub CalcGPressureandOffRoadSpeed()
'////////////////////////////////////////////////////////////
'' get the contact area
Dim element As Object
Dim ContactArea As Single
Dim TempContactArea As Single
Dim dType As Integer
Dim i As Long

If mvarKeyChain(1) = "" Then 'exit sub if the keychain does not contain any propulsion systems
    mvargPressureDescription = ""
    mvargOffRoad = 0
    mvargPressure = 0
    Exit Sub
End If


If Veh.Components(mvarMotiveAssemblyKey).SubType = "railway" Then
    ContactArea = Veh.Components(mvarMotiveAssemblyKey).SurfaceArea / 66
ElseIf Veh.Components(mvarMotiveAssemblyKey).SubType = "off-road" Then
    ContactArea = Veh.Components(mvarMotiveAssemblyKey).SurfaceArea / 33
Else
    If gVehicleTL >= 6 Then
        ContactArea = Veh.Components(mvarMotiveAssemblyKey).SurfaceArea / 50
    Else
        ContactArea = Veh.Components(mvarMotiveAssemblyKey).SurfaceArea / 66
    End If
End If


' get the ground pressure
Dim TempWeight As Single
Dim MinWeight As Single ' contragrav cant reduce weight to less than 10% of original
Dim sDescription As String
Dim TempPressure As Single
Dim First As Integer
Dim Second As Integer
Dim arrPT(7, 28)
'fill the pressure table
arrPT(1, 1) = 1
arrPT(1, 2) = 1
arrPT(1, 3) = 4 / 5
arrPT(1, 4) = 2 / 3
arrPT(2, 1) = 1
arrPT(2, 2) = 4 / 5
arrPT(2, 3) = 2 / 3
arrPT(2, 4) = 0.5
arrPT(3, 1) = 4 / 5
arrPT(3, 2) = 2 / 3
arrPT(3, 3) = 0.5
arrPT(3, 4) = 1 / 3
arrPT(4, 1) = 2 / 3
arrPT(4, 2) = 0.5
arrPT(4, 3) = 1 / 3
arrPT(4, 4) = 1 / 4
arrPT(5, 1) = 0.5
arrPT(5, 2) = 1 / 3
arrPT(5, 3) = 1 / 4
arrPT(5, 4) = 1 / 6
arrPT(6, 1) = 1 / 3
arrPT(6, 2) = 1 / 4
arrPT(6, 3) = 1 / 6
arrPT(6, 4) = 1 / 8
arrPT(7, 1) = 1 / 4
arrPT(7, 2) = 1 / 6
arrPT(7, 3) = 1 / 8
arrPT(7, 4) = 0

TempWeight = m_VWeight
MinWeight = TempWeight * 0.1
TempWeight = TempWeight - CalcTotalContragravLift
If TempWeight < MinWeight Then TempWeight = MinWeight 'make sure contragrav reduction leaves at least 10% of loaded weight

If ContactArea = 0 Then
    TempWeight = 0
Else
    TempWeight = TempWeight / ContactArea 'this gives us our actual ground pressure
End If
mvargPressure = TempWeight

If TempWeight <= 150 Then
    First = 1
    mvargPressureDescription = "extremely low"
ElseIf TempWeight <= 900 Then
    First = 2
    mvargPressureDescription = "very low"
ElseIf TempWeight <= 1800 Then
    First = 3
    mvargPressureDescription = "low"
ElseIf TempWeight <= 2700 Then
    First = 4
    mvargPressureDescription = "moderate"
ElseIf TempWeight <= 7500 Then
    First = 5
    mvargPressureDescription = "high"
ElseIf TempWeight <= 15000 Then
    First = 6
    mvargPressureDescription = "very high"
Else
    First = 7
    mvargPressureDescription = "extremely high"
End If


'TODO check ruling about vehicles with small wheels. See page 130!
'if small wheels do count, then the below code is complete.
'or if small wheels count but have reduced speed, i must add that code
If mvarKeyChain(1) <> "" Then
    For i = 1 To UBound(KeyChain)
        If Veh.Components(mvarKeyChain(i)).Datatype = AllWheelDriveWheeledDrivetrain Then
            Second = 3
            Exit For
        Else: Second = 4
        End If
    Next
End If

' get off road speed
If Second = 5 Then
    mvargOffRoad = 0 ' if it has railway wheels
Else
    mvargOffRoad = mvargTopSpeed * arrPT(First, Second)
End If

End Sub

Private Sub Class_Initialize()
    ReDim mvarKeyChain(1)
    mvarDatatype = PERFORMANCEWHEEL
    
    
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







