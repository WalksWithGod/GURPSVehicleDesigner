VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerformanceLeg"
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


Private mvargTotalMotivePower As Double
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
    If vdata > 100 Then
         vdata = 100
    End If
    If vdata < 0 Then
         vdata = 0
    End If
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
    If vdata > 100 Then
         vdata = 100
    End If
    If vdata < 0 Then
         vdata = 0
    End If
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
    If vdata > 100 Then
         vdata = 100
    End If
    If vdata < 0 Then
         vdata = 0
    End If
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
    If vdata > 100 Then
         vdata = 100
    End If
    If vdata < 0 Then
         vdata = 0
    End If
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
    If vdata > 100 Then
         vdata = 100
    End If
    If vdata < 0 Then
         vdata = 0
    End If
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
    If vdata > 100 Then
         vdata = 100
    End If
    If vdata < 0 Then
         vdata = 0
    End If
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
    If vdata > 100 Then
         vdata = 100
    End If
    If vdata < 0 Then
         vdata = 0
    End If
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
    If vdata > 100 Then
         vdata = 100
    End If
    If vdata < 0 Then
         vdata = 0
    End If
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


Public Property Let gTotalMotivePower(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gTotalMotivePower = 5

    mvargTotalMotivePower = vdata
End Property

Public Property Get gTotalMotivePower() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gTotalMotivePower
    gTotalMotivePower = mvargTotalMotivePower
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
        InfoPrint 0, "No Leg Subassembly detected on vehicle.  Cannot caclulate Legged Ground performance."
        Exit Sub
    Else
        If Not TypeOf Veh.Components(mvarMotiveAssemblyKey) Is clsLeg Then
            InfoPrint 0, "No Leg Subassemblies detected on vehicle.  Cannot caclulate Legged Ground performance."
            Exit Sub
        End If
    End If

    'determine if vehicle has responsive structure
    mbResponsive = VehicleHasResponsiveStruct

    Call GetVehicleWeight(PERFORMANCELEG, _
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
Dim i As Long
Dim sKey As String

On Error GoTo errorhandler

    'if there are no propulsion systems on the Keychain, exit the function
    If mvarKeyChain(1) = "" Then
         Exit Function
    End If


    For i = 1 To UBound(mvarKeyChain)
    sKey = mvarKeyChain(i)
    dType = Veh.Components(sKey).Datatype
        ' get the leg drivetrain's motive power
        If dType = LegDrivetrain Then
            mvargTotalMotivePower = mvargTotalMotivePower + Veh.Components(sKey).MotivePower
        End If
    Next

   mvargTotalMotivePower = mvargTotalMotivePower * mvarPercentThrust
   TempSpeed = Sqr(mvargTotalMotivePower / m_VMass)
   TempSpeed = TempSpeed * gSpeedFactor


    ' add in the streamlining effects
    With Veh.surface
    If TempSpeed > 50 Then
        If .StreamLining = "none" Then
            If TempSpeed > 600 Then
                 TempSpeed = 600
            End If
        ElseIf .StreamLining = "fair" Then
            TempSpeed = TempSpeed + (TempSpeed * 0.05)
            If TempSpeed > 600 Then
                 TempSpeed = 600
            End If
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

errorhandler:
    Debug.Print "clsPerformanceLeg.CalcGroundSpeed - Error # " & err.Number & " " & err.Description
    If err.Number = 9 Then  'subscript out of range check for the Keychain if it hasnt been intialized yet
        Exit Function
    End If
End Function

Function CalcGroundSpeedFactor()
Dim Bonus As Integer
Dim legcount As Integer
Dim ImpvdSuspension As Boolean
Dim legarray() As String
Dim element


     'find the improved suspension bonus
    If Veh.Components(mvarMotiveAssemblyKey).ImprovedSuspension Then
        ImpvdSuspension = True
        Bonus = 1
    End If

    ' get the Final total speed factor

        ' Find how many legs are on the vehicle
        legarray = Veh.KeyManager.GetCurrentLegKeys
        legcount = UBound(legarray)

        If legcount = 1 Then
            MsgBox "You have not added at least 2 legs. Using _1_ for Ground Speed Factor"
            CalcGroundSpeedFactor = 1
            Exit Function
        ElseIf legcount = 2 Then
            CalcGroundSpeedFactor = 8 + Bonus
            Exit Function
        ElseIf legcount = 3 Then
            CalcGroundSpeedFactor = 10 + Bonus
            Exit Function
        ElseIf legcount >= 4 Then
            CalcGroundSpeedFactor = 12 + Bonus
            Exit Function
        End If
End Function

      
Function CalcGroundAcceleration(ByVal Speedfactor As Long, ByVal TopSpeed As Single) As Single
Dim TempAcceleration As Long
Dim Bonus As Single
Dim legcount As Integer
Dim legarray() As String
On Error GoTo errorhandler

   ' Find how many legs are on the vehicle
    legarray = Veh.KeyManager.GetCurrentLegKeys
    legcount = UBound(legarray)
    If legarray(1) = "" Or legcount = 1 Then
        Bonus = 0
    ElseIf legcount = 2 Then
        Bonus = 12
    ElseIf legcount = 3 Then
        Bonus = 9.6
    ElseIf legcount >= 4 Then
        Bonus = 8
    End If
    Debug.Assert Speedfactor <> 0
    TempAcceleration = ((TopSpeed / Speedfactor) * 0.8) + Bonus

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
mvargDeceleration = 20 'covers legs and flexibody

End Function

Sub CalcGSRandMR()
'///////////////////////////////////////////////////////////
'now get the GroundStability and Manuever Ratings
Dim MotiveSystem As Integer
Dim TempMR As Single
Dim TempSR As Single
Dim BodyVolume As Single
Dim VehicleWornAsHarness As Boolean
Dim ImpvdSuspension As Boolean
Dim legarray() As String
Dim legcount As Long
Dim element As Object 'used to finding exceptions for sails, harnessed animals, etc.

BodyVolume = Veh.Components(BODY_KEY).Volume


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

    ' Find how many legs are on the vehicle
    legarray = Veh.KeyManager.GetCurrentLegKeys
    legcount = UBound(legarray)

    If legcount = 2 Then
        MotiveSystem = 10
    ElseIf legcount = 3 Then
        MotiveSystem = 11
    ElseIf legcount >= 4 Then
        MotiveSystem = 12
    End If

' get the actual values from the table

If BodyVolume <= 30 Then
    TempMR = GroundStabMatrix(MotiveSystem).M1
    TempSR = GroundStabMatrix(MotiveSystem).S1
ElseIf BodyVolume <= 100 Then
    TempMR = GroundStabMatrix(MotiveSystem).M2
    TempSR = GroundStabMatrix(MotiveSystem).S2
ElseIf BodyVolume <= 300 Then
    TempMR = GroundStabMatrix(MotiveSystem).M3
    TempSR = GroundStabMatrix(MotiveSystem).S3
ElseIf BodyVolume <= 3000 Then
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

    ' add modifiers for legs built at or prior to TL 7
    If gVehicleTL <= 7 Then
         TempSR = TempSR - 1
    End If

' conduct final exception checks for harnessed animals, sails or non-folded wings and Rotors
For Each element In Veh.Components
    If TypeOf element Is clsSail Then
        If TempMR > 0.5 Then 'vehicle with sails cant exceed .5 MR
             TempMR = 0.5
        End If
    ElseIf TypeOf element Is clsHarness Then
        If TempMR > 0.5 Then 'vehicle with harnessed animals cant exceed .5MR
             TempMR = 0.5
        End If
    ElseIf TypeOf element Is clsWing Then
        If element.Folding <> True Then
            If TempMR > 0.5 Then 'vehicles with non folded wings limited to .5MR
                 TempMR = 0.5
            End If
        End If
    ElseIf TypeOf element Is clsRotor Then
        If element.Folding <> True Then
            If TempMR > 0.5 Then 'vehicles with non folded Rotors limited to .5MR
                 TempMR = 0.5
            End If
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
Dim legarray() As String
Dim i As Long

If mvarKeyChain(1) = "" Then 'exit sub if the keychain does not contain any propulsion systems
    mvargPressureDescription = ""
    mvargOffRoad = 0
    mvargPressure = 0
    Exit Sub
End If



legarray = Veh.KeyManager.GetCurrentLegKeys
If legarray(1) <> "" Then
    For i = 1 To UBound(legarray)
        TempContactArea = TempContactArea + Veh.Components(legarray(i)).SurfaceArea
    Next
End If
ContactArea = TempContactArea / 12.5

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
If TempWeight < MinWeight Then 'make sure contragrav reduction leaves at least 10% of loaded weight
     TempWeight = MinWeight
End If

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

' should always be ONE for Leg performance
Second = 1

' get off road speed
If Second = 5 Then
    mvargOffRoad = 0 ' if it has railway wheels
Else
    mvargOffRoad = mvargTopSpeed * arrPT(First, Second)
End If

End Sub




Private Sub Class_Initialize()
    ReDim mvarKeyChain(1)
    mvarDatatype = PERFORMANCELEG

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








