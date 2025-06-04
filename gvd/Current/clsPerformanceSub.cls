VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerformanceSubmerged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mvarsuAcceleration As Single
Private mvarsuCrushDepth As Single
Private mvarsuDeceleration As Single
Private mvarsuIDeceleration As Single
Private mvarsuHydroDrag As Single
Private mvarsuDraft As Single
Private mvarsuManeuverability As Single
Private mvarsuStability As Single
Private mvarsuTopSpeed As Single
Private mvarsuTotalAquaticThrust As Single


Private mvarKey As String
Private mvarParent As String
Private mvarDatatype As Long
Private mvarMotiveAssembly As String
Private mvarMotiveAssemblyKey As String
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


Public Property Let suTopSpeed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suTopSpeed = 5
    mvarsuTopSpeed = vdata
End Property

Public Property Get suTopSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suTopSpeed
    suTopSpeed = mvarsuTopSpeed
End Property

Public Property Let suTotalAquaticThrust(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suTotalAquaticThrust = 5
    mvarsuTotalAquaticThrust = vdata
End Property

Public Property Get suTotalAquaticThrust() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suDeceleration
    suTotalAquaticThrust = mvarsuTotalAquaticThrust
End Property

Public Property Let suStability(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suStability = 5
    mvarsuStability = vdata
End Property

Public Property Get suStability() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suStability
    suStability = mvarsuStability
End Property


Public Property Let suManeuverability(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suManeuverability = 5
    mvarsuManeuverability = vdata
End Property

Public Property Get suManeuverability() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suManeuverability
    suManeuverability = mvarsuManeuverability
End Property

Public Property Let suDraft(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suDraft = 5
    mvarsuDraft = vdata
End Property

Public Property Get suDraft() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suDraft
    suDraft = mvarsuDraft
End Property

Public Property Let suHydroDrag(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suHydroDrag = 5
    mvarsuHydroDrag = vdata
End Property

Public Property Get suHydroDrag() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suHydroDrag
    suHydroDrag = mvarsuHydroDrag
End Property

Public Property Let suDeceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suDeceleration = 5
    mvarsuDeceleration = vdata
End Property

Public Property Get suDeceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suDeceleration
    suDeceleration = mvarsuDeceleration
End Property

Public Property Let suIDeceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suIDeceleration = 5
    mvarsuIDeceleration = vdata
End Property

Public Property Get suIDeceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suIDeceleration
    suIDeceleration = mvarsuIDeceleration
End Property

Public Property Let suCrushDepth(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suCrushDepth = 5
    mvarsuCrushDepth = vdata
End Property

Public Property Get suCrushDepth() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suCrushDepth
    suCrushDepth = mvarsuCrushDepth
End Property

Public Property Let suAcceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suAcceleration = 5
    mvarsuAcceleration = vdata
End Property

Public Property Get suAcceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suAcceleration
    suAcceleration = mvarsuAcceleration
End Property



Public Sub CalcPerformance()
    mvarAdvisory = ""


    'determine if vehicle has responsive structure
    mbResponsive = VehicleHasResponsiveStruct

    Call GetVehicleWeight(PERFORMANCESUB, _
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


    mvarsuHydroDrag = CalcSuHydroDrag
    mvarsuTopSpeed = CalcSuTopSpeed
    mvarsuAcceleration = CalcWaterAcceleration(mvarsuTotalAquaticThrust, m_VWeight)
    Call CalcWaterMRandSR(mvarsuStability, mvarsuManeuverability, mvarKeyChain, mbResponsive)
    Call CalcWaterDeceleration(mvarsuManeuverability, mvarsuAcceleration, mvarsuDeceleration, mvarsuIDeceleration)  'this does both Deceleration and Increased Deceleration
    
    mvarsuDraft = CalcSuDraft
    mvarsuCrushDepth = CalcCrushDepth

    
End Sub


Function CalcSuHydroDrag() As Single
    Dim Ls As Integer
    Dim TempDrag As Single
    Dim TempWeight As Single
    Dim sLines As String
    
    sLines = Veh.surface.HydrodynamicLines
    
    Select Case sLines
        Case "submarine"
            Ls = 10
        Case "very fine"
            Ls = 6
        Case "fine"
            Ls = 4
        Case "average"
            Ls = 3
        Case "mediocre"
            Ls = 2
        Case "none"
            Ls = 1
        End Select
        TempWeight = m_VWeight
    End With
        
    TempDrag = ((TempWeight ^ (1 / 3)) ^ 2) / Ls
    CalcSuHydroDrag = Round(TempDrag, 0)
End Function


Function CalcSuTopSpeed() As Single
    Dim Animal As Boolean
    Dim SlowestAnimal As Single
    Dim TempSpeed As Single
    Dim dType As String
    Dim sKey As String
    Dim i As Long
    Dim TotalMotivePower As Single
    
    If mvarKeyChain(1) = "" Then
        mvarsuTotalAquaticThrust = 0
        Exit Function 'exit if there are no propulsion systems in the keychain
    End If
    
    For i = 1 To UBound(mvarKeyChain)
    sKey = mvarKeyChain(i)
    dType = Veh.Components(sKey).Datatype
        ' add motive powers for all other types of thrust components
        Select Case dType
            Case LegDrivetrain
                TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotivePower * 2 * mvarPercentThrust)
            Case FlexibodyDrivetrain
                TotalMotivePower = TotalMotivePower + (Veh.Components(sKey).MotivePower * 5 * mvarPercentThrust)
            'Divide the remaining thrust components into those that use MotivePower and
            'those that use Motive Thrust
            Case PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, _
                    Hydrojet, MHDTunnel, RopeHarness, _
                    YokeandPoleHarness, ShaftandCollarHarness, _
                    WhiffletreeHarness, _
                    LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                    OptimizedFusion, AntimatterThermal, AntimatterPion, _
                    StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
                    
                TotalMotivePower = TotalMotivePower + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
        End Select
        Next
    mvarsuTotalAquaticThrust = TotalMotivePower 'save the submerged totalaquaticthrust
    If mvarsuHydroDrag = 0 Then
        TempSpeed = 0
    Else
        TempSpeed = ((TotalMotivePower / mvarsuHydroDrag) ^ (1 / 3)) * 6
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
    
    'return the function's value
    CalcSuTopSpeed = TempSpeed
End Function


Function CalcSuDraft() As Single
    CalcSuDraft = ((m_VWeight) ^ (1 / 3)) / 3
End Function


Function CalcCrushDepth() As Single
Dim element As Object
Dim LowestDR As Long
Dim childElement As Object
Dim TempCrush As Single
Dim SMod As Single
Dim bArmored As Boolean
LowestDR = 0 'init

'On Error Resume Next
Dim sParentKey As String
For Each element In Veh.Components
    ' find all subassemblies with accomodations or crew stations
    If (TypeOf element Is clsAccommodation) Or (TypeOf element Is clsCrewStation) Then
        sParentKey = element.LogicalParent 'get the key of the subassembly this is contained in
        bArmored = False
        ' find the armor DR of it
        For Each childElement In Veh.Components
            If TypeOf childElement Is clsArmor Then
                If childElement.LogicalParent = sParentKey Then
                    LowestDR = MinimumNonZero(childElement.GetLowestCrushDepthDR, LowestDR)
                    bArmored = True
                End If
            End If
        Next
        
        '//the below is commented out because it will cause OveralArmor which are attached tojust the Body
        '//yet cover the entire vehicle to have 0 DR's for all other subassemblies
         '//todo: In future, i can set a Flag if Overall armor is detected.  If not detected, then i can use 0
        '//for items that dont have any armor but should.
        'If Not bArmored Then
        '    LowestDR = 0 'this subassembly has no DR
        '    Exit For
        'End If
    End If
Next
        
'get the structure modifier
With Veh.Components(BODY_KEY) 'todo do i have to find the weakest subassembly or just the body here?
    If .FrameStrength = "super-light" Then
        SMod = 0.1
    ElseIf .FrameStrength = "extra-light" Then
        SMod = 0.25
    ElseIf .FrameStrength = "light" Then
        SMod = 0.5
    ElseIf .FrameStrength = "medium" Then
        SMod = 1
    ElseIf .FrameStrength = "heavy" Then
        SMod = 2
    ElseIf .FrameStrength = "extra-heavy" Then
        SMod = 4
    End If
End With

If bArmored Then

    'do final calculations to yield Crush Depth in yards
    TempCrush = (LowestDR + 10) * SMod * 10
    If Veh.surface.Submersible = False Then
        TempCrush = TempCrush / 2
    End If
Else
    ' there is no crush depth which we represent as -1
    ' in the output, if we see a -1, then we type out "No Crush Depth"
    TempCrush = -1
End If


CalcCrushDepth = Round(TempCrush, 0) 'round to nearest whole number

End Function


Function CalcDraft() As Single

Dim Hl As Single
Dim TempDraft As Single
Dim TempWeight As Single
Dim sLines As String

sLines = Veh.surface.HydrodynamicLines
TempWeight = m_VWeight
 
Select Case sLines
    Case "none"
        Hl = 1
    Case "mediocre"
        Hl = 1.1
    Case "average"
        Hl = 1.2
    Case "fine"
        Hl = 1.3
    Case "very fine"
        Hl = 1.4
    Case "submarine"
        Hl = 2
    Case Else
        Debug.Print "clsPerformanceSubmerged:CalcDraft() -- ERROR.  Invalid Case"
    End Select
End With
TempDraft = ((TempWeight ^ (1 / 3)) / 15) * Hl
CalcDraft = Round(TempDraft, 1) 'round to one decimal place

End Function


Private Sub Class_Initialize()
    ReDim mvarKeyChain(1)
    mvarDatatype = PERFORMANCEPROFILE
    
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






