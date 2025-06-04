VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerformanceMagLev"
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

Private mvarmlStallSpeed As Single
Private mvarmlMotiveThrust As Single
Private mvarmlDrag As Single
Private mvarmlTopSpeed As Single
Private mvarmlAcceleration As Single
Private mvarmlManeuverability As Single
Private mvarmlStability As Single
Private mvarmlDeceleration As Single

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

Public Property Get DesignCheckString() As String
    DesignCheckString = g_sDC
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

Public Property Let mlManeuverability(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mlManeuverability = 5
    mvarmlManeuverability = vdata
End Property


Public Property Get mlManeuverability() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mlManeuverability
    mlManeuverability = mvarmlManeuverability
End Property



Public Property Let mlDrag(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mlDrag = 5
    mvarmlDrag = vdata
End Property


Public Property Get mlDrag() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mlDrag
    mlDrag = mvarmlDrag
End Property



Public Property Let mlDeceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mlDeceleration = 5
    mvarmlDeceleration = vdata
End Property


Public Property Get mlDeceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mlDeceleration
    mlDeceleration = mvarmlDeceleration
End Property



Public Property Let mlAcceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mlAcceleration = 5
    mvarmlAcceleration = vdata
End Property


Public Property Get mlAcceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mlAcceleration
    mlAcceleration = mvarmlAcceleration
End Property



Public Property Let mlTopSpeed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mlTopSpeed = 5
    mvarmlTopSpeed = vdata
End Property


Public Property Get mlTopSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mlTopSpeed
    mlTopSpeed = mvarmlTopSpeed
End Property



Public Property Let mlStallSpeed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mlStallSpeed = 5
    mvarmlStallSpeed = vdata
End Property


Public Property Get mlStallSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mlStallSpeed
    mlStallSpeed = mvarmlStallSpeed
End Property



Public Property Let mlStability(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mlStability = 5
    mvarmlStability = vdata
End Property


Public Property Get mlStability() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mlStability
    mlStability = mvarmlStability
End Property



Public Property Let mlMotiveThrust(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mlMotiveThrust = 5
    mvarmlMotiveThrust = vdata
End Property


Public Property Get mlMotiveThrust() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mlMotiveThrust
    mlMotiveThrust = mvarmlMotiveThrust
End Property



Public Sub CalcPerformance()

    Dim sPType As String
    mvarAdvisory = ""
    sPType = mvarDatatype

    'determine if vehicle has responsive structure
    mbResponsive = VehicleHasResponsiveStruct

    Call GetVehicleWeight(PERFORMANCEMAGLEV, _
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

    
        'mvarStaticLift = CalcTotalStaticLift 'todo: needed?
        mvarmlStallSpeed = 0
        mvarmlMotiveThrust = CalcMLMotiveThrust
        mvarmlDrag = CalcADrag(mvarPopTurretsExtended, _
                             mvarWheelsSkidsExtended, _
                             mbResponsive) 'uses same as air performance
                             
        mvarmlTopSpeed = CalcATopSpeed(mvarmlDrag, _
                       mvarmlMotiveThrust, _
                       mvarAfterBurnersOn, _
                       mvartotalramjetthrust, _
                       mvarTotalTurboRamJetThrust, _
                       mvarTotalRamJetThrustAB, _
                       mvarTotalTurboramJetThrustAB) 'uses same as air
                       
        'check for Max speed limits
        mvarmlTopSpeed = CalcAMaxSpeed(mvarmlTopSpeed, _
                                        mvarKeyChain, _
                                        mvarTreatTiltRotorsAsPropellers, _
                                        g_sDC)
        mvarmlTopSpeed = Round(mvarmlTopSpeed / 5, 0) * 5 'round to nearest 5mph

        mvarmlAcceleration = CalcAAcceleration(mvarmlMotiveThrust, m_VWeight)
        mvarmlManeuverability = CalcMLMR
        mvarmlStability = 6 'simple fixed rule
     mvarmlDeceleration = 40 * mvarmlManeuverability 'simple rule

    
End Sub
Function CalcMLMotiveThrust() As Single
Dim i As Integer
Dim TempThrust As Single
Dim sKey As String
Dim dType As Integer

If mvarKeyChain(1) = "" Then Exit Function 'exit if there are no propulsion systems in the keychain

For i = 1 To UBound(mvarKeyChain)
sKey = mvarKeyChain(i)
dType = Veh.Components(sKey).Datatype

    ' add motive powers for all other types of thrust components
    Select Case dType
        Case MagLevLifter
            TempThrust = TempThrust + Veh.Components(sKey).Lift * 0.2
        Case AerialPropeller, DuctedFan, _
                LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                OptimizedFusion, AntimatterThermal, AntimatterPion, _
                StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine
                
            TempThrust = TempThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
    
        Case Turbojet, Turbofan, Hyperfan, FusionAirRam
            'use Afterburner Thrust if this option is enabled by the user
            If mvarAfterBurnersOn Then
                'determine if this engine has afterburners or not
                If Veh.Components(sKey).Afterburner Then
                    TempThrust = TempThrust + Veh.Components(sKey).ABThrust 'PPP This should only include AB if its checked
                Else
                    TempThrust = TempThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
                End If
            Else 'use normal engine thrust without afterburner
                TempThrust = TempThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
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
                    TempThrust = TempThrust + Veh.Components(sKey).ABThrust
                Else
                    TempThrust = TempThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
                End If
            Else 'use normal engine thrust without afterburner
                TempThrust = TempThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
            End If
    End Select
    Next
    
    '//add in any forward thrust from tilt rotors
    TempThrust = TempThrust + mvarTiltRotorForwardThrust
    

CalcMLMotiveThrust = TempThrust
End Function

Function CalcMLMR() As Single
Dim TempMR As Single

If gVehicleTL <= 7 Then
    CalcMLMR = 0.125
Else: CalcMLMR = 0.25
End If
End Function



Private Sub Class_Initialize()
    ReDim mvarKeyChain(1)
    mvarDatatype = PERFORMANCEMAGLEV
    
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





