VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerformanceAir"
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


Private mvarStaticLift As Single
Private mvaraCanFly As Boolean
Private mvaraTakeOffRun As Single
Private mvaraLandingRun As Single
Private mvaraAcceleration As Single
Private mvaraDeceleration As Single
Private mvaraDrag As Single
Private mvaraManeuverability As Single
Private mvaraMotiveThrust As Single
Private mvaraStability As Single
Private mvaraStallSpeed As Single
Private mvaraTopSpeed As Single
Private mvaraReservedVectoredThrust As Single
Private mvarCeiling As Single

Private mvarhAcceleration As Single
Private mvarhDeceleration As Single
Private mvarhSEVSidewalls As Boolean
Private mvarhGEVSkirt As Boolean
Private mvarhHoverAltitude As Single
Private mvarhMotiveThrust As Single
Private mvarhDrag As Single
Private mvarhTopSpeed As Single
Private mvarhManeuverability As Single
Private mvarhStability As Single
Private mvarhReservedHoverThrust As Single '//this doesnt actually need a Get/Let since it doesnt need to be saved in .VEH


Private mvarKey As String
Private mvarParent As String
Private mvarDescription As String
Private mvarDatatype As Long


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



Public Property Let aCanFly(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aCanFly = 5
    mvaraCanFly = vdata
End Property


Public Property Get aCanFly() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aCanFly
    aCanFly = mvaraCanFly
End Property


Public Property Let StaticLift(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StaticLift = 5
    mvarStaticLift = vdata
End Property

Public Property Get StaticLift() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StaticLift
    StaticLift = mvarStaticLift
End Property


Public Property Let aLandingRun(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aLandingRun = 5
    mvaraLandingRun = vdata
End Property

Public Property Get aLandingRun() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aLandingRun
    aLandingRun = mvaraLandingRun
End Property

Public Property Let aTakeOffRun(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aTakeOffRun = 5
    mvaraTakeOffRun = vdata
End Property


Public Property Get aTakeOffRun() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aTakeOffRun
    aTakeOffRun = mvaraTakeOffRun
End Property




Public Property Let aTopSpeed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aTopSpeed = 5
    mvaraTopSpeed = vdata
End Property


Public Property Get aTopSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aTopSpeed
    aTopSpeed = mvaraTopSpeed
End Property



Public Property Let aStallSpeed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aStallSpeed = 5
    mvaraStallSpeed = vdata
End Property


Public Property Get aStallSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aStallSpeed
    aStallSpeed = mvaraStallSpeed
End Property



Public Property Let aStability(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aStability = 5
    mvaraStability = vdata
End Property


Public Property Get aStability() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aStability
    aStability = mvaraStability
End Property



Public Property Let aMotiveThrust(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aMotiveThrust = 5
    mvaraMotiveThrust = vdata
End Property


Public Property Get aMotiveThrust() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aMotiveThrust
    aMotiveThrust = mvaraMotiveThrust
End Property



Public Property Let aManeuverability(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aManeuverability = 5
    mvaraManeuverability = vdata
End Property


Public Property Get aManeuverability() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aManeuverability
    aManeuverability = mvaraManeuverability
End Property



Public Property Let aDrag(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aDrag = 5
    mvaraDrag = vdata
End Property


Public Property Get aDrag() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aDrag
    aDrag = mvaraDrag
End Property



Public Property Let aDeceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aDeceleration = 5
    mvaraDeceleration = vdata
End Property


Public Property Get aDeceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aDeceleration
    aDeceleration = mvaraDeceleration
End Property



Public Property Let aAcceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aAcceleration = 5
    mvaraAcceleration = vdata
End Property


Public Property Get aAcceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aAcceleration
    aAcceleration = mvaraAcceleration
End Property

Public Property Let hstability(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hstability  = 5
    mvarhStability = vdata
End Property


Public Property Get hstability() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hstability
    hstability = mvarhStability
End Property


Public Property Let hmaneuverability(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hmaneuverability  = 5
    mvarhManeuverability = vdata
End Property


Public Property Get hmaneuverability() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hmaneuverability
    hmaneuverability = mvarhManeuverability
End Property

Public Property Let hTopSpeed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hTopSpeed  = 5
    mvarhTopSpeed = vdata
End Property


Public Property Get hTopSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hTopSpeed
    hTopSpeed = mvarhTopSpeed
End Property



Public Property Let hDrag(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hDrag  = 5
    mvarhDrag = vdata
End Property


Public Property Get hDrag() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hDrag
    hDrag = mvarhDrag
End Property


Public Property Let hMotiveThrust(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hMotiveThrust  = 5
    mvarhMotiveThrust = vdata
End Property


Public Property Get hMotiveThrust() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hMotiveThrust
    hMotiveThrust = mvarhMotiveThrust
End Property


Public Property Let hHoverAltitude(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hHoverAltitude = 5
    mvarhHoverAltitude = vdata
End Property


Public Property Get hHoverAltitude() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hHoverAltitude
    hHoverAltitude = mvarhHoverAltitude
End Property


Public Property Let hSEVSidewalls(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hSEVSidewalls = 5
    mvarhSEVSidewalls = vdata
End Property


Public Property Get hSEVSidewalls() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hSEVSidewalls
    hSEVSidewalls = mvarhSEVSidewalls
End Property


Public Property Let hGEVSkirt(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hGEVSkirt = 5
    mvarhGEVSkirt = vdata
End Property


Public Property Get hGEVSkirt() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hGEVSkirt
    hGEVSkirt = mvarhGEVSkirt
End Property

Public Property Let hDeceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hDeceleration = 5
    mvarhDeceleration = vdata
End Property


Public Property Get hDeceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hDeceleration
    hDeceleration = mvarhDeceleration
End Property

Public Property Let hAcceleration(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hAcceleration = 5
    mvarhAcceleration = vdata
End Property

Public Property Get hAcceleration() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hAcceleration
    hAcceleration = mvarhAcceleration
End Property




Public Sub CalcPerformance()
    Dim sHovType As String
    
    mvarAdvisory = ""

    'determine if vehicle has responsive structure
    mbResponsive = VehicleHasResponsiveStruct

    Call GetVehicleWeight(PERFORMANCEAIR, _
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

    
   
    'aerial performance
    mvarStaticLift = CalcTotalStaticLift(mvarKeyChain, _
                                    mvarTreatTiltRotorsAsPropellers, _
                                    mvarhSEVSidewalls, _
                                    mvarhGEVSkirt, _
                                    PERFORMANCEAIR, _
                                    mvarPercentThrust, _
                                    m_VWeight)
                                    
                                    
   
        'this must be done first
        mvaraStallSpeed = CalcAStallSpeed
        '    mvargSpeedFactor = CalcGroundSpeedFactor 'this one cant be an mvar unless i move
        '    mvargTopSpeed = CalcGroundSpeed
        '    mvargAcceleration = CalcGroundAcceleration(mvargSpeedFactor, mvargTopSpeed)
        '    mvaraTakeOffRun = CalcTakeOffRun
        '    CalcGGroundDeceleration
        'mvaraLandingRun = CalcLandingRun
       ' mvaraCanFly = CalcACanFly
        mvaraMotiveThrust = CalcAMotiveThrust(mvarKeyChain, _
                        mvaraReservedVectoredThrust, _
                        mvarTreatTiltRotorsAsPropellers, _
                        mvarTiltRotorForwardThrust, _
                        mvarPercentThrust, _
                        mvarAfterBurnersOn, _
                        mvarDatatype, _
                        mvartotalramjetthrust, _
                        mvarTotalRamJetThrustAB, _
                        mvarTotalTurboRamJetThrust, _
                        mvarTotalTurboramJetThrustAB)
                        
        mvaraDrag = CalcADrag(mvarPopTurretsExtended, _
                             mvarWheelsSkidsExtended, _
                             mbResponsive)
                   
        mvaraTopSpeed = CalcATopSpeed(mvaraDrag, _
                       mvaraMotiveThrust, _
                       mvarAfterBurnersOn, _
                       mvartotalramjetthrust, _
                       mvarTotalTurboRamJetThrust, _
                       mvarTotalRamJetThrustAB, _
                       mvarTotalTurboramJetThrustAB)
                       
        'check for Max speed limits
        mvaraTopSpeed = CalcAMaxSpeed(mvaraTopSpeed, _
                                      mvarKeyChain, _
                                      mvarTreatTiltRotorsAsPropellers, _
                                      g_sDC)
        
        ' round to nearest 5
        mvaraTopSpeed = Round(mvaraTopSpeed / 5, 0) * 5 'round to nearest 5mph

        mvaraAcceleration = CalcAAcceleration(mvaraMotiveThrust, m_VWeight)
        mvaraManeuverability = CalcAManeuverability(mvaraStallSpeed, mbResponsive, m_VWeight)
        mvaraStability = CalcAStability
        mvaraDeceleration = CalcADeceleration(mvaraManeuverability)

  
End Sub
       


'Function CalcTakeOffRun() As Single
'JAW 2000.06.16

'On Error Resume Next
'    If mvargAcceleration < 0.1 Then
'        CalcTakeOffRun = (mvaraStallSpeed ^ 2) / (4 * 0.1)
'    Else
'        CalcTakeOffRun = (mvaraStallSpeed ^ 2) / (4 * mvargAcceleration)
'    End If
'    If CalcTakeOffRun > Int(CalcTakeOffRun) Then CalcTakeOffRun = Int(CalcTakeOffRun) + 1
'
'End Function
    
'Function CalcLandingRun() As Single
''JAW 2000.06.16

'On Error Resume Next
'    'guard against div by 0
'    If mvargDeceleration < 0.1 Then
'        CalcLandingRun = (mvaraStallSpeed ^ 2) / (4 * 0.1)
'    Else
'        CalcLandingRun = (mvaraStallSpeed ^ 2) / (4 * mvargDeceleration)
'    End If
'    If CalcLandingRun > Int(CalcLandingRun) Then CalcLandingRun = Int(CalcLandingRun) + 1
'End Function
    
'Function CalcACanFly() As Boolean
'JAW 2000.06.16
'On Error Resume Next
'    If mvaraStallSpeed <= mvargTopSpeed Then
'        CalcACanFly = True
'    Else
'        CalcACanFly = False
'        Advisory = Advisory & "CANT TAKE OFF."
'    End If
'End Function


Function CalcAStallSpeed() As Single
    Dim Sl As Single
    Dim RS As Single
    Dim TempSpeed As Single
    Dim LiftArea As Single
    Dim element As Object
    Dim sStreamlining As String
    Dim sngVThrustNeeded As Single

'JAW 2000.05.25
'replaced the section that calculates weight in need of support by other means than
'static lift.

sngVThrustNeeded = m_VWeight - mvarStaticLift
If sngVThrustNeeded <= 0 Then
        CalcAStallSpeed = 0
        Exit Function
End If

            'see if total static lift exceeds vehicles loadedweight
            sngVThrustNeeded = m_VWeight - mvarStaticLift

            If sngVThrustNeeded > 0 Then
                mvaraReservedVectoredThrust = GetUseableVectoredThrust(sngVThrustNeeded, _
                                                mvarKeyChain, mvarPercentThrust)
                '//if we already know we have enough static lift to cancel out our loaded
                '//weight then we know our stall speed is 0 and we can exit
                If sngVThrustNeeded <= mvaraReservedVectoredThrust Then
                    CalcAStallSpeed = 0
                    'also, note we will not reset the vectored thrust!
                    Exit Function
                End If
            End If

    'determine total lift area
    For Each element In Veh.Components
        If TypeOf element Is clsWing Then
            If element.SubType = "STOL" Then
                LiftArea = LiftArea + (1.5 * element.SurfaceArea)
            ElseIf element.SubType = "flarecraft" Then
                LiftArea = LiftArea + (3 * element.SurfaceArea)
            Else
                LiftArea = LiftArea + element.SurfaceArea
            End If
        ElseIf TypeOf element Is clsRotor Then
            LiftArea = LiftArea + (3 * element.SurfaceArea)
        End If
    Next
    If Veh.Components(BODY_KEY).LiftingBody Then
        LiftArea = LiftArea + (0.3 * Veh.Components(BODY_KEY).SurfaceArea)
    Else
        LiftArea = LiftArea + (0.1 * Veh.Components(BODY_KEY).SurfaceArea)
    End If
    
    'determine the StreamLining modifier
    With Veh.surface
        sStreamlining = .StreamLining
        If sStreamlining = "none" Then
            Sl = 1
        ElseIf sStreamlining = "fair" Then Sl = 1
        ElseIf sStreamlining = "good" Then Sl = 1.05
        ElseIf sStreamlining = "very good" Then Sl = 1.1
        ElseIf sStreamlining = "superior" Then Sl = 1.15
        ElseIf sStreamlining = "excellent" Then Sl = 1.2
        ElseIf sStreamlining = "radical" Then Sl = 1.3
        End If
    End With
    
    'determine Responsive structure modififer
    If mbResponsive Then
        RS = 1.5
    Else: RS = 2
    End If


    'check for possible divide by zero
    If LiftArea = 0 Then
        CalcAStallSpeed = 0
        Advisory = Advisory & "No Lift Area. "
        Exit Function
    End If
    
    'do final calculation
    TempSpeed = ((m_VWeight - mvarStaticLift - mvaraReservedVectoredThrust) / LiftArea) * Sl * RS
    
    'do final rounding
    If TempSpeed < 0 Then
        CalcAStallSpeed = 0
    Else
        CalcAStallSpeed = Round(TempSpeed / 5, 0) * 5 'round to nearest 5mph
    End If
End Function




Public Property Get DesignCheckString() As String
    DesignCheckString = g_sDC
End Property


Function CalcADeceleration(MR As Single) As Single
    CalcADeceleration = MR * 4
End Function



Private Sub Class_Initialize()
    ReDim mvarKeyChain(1)
    mvarDatatype = PERFORMANCEAIR
       
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







