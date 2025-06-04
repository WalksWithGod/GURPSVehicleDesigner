VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerformanceHover"
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
Private mvarStaticLift As Single 'todo: i think i can delete the get/lets for this right in Air too?

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


Public Sub CalcPerformance()

    Dim sPType As String
    mvarAdvisory = ""
    sPType = mvarDatatype

    'determine if vehicle has responsive structure
    mbResponsive = VehicleHasResponsiveStruct

    Call GetVehicleWeight(PERFORMANCEHOVER, _
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

    
        Dim sHovType As String
    
        mvarStaticLift = CalcTotalStaticLift(mvarKeyChain, _
                                    mvarTreatTiltRotorsAsPropellers, _
                                    mvarhSEVSidewalls, _
                                    mvarhGEVSkirt, _
                                    PERFORMANCEAIR, _
                                    mvarPercentThrust, _
                                    m_VWeight)
                                    
        sHovType = GetHovercraftType
        If sHovType = "SEV" Then
           hSEVSidewalls = True
        ElseIf sHovType = "GEV" Then
            hGEVSkirt = True
        End If
        mvarhHoverAltitude = Round(CalcHHoverAltitude(mvarStaticLift), 2)
        
        mvarhMotiveThrust = CalcAMotiveThrust(mvarKeyChain, _
                        mvarhReservedHoverThrust, _
                        mvarTreatTiltRotorsAsPropellers, _
                        mvarTiltRotorForwardThrust, _
                        mvarPercentThrust, _
                        mvarAfterBurnersOn, _
                        mvarDatatype, _
                        mvartotalramjetthrust, _
                        mvarTotalRamJetThrustAB, _
                        mvarTotalTurboRamJetThrust, _
                        mvarTotalTurboramJetThrustAB)
        
        mvarhMotiveThrust = Round(mvarhMotiveThrust - mvarhReservedHoverThrust, 2)
        
        
        'Debug.Assert mvarhMotiveThrust >= 0
        mvarhDrag = CalcADrag(mvarPopTurretsExtended, _
                             mvarWheelsSkidsExtended, _
                             mbResponsive)
        mvarhDrag = Round(mvarhDrag, 2)
        
        mvarhTopSpeed = CalcHTopSpeed
        mvarhAcceleration = Round(CalcAAcceleration(mvarhMotiveThrust, m_VWeight), 2)
        mvarhManeuverability = Round(CalcAManeuverability(0, mbResponsive, m_VWeight), 2) 'todo: i believe its ok to just pass 0 for stall speed.  Double Check rules
        If hSEVSidewalls Then mvarhManeuverability = mvarhManeuverability + 0.25
        
        mvarhStability = Round(CalcAStability, 2): If hSEVSidewalls Then mvarhStability = mvarhStability + 1
        mvarhDeceleration = 4 * mvarhManeuverability 'simple rule

End Sub

Function CalcHTopSpeed() As Single
    Dim TempSpeed As Single
    'caclulate the top hover speed
    'essentially this is exactly the same as Aerial Top Speed except
    'it has a max speed of 300
    
    TempSpeed = CalcATopSpeed(mvarhDrag, _
                           mvarhMotiveThrust, _
                           mvarAfterBurnersOn, _
                           mvartotalramjetthrust, _
                           mvarTotalTurboRamJetThrust, _
                           mvarTotalRamJetThrustAB, _
                           mvarTotalTurboramJetThrustAB)
    
    'check for Max speed limits
    TempSpeed = CalcAMaxSpeed(TempSpeed, _
                              mvarKeyChain, _
                              mvarTreatTiltRotorsAsPropellers, _
                              g_sDC)
                              
    TempSpeed = Round(TempSpeed / 5, 0) * 5 'round to nearest 5mph
    
    'multiply by .8 if it has SEV sidewalls
    If hSEVSidewalls Then TempSpeed = TempSpeed * 0.8
    
    'make sure we havent exceeded max speed of 300
    TempSpeed = Minimum(TempSpeed, 300)
    
    CalcHTopSpeed = TempSpeed
End Function


Function CalcHHoverAltitude(ByVal StaticLift As Single) As Single
    Dim Lwt As Single
    Dim TempHover As Single
    Dim element As Object
    Dim sngVectoredThrustNeeded As Single
    Dim sngVectoredThrust As Single
    
    Lwt = m_VWeight
    
    'check for divide by zero
    If Lwt = 0 Then Exit Function
    
    sngVectoredThrustNeeded = Lwt - StaticLift
    If sngVectoredThrustNeeded > 0 Then
        mvarhReservedHoverThrust = GetUseableVectoredThrust(sngVectoredThrustNeeded, _
                                                            mvarKeyChain, _
                                                            mvarPercentThrust)
    Else
        mvarhReservedHoverThrust = 0
    End If
    
    If StaticLift + mvarhReservedHoverThrust < Lwt Then
        'if our static lift + vectored thrust is still not enough, we cant hover
        CalcHHoverAltitude = 0
        Advisory = Advisory & "LIFT TOO LOW. "
        Exit Function
    Else
        TempHover = (2 * (StaticLift + mvarhReservedHoverThrust) / Lwt)
    End If
    
    CalcHHoverAltitude = Minimum(TempHover, 6) 'make sure we dont exceed 6 foot hover max

End Function



Private Sub Class_Initialize()
    ReDim mvarKeyChain(1)
    mvarDatatype = PERFORMANCEHOVER
    
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







