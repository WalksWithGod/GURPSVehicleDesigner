VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerformanceSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsManeuverability As Single
Private mvarsTurnAroundTime As Single
Private mvarsMotiveThrust As Single
Private mvarsAccelerationG As Single
Private mvarsAccelerationMPH As Single
Private mvarsHyperSpeed As Single
Private mvarsWarpSpeed As Single
Private mvarsJumpDriveable As Boolean
Private mvarsTeleportationDriveable As Boolean

Private mvarKey As String
Private mvarParent As String
Private mvarDatatype As Long
Private mvarDescription As String

Private mbResponsive As Boolean

Private mvarHardPointsOn As Boolean

        
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

Public Property Let HardPointsOn(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HardPointsOn = 5
    mvarHardPointsOn = vdata
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


Public Property Get HardPointsOn() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HardPointsOn
    HardPointsOn = mvarHardPointsOn
End Property


Public Property Let sMotiveThrust(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sMotiveThrust = 5
    mvarsMotiveThrust = vdata
End Property

Public Property Get sMotiveThrust() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sMotiveThrust
    sMotiveThrust = mvarsMotiveThrust
End Property



Public Property Let sAccelerationG(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sAccelerationG = 5
    mvarsAccelerationG = vdata
End Property

Public Property Get sAccelerationG() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sAccelerationG
    sAccelerationG = mvarsAccelerationG
End Property


Public Property Let sAccelerationMPH(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sAccelerationMPH = 5
    mvarsAccelerationMPH = vdata
End Property

Public Property Get sAccelerationMPH() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sAccelerationMPH
    sAccelerationMPH = mvarsAccelerationMPH
End Property


Public Property Let sTurnAroundTime(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTurnAroundTime = 5
    mvarsTurnAroundTime = vdata
End Property

Public Property Get sTurnAroundTime() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTurnAroundTime
    sTurnAroundTime = mvarsTurnAroundTime
End Property


Public Property Let sHyperSpeed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sHyperSpeed = 5
    mvarsHyperSpeed = vdata
End Property

Public Property Get sHyperSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sHyperSpeed
    sHyperSpeed = mvarsHyperSpeed
End Property

Public Property Let sWarpSpeed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sWarpSpeed = 5
    mvarsWarpSpeed = vdata
End Property

Public Property Get sWarpSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sWarpSpeed
    sWarpSpeed = mvarsWarpSpeed
End Property

Public Property Let sManeuverability(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sManeuverability = 5
    mvarsManeuverability = vdata
End Property

Public Property Get sManeuverability() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sManeuverability
    sManeuverability = mvarsManeuverability
End Property

Public Property Let sTeleportationDriveable(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTeleportationDriveable = 5
    mvarsTeleportationDriveable = vdata
End Property

Public Property Get sTeleportationDriveable() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTeleportationDriveable
    sTeleportationDriveable = mvarsTeleportationDriveable
End Property

Public Property Let sJumpDriveable(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sJumpDriveable = 5
    mvarsJumpDriveable = vdata
End Property

Public Property Get sJumpDriveable() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sJumpDriveable
    sJumpDriveable = mvarsJumpDriveable
End Property


Public Sub CalcPerformance()


    mvarAdvisory = ""


    'determine if vehicle has responsive structure
    mbResponsive = VehicleHasResponsiveStruct

    Call GetVehicleWeight(PERFORMANCESPACE, _
                            mvarPercentAuxVehicleWeight, _
                            mvarPercentCargoWeight, _
                            mvarPercentAmmunitionWeight, _
                            mvarPercentHardpointWeight, _
                            mvarPercentFuelWeight, _
                            mvarPercentProvisionWeight, _
                            m_VWeight, m_VMass)
                            
    


        mvarsMotiveThrust = Round(CalcSMotiveThrust, 2)
        mvarsAccelerationG = Round(CalcSAccelerationG, 4)
        mvarsAccelerationMPH = Round(CalcSAccelerationMPH, 4)
        mvarsManeuverability = Round(mvarsAccelerationG, 2) 'this equals acceleration
        mvarsTurnAroundTime = Round(CalcSTurnAroundTime, 2)
        mvarsHyperSpeed = Round(CalcSHyperspeed, 2)
        mvarsWarpSpeed = Round(CalcSWarpSpeed, 2)
        mvarsJumpDriveable = CalcJumpDriveable
        mvarsTeleportationDriveable = CalcTeleportationDriveable
 
End Sub

Function CalcSTurnAroundTime() As Single
    CalcSTurnAroundTime = Maximum(Veh.Stats.SizeModifier * 10, 1)
End Function

Function CalcSMotiveThrust()
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

        Case LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                OptimizedFusion, AntimatterThermal, AntimatterPion, _
                StandardThruster, SuperThruster, MegaThruster, SolidRocketEngine, OrionEngine
                
            TempThrust = TempThrust + Veh.Components(sKey).MotiveThrust * mvarPercentThrust
        
        Case lightSail
            TempThrust = TempThrust + Veh.Components(sKey).Thrust * mvarPercentThrust
    End Select
    Next

CalcSMotiveThrust = TempThrust


End Function

Function CalcSAccelerationG() As Single
If m_VWeight <= 0 Then
    CalcSAccelerationG = 0
Else
    CalcSAccelerationG = mvarsMotiveThrust / m_VWeight
End If
End Function

Function CalcSAccelerationMPH() As Single
Const GravitiesToMPH = 21.9

CalcSAccelerationMPH = mvarsAccelerationG * GravitiesToMPH
End Function

Function CalcSHyperspeed() As Single
'Note: this only uses Hyperspeed drives that are added to the keychain since
'other hyperdrives can be carried as just cargo

'Note: the errate allows speeds higher than the flat .2 per day.  The book incorrectly states
' a max of .2 parsecs a day

Const HyperSpeedConstant = 0.2 '.2 parsecs
Dim Lmass As Single
Dim i As Long
Dim HP As Single
Dim sKey As String
Dim dType As Integer

If mvarKeyChain(1) = "" Then Exit Function 'exit if there are no propulsion systems in the keychain

Lmass = m_VMass

For i = 1 To UBound(mvarKeyChain)
    sKey = mvarKeyChain(i)
    dType = Veh.Components(sKey).Datatype
    If dType = Hyperdrive Then
        HP = HP + Veh.Components(sKey).capacity
    End If
Next

CalcSHyperspeed = HP / Lmass * HyperSpeedConstant
End Function

Function CalcSWarpSpeed() As Single
'Note: this only uses Warp drives that are added to the keychain since
'other Warp drives can be carried as just cargo
Const WarpSpeedConstant = 1 '1 parsec per day
Dim Lmass As Single
Dim i As Long
Dim sKey As String
Dim dType As Integer
Dim WF As Single

If mvarKeyChain(1) = "" Then Exit Function 'exit if there are no propulsion systems in the keychain

Lmass = m_VMass

For i = 1 To UBound(mvarKeyChain)
    sKey = mvarKeyChain(i)
    dType = Veh.Components(sKey).Datatype
    If dType = WarpDrive Then
        WF = WF + Veh.Components(sKey).capacity
    End If
Next

CalcSWarpSpeed = WF / Lmass * WarpSpeedConstant

End Function

Function CalcJumpDriveable() As Boolean
'determine if the ship is capable of Jump travel via jump drives
'Note: this only uses Jump drives that are added to the keychain since
'other Jump can be carried as just cargo
Dim TempRating As Single
Dim i As Long
Dim sKey As String
Dim dType As Integer

If mvarKeyChain(1) = "" Then Exit Function 'exit if there are no propulsion systems in the keychain

For i = 1 To UBound(mvarKeyChain)
    sKey = mvarKeyChain(i)
    dType = Veh.Components(sKey).Datatype
    If dType = JumpDrive Then
        TempRating = TempRating + Veh.Components(sKey).capacity
    End If
Next

If TempRating >= m_VMass Then
    CalcJumpDriveable = True
Else
    CalcJumpDriveable = False
End If

End Function

Function CalcTeleportationDriveable() As Boolean
'determine if the ship is capable of Teleportation travel
'Note: this only uses Teleportation drives that are added to the keychain since
'other Teleportation drives can be carried as just cargo
Dim i As Long
Dim sKey As String
Dim dType As Integer

If mvarKeyChain(1) = "" Then Exit Function 'exit if there are no propulsion systems in the keychain

For i = 1 To UBound(mvarKeyChain)
    sKey = mvarKeyChain(i)
    dType = Veh.Components(sKey).Datatype
    If dType = TeleportationDrive Then
        CalcTeleportationDriveable = True
        Exit Function
    End If
Next
    
CalcTeleportationDriveable = False
    
End Function


Private Sub Class_Initialize()
    ReDim mvarKeyChain(1)
    mvarDatatype = PERFORMANCESPACE
    
    'mvarAfterBurnersOn = True ' not used in space
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



