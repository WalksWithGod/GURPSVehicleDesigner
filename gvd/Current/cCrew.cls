VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCrew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'todo:  We need to contemplate this hiearchy since we want to handle crew, crew station assignments,
'       station device assignments, computer/terminal device assignments, and crew shift profiles


' Implements cInode
' Implements cIDisplay
' Implements cIPersist

' Private m_Crew() as PERSON  ' eventually if we are going to want to store 'crew' as real individuals we can reference
'                             ' we'll need to keep a list of "crew" and "passenger" objects.
' Private m_Passengers() as PERSON

' Type SKILL
'   id as long
'   level as long
' End Type
' Type PERSON  '<-- needs to be an object
'   jobID as long ' this ID can be a position such as captain, mechanic, etc or a passenger
'   skills() as SKILL
' End Type

' ENUM JOBS
'   JOB_CAPTAIN = 1
' End ENUM
Private mvarOccupancy As String

Private mvarNumShifts As Long
Private mvarNumCaptains As Long
Private mvarNumOfficers As Long
Private mvarNumCrewStationOperators As Long
Private mvarNumWeaponLoaders As Long
Private mvarNumRowers As Long
Private mvarNumSailors As Long
Private mvarNumRiggers As Long
Private mvarNumFuelStokers As Long
Private mvarNumMechanics As Long
Private mvarNumServiceCrewmen As Long
Private mvarNumMedics As Long
Private mvarNumScientists As Long
Private mvarNumAuxiliaryVehicleCrew As Long
Private mvarNumStewards As Long
Private mvarNumLuxury As Long
Private mvarNumFirstClass As Long
Private mvarNumSecondClass As Long
Private mvarNumSteerage As Long
Private mvarTotalNumberCrewPassengers As Long
Private mvarUseRecommendedCrew As Boolean
Private mvarMilitaryVehicle As Boolean


Public Property Let MilitaryVehicle(ByVal vdata As Boolean)
    mvarMilitaryVehicle = vdata
End Property
Public Property Get MilitaryVehicle() As Boolean
    MilitaryVehicle = mvarMilitaryVehicle
End Property
Public Property Let UseRecommendedCrew(ByVal vdata As Boolean)
    mvarUseRecommendedCrew = vdata
End Property
Public Property Get UseRecommendedCrew() As Boolean
    UseRecommendedCrew = mvarUseRecommendedCrew
End Property
Public Property Let NumOfficers(ByVal vdata As Long)
    mvarNumOfficers = vdata
End Property
Public Property Get NumOfficers() As Long
    NumOfficers = mvarNumOfficers
End Property
Public Property Let NumShifts(ByVal vdata As Long)
    mvarNumShifts = vdata
End Property
Public Property Get NumShifts() As Long
     NumShifts = mvarNumShifts
End Property
Public Property Let NumCaptains(ByVal vdata As Long)
    mvarNumCaptains = vdata
End Property
Public Property Get NumCaptains() As Long
     NumCaptains = mvarNumCaptains
End Property
Public Property Let NumCrewStationOperators(ByVal vdata As Long)
    mvarNumCrewStationOperators = vdata
End Property
Public Property Get NumCrewStationOperators() As Long
     NumCrewStationOperators = mvarNumCrewStationOperators
End Property
Public Property Let NumWeaponLoaders(ByVal vdata As Long)
    mvarNumWeaponLoaders = vdata
End Property
Public Property Get NumWeaponLoaders() As Long
     NumWeaponLoaders = mvarNumWeaponLoaders
End Property
Public Property Let NumRowers(ByVal vdata As Long)
    mvarNumRowers = vdata
End Property
Public Property Get NumRowers() As Long
     NumRowers = mvarNumRowers
End Property
Public Property Let NumSailors(ByVal vdata As Long)
    mvarNumSailors = vdata
End Property
Public Property Get NumSailors() As Long
     NumSailors = mvarNumSailors
End Property
Public Property Let NumRiggers(ByVal vdata As Long)
    mvarNumRiggers = vdata
End Property
Public Property Get NumRiggers() As Long
     NumRiggers = mvarNumRiggers
End Property
Public Property Let NumFuelStokers(ByVal vdata As Long)
    mvarNumFuelStokers = vdata
End Property
Public Property Get NumFuelStokers() As Long
     NumFuelStokers = mvarNumFuelStokers
End Property
Public Property Let NumMechanics(ByVal vdata As Long)
    mvarNumMechanics = vdata
End Property
Public Property Get NumMechanics() As Long
     NumMechanics = mvarNumMechanics
End Property
Public Property Let NumServiceCrewmen(ByVal vdata As Long)
    mvarNumServiceCrewmen = vdata
End Property
Public Property Get NumServiceCrewmen() As Long
     NumServiceCrewmen = mvarNumServiceCrewmen
End Property
Public Property Let NumMedics(ByVal vdata As Long)
    mvarNumMedics = vdata
End Property
Public Property Get NumMedics() As Long
     NumMedics = mvarNumMedics
End Property
Public Property Let NumScientists(ByVal vdata As Long)
    mvarNumScientists = vdata
End Property
Public Property Get NumScientists() As Long
     NumScientists = mvarNumScientists
End Property
Public Property Let NumAuxiliaryVehicleCrew(ByVal vdata As Long)
    mvarNumAuxiliaryVehicleCrew = vdata
End Property
Public Property Get NumAuxiliaryVehicleCrew() As Long
     NumAuxiliaryVehicleCrew = mvarNumAuxiliaryVehicleCrew
End Property
Public Property Let NumStewards(ByVal vdata As Long)
    mvarNumStewards = vdata
End Property
Public Property Get NumStewards() As Long
     NumStewards = mvarNumStewards
End Property
Public Property Let NumLuxury(ByVal vdata As Long)
    mvarNumLuxury = vdata
End Property
Public Property Get NumLuxury() As Long
     NumLuxury = mvarNumLuxury
End Property
Public Property Let NumFirstClass(ByVal vdata As Long)
    mvarNumFirstClass = vdata
End Property
Public Property Get NumFirstClass() As Long
     NumFirstClass = mvarNumFirstClass
End Property
Public Property Let NumSecondClass(ByVal vdata As Long)
    mvarNumSecondClass = vdata
End Property
Public Property Get NumSecondClass() As Long
     NumSecondClass = mvarNumSecondClass
End Property
Public Property Let NumSteerage(ByVal vdata As Long)
    mvarNumSteerage = vdata
End Property
Public Property Get NumSteerage() As Long
     NumSteerage = mvarNumSteerage
End Property
Public Property Let Occupancy(ByVal vdata As String)
    mvarOccupancy = vdata
End Property
Public Property Get Occupancy() As String
    Occupancy = mvarOccupancy
End Property
Public Property Let TotalNumberCrewPassengers(ByVal vdata As Long)
    mvarTotalNumberCrewPassengers = vdata
End Property
Public Property Get TotalNumberCrewPassengers() As Long
     TotalNumberCrewPassengers = mvarTotalNumberCrewPassengers
End Property



Sub StatsUpdate()
    Dim NumCaptains As Long
    Dim NumOfficers As Long
    Dim NumCrewStationOperators As Long
    Dim NumLoaders As Long
    Dim NumRowers As Long
    Dim NumSailors As Long
    Dim NumGasRiggers As Long
    Dim NumFuelStokers As Long
    Dim NumMechanics As Long
    Dim NumServiceCrewmen As Long
    Dim NumMedics As Long
    Dim NumScientists As Long
    Dim NumAuxiliaryVehicleCrewmen As Long
    Dim NumStewards As Single 'depends on number of each class of passenger '6.02.2000 changed from Long to Single so it can be rounded up after NumShifts
    Dim NumLuxury As Long 'user defined
    Dim NumFirstClass As Long 'user defined
    Dim NumSecondClass As Long 'user defined
    Dim NumSteerage As Long 'user defined
    Dim TotalPassengers As Long
    Dim TotalCrewSize As Long
    Dim element As Object
    Dim SailDivisor As Integer
    Dim GasVolume As Double
    Dim NumShifts As Integer
    Dim LongTerm As Boolean
    Dim MilitaryVehicle As Boolean
    Dim i As Long
    Dim dType As Integer
    Dim nManeuverControlCount As Long
    Dim element2 As Object
    Dim bMechanical As Boolean
    Dim arrSubs() As String
    Dim j As Long
    Dim TotalMastHeight As Long
    Dim lngDiv As Long
    Dim dblTempPower As Double
    Dim bRequiresMedics As Boolean
    Dim lngTempMedics As Long
    
    
    'init totalcrewsize to 0
    TotalCrewSize = 0
    
    'determine if long or short term occupancy
    If mvarOccupancy = "long" Then
        LongTerm = True
    End If
    
    'determine if military vehicle
    If mvarMilitaryVehicle Then
        MilitaryVehicle = True
    End If
    
    'get the number of work shifts
    NumShifts = mvarNumShifts

    'get the total mast height
    arrSubs = Veh.KeyManager.GetCurrentSubAssembliesKeys
    If arrSubs(1) <> "" Then
        For j = 1 To UBound(arrSubs)
            If Veh.Components(arrSubs(j)).Datatype = Mast Then
                TotalMastHeight = TotalMastHeight + (Veh.Components(arrSubs(j)).Quantity * Veh.Components(arrSubs(j)).Height)
            End If
        Next
    End If
    'determine maneuver control modifiers which will be used for # of Sailors needed
    If gVehicleTL >= 7 Then
        For Each element2 In Veh.Components
            dType = element2.Datatype
            If (dType = MechanicalManeuverControl) Or (dType = MechanicalDivingControl) Then
                bMechanical = True
                Exit For
            ElseIf TypeOf element2 Is clsManeuverControl Then
                nManeuverControlCount = 1
                Exit For
            End If
        Next
    End If
    '//now we can search the vehicle and determine how many crew we need based on
    'what we find
    For Each element In Veh.Components
        dType = element.Datatype
        If (dType = RoomyCrewStation) Or (dType = NormalCrewStation) Or (dType = CrampedCrewStation) Or (dType = CycleCrewStation) Or (dType = HarnessCrewStation) Then
            NumCrewStationOperators = NumCrewStationOperators + element.Quantity
        ElseIf (dType = BattlesuitSystem) Or (dType = FormFittingBattleSuitSystem) Then
            NumCrewStationOperators = NumCrewStationOperators + 1
        ElseIf (TypeOf element Is clsWeaponStoneBoltThrower) Or (TypeOf element Is clsWeaponGun) Or (TypeOf element Is clsWeaponLauncher) Then
                NumLoaders = NumLoaders + element.Loaders
        ElseIf dType = RowingPositions Then
                NumRowers = NumRowers + element.Quantity
        ElseIf (dType = HotAir) Or (dType = Hydrogen) Or (dType = Helium) Then
            GasVolume = GasVolume + element.Volume
        ElseIf TypeOf element Is clsLabandWorkshop Then
            NumScientists = NumScientists + element.Quantity
        ElseIf (dType = SquareRig) Or (dType = ForeandAftRig) Or (dType = FullRig) Or (dType = AerialSail) Or (dType = AerialSailForeAftRig) Then
            If (gVehicleTL >= 7) And (bMechanical = False) And (nManeuverControlCount <> 0) Then
                NumSailors = 0
            Else
                If dType = ForeandAftRig Then SailDivisor = 5 Else SailDivisor = 4
                NumSailors = Sqr(TotalMastHeight) / SailDivisor
            End If
        End If
    Next
    
    'get the number of mechanics
    For Each element2 In Veh.Components
        Select Case element2.Datatype
        
            Case TrackedDrivetrain, LegDrivetrain, FlexibodyDrivetrain, _
                 WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, _
                 PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, _
                 Hydrojet, MHDTunnel, DuctedFan, AerialPropeller, _
                 LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                 OptimizedFusion, AntimatterThermal, AntimatterPion, _
                 StandardThruster, SuperThruster, MegaThruster, _
                 OrnithopterDrivetrain, _
                MagLevLifter, TeleportationDrive, JumpDrive, WarpDrive, _
                ContraGravGenerator, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain
        
                dblTempPower = dblTempPower + element2.PowerReqt
                
            'Case Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam
                '//jet engines dont have a power reqt so i guess they dont apply since the
                '//rules on page 75 for mechanics only mention Power Reqt systems
                
            Case Hyperdrive
                dblTempPower = dblTempPower + element2.SustainedPower
        End Select
    Next
    
    '//get the divisor
    Select Case gVehicleTL
        Case Is <= 5
            lngDiv = 250
        Case 6
            lngDiv = 500
        Case 7
            lngDiv = 1000
        Case 8
            lngDiv = 5000
        Case 9
            lngDiv = 50000
        Case 10
            lngDiv = 250000
        Case 11
            lngDiv = 500000
        Case 12, 13
            lngDiv = 1000000
        Case Is >= 14
            lngDiv = 5000000
    End Select
    NumMechanics = Int(Sqr(dblTempPower / lngDiv))
    
    'get number of stewards
    NumLuxury = mvarNumLuxury
    NumFirstClass = mvarNumFirstClass
    NumSecondClass = mvarNumSecondClass
    NumSteerage = mvarNumSteerage
    TotalPassengers = NumSteerage + NumSecondClass + NumFirstClass + NumLuxury  ' MPJ 06.02.2000 Medics will now take into account passenger count too
         
    NumStewards = NumSteerage / 100
    NumStewards = NumStewards + NumSecondClass / 50
    NumStewards = NumStewards + NumFirstClass / 20
    NumStewards = NumStewards + NumLuxury / 4
   
    
    'get number of gasriggers required (Linked to shifts)
    NumGasRiggers = (GasVolume / 200000)
    If GasVolume < 200000 Then NumGasRiggers = 0   'MPJ 05.19.2000
    
    'get number of medics needed (not linked to Shifts since it dependant on totalcrewsize)
    For Each element In Veh.Components     'MPJ 05.18.2000 - Updated this entire medics crew search
        dType = element.Datatype
        If (dType = StretcherPallet) Or (dType = DiagnosisTable) Then
            bRequiresMedics = True
        ElseIf dType = OperatingRoom Then
            NumMedics = NumMedics + 1
        End If
    Next
    ' if we require at least one medic but have no operating tables, set NumMedics to 1
    If NumMedics < 1 And (bRequiresMedics) Then NumMedics = 1
    ' make sure we have at least as many doctors needed for total crew size
    If LongTerm Then
        If gVehicleTL <= 4 Then
            lngTempMedics = TotalCrewSize + TotalPassengers / 50
        ElseIf gVehicleTL <= 7 Then
            lngTempMedics = TotalCrewSize + TotalPassengers / 100
        Else
            lngTempMedics = TotalCrewSize + TotalPassengers / 200
        End If
    End If
    If lngTempMedics > NumMedics Then NumMedics = lngTempMedics
    
    
    'For all of the above ones, get final number based on number of Work Shifts
    NumMedics = NumMedics * NumShifts
    NumScientists = NumScientists * NumShifts
    NumGasRiggers = NumGasRiggers * NumShifts
    NumCrewStationOperators = NumCrewStationOperators * NumShifts
    NumRowers = NumRowers * NumShifts
    NumLoaders = NumLoaders * NumShifts
    NumSailors = NumSailors * NumShifts
    NumMechanics = NumMechanics * NumShifts
    NumFuelStokers = NumFuelStokers * NumShifts
    NumStewards = Round(NumStewards * NumShifts, 0)
    

    'NumAuxiliaryVehicleCrewmen  'TODO: Right now it doesnt calc Auxiliary Crew.  This should
    'be a user defined choice anyway since its beyond the scope of GVD to calc crew of
    ' other vehicles loaded onto this vehicle.
    NumAuxiliaryVehicleCrewmen = mvarNumAuxiliaryVehicleCrew
               
    'add them all up so far.  MPJ 06.02.2000 Total Passengers Now added Here instead of further below. (Note previous version appeared to add them here too, but the values were never set for them)
    TotalCrewSize = TotalPassengers + NumCaptains + NumOfficers + NumCrewStationOperators + NumLoaders + NumRowers + NumSailors + NumGasRiggers + NumFuelStokers + NumMechanics + NumServiceCrewmen + NumScientists + NumAuxiliaryVehicleCrewmen + NumStewards
    
    'get number of service crewmen needed (not linked to shifts)
    If LongTerm Then
        NumServiceCrewmen = TotalCrewSize / 20
    End If
    TotalCrewSize = TotalCrewSize + NumServiceCrewmen
    
   
    'update the total crew size
    TotalCrewSize = TotalCrewSize + NumMedics
    
    'get number of captains/commanders (Linked to Shifts)
    If (MilitaryVehicle) And (TotalCrewSize > 1) Then
        NumCaptains = 1 * NumShifts
    ElseIf TotalCrewSize > 12 Then
        NumCaptains = 1 * NumShifts
    End If
    
    'get number of officers needed (not linked to shifts since its dependant on TotalCrewSize)
    NumOfficers = TotalCrewSize / 8
    
    
    TotalCrewSize = TotalCrewSize + NumOfficers
    
    'save it all
    mvarNumCaptains = NumCaptains
    mvarNumOfficers = NumOfficers
    mvarNumCrewStationOperators = NumCrewStationOperators
    mvarNumWeaponLoaders = NumLoaders
    mvarNumRowers = NumRowers
    mvarNumSailors = NumSailors
    mvarNumRiggers = NumGasRiggers
    mvarNumMechanics = NumMechanics
    mvarNumFuelStokers = NumFuelStokers
    mvarNumServiceCrewmen = NumServiceCrewmen
    mvarNumMedics = NumMedics
    mvarNumScientists = NumScientists
    mvarNumAuxiliaryVehicleCrew = NumAuxiliaryVehicleCrewmen
    mvarNumStewards = NumStewards
        
    'note that the passengers arent saved because the program
    'cant calculate how many passengers there should be.  User must enter.
    mvarTotalNumberCrewPassengers = TotalCrewSize

End Sub


Private Sub Class_Initialize()
    mvarNumShifts = 1
    mvarUseRecommendedCrew = True
    mvarOccupancy = "long"
End Sub
