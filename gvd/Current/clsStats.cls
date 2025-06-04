VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//////////////////////////////////////////////////////
' Total Vehicle Stats
Private mvarFloatationRating As Single
Private mvarTotalPrice As Single  'todo: rename these to just price, surface, weight, etc to be consistant. In other words, we dont need "total" in each variables name
Private mvarTotalSurfaceArea As Single
Private mvarStructuralSurfaceArea As Single
Private mvarStructuralCost As Single
Private mvarStructuralWeight As Single
Private mvarEmptyWeight As Single
Private mvarHLoadedWeight As Single
Private mvarLoadedWeight As Single
Private mvarLoadedMass As Single
Private mvarHLoadedMass As Single
Private mvarTotalVolume As Single
Private mvarSubmergedWeight As Single
Private mvarSubmergedMass As Single
Private mvarHSubmergedWeight As Single ' with hardpoints loaded
Private mvarHSubmergedMass As Single
Private mvarUsualInternalPayload As Single
Private mvarStructuralHealth As Long 'always with hardpoints loaded
Private mvarSizeModifier As Integer
Private mvarTotalContinuousPower As Single 'todo: these may need to be stats of the power profiles since it depends on what devices the user is running
Private mvarTotalContinuousPowerConsumption As Single
Private mvarTotalStoredPower As Single
Private mvarTotalStoredPowerConsumption As Single
Private mvarTotalHardPointConnections As Long

' option settings which should probably be moved to a dialog, but still loaded
' via a class (e.g. cOptions duh?(
Private mvarBattleSuit As String  ' setting type "none", "form-fitting" etc
Private mvarQuality As String
Private mvarPerPersonWeight As Single
Private mvarPerCargoWeight As Single
Private mvarRecommendedPayload As Single
Private mvarRecommendedAccessSpace As Single
Private mvarAccessSpaceVolumeMod As Single
Private mvarUseHardpointMountedWeights As Boolean
Private mvarUseSurfaceAreaTable As Boolean

Public Property Let BattleSuit(ByVal vdata As String)
    mvarBattleSuit = vdata
End Property
Public Property Get BattleSuit() As String
    BattleSuit = mvarBattleSuit
End Property
Public Property Let Quality(ByVal vdata As String)
    mvarQuality = vdata
End Property
Public Property Get Quality() As String
    Quality = mvarQuality
End Property
Public Property Let UseSurfaceAreaTable(ByVal vdata As Boolean)
    mvarUseSurfaceAreaTable = vdata
End Property
Public Property Get UseSurfaceAreaTable() As Boolean
    UseSurfaceAreaTable = mvarUseSurfaceAreaTable
End Property
Public Property Let UseHardpointMountedWeights(ByVal vdata As Boolean)
    mvarUseHardpointMountedWeights = vdata
End Property
Public Property Get UseHardpointMountedWeights() As Boolean
    UseHardpointMountedWeights = mvarUseHardpointMountedWeights
End Property
Public Property Let RecommendedAccessSpace(ByVal vdata As Boolean)
    mvarRecommendedAccessSpace = vdata
End Property
Public Property Get RecommendedAccessSpace() As Boolean
    RecommendedAccessSpace = mvarRecommendedAccessSpace
End Property
Public Property Let RecommendedPayload(ByVal vdata As Boolean)
    mvarRecommendedPayload = vdata
End Property
Public Property Get RecommendedPayload() As Boolean
    RecommendedPayload = mvarRecommendedPayload
End Property
Public Property Let PerPersonWeight(ByVal vdata As Single)
    mvarPerPersonWeight = vdata
End Property
Public Property Get PerPersonWeight() As Single
     PerPersonWeight = mvarPerPersonWeight
End Property
Public Property Let PerCargoWeight(ByVal vdata As Single)
    mvarPerCargoWeight = vdata
End Property
Public Property Get PerCargoWeight() As Single
     PerCargoWeight = mvarPerCargoWeight
End Property
Public Property Let AccessSpaceVolumeMod(ByVal vdata As Single)
    mvarAccessSpaceVolumeMod = vdata
End Property
Public Property Get AccessSpaceVolumeMod() As Single
     AccessSpaceVolumeMod = mvarAccessSpaceVolumeMod
End Property


Public Property Let TotalHardPointConnections(ByVal vdata As Long)
    mvarTotalHardPointConnections = vdata
End Property
Public Property Get TotalHardPointConnections() As Long
    TotalHardPointConnections = mvarTotalHardPointConnections
End Property

Public Property Let UsualInternalPayload(ByVal vdata As Double)
    mvarUsualInternalPayload = vdata
End Property
Public Property Get UsualInternalPayload() As Double
    UsualInternalPayload = mvarUsualInternalPayload
End Property
Public Property Let StructuralSurfaceArea(ByVal vdata As Double)
    mvarStructuralSurfaceArea = vdata
End Property
Public Property Get StructuralSurfaceArea() As Double
    StructuralSurfaceArea = mvarStructuralSurfaceArea
End Property
Public Property Let totalSurfaceArea(ByVal vdata As Double)
    mvarTotalSurfaceArea = vdata
End Property
Public Property Get totalSurfaceArea() As Double
    totalSurfaceArea = mvarTotalSurfaceArea
End Property
Public Property Let StructuralCost(ByVal vdata As Double)
    mvarStructuralCost = vdata
End Property
Public Property Get StructuralCost() As Double
    StructuralCost = mvarStructuralCost
End Property
Public Property Let StructuralWeight(ByVal vdata As Double)
    mvarStructuralWeight = vdata
End Property
Public Property Get StructuralWeight() As Double
    StructuralWeight = mvarStructuralWeight
End Property
Public Property Let SizeModifier(ByVal vdata As Integer)
    mvarSizeModifier = vdata
End Property
Public Property Get SizeModifier() As Integer
    SizeModifier = mvarSizeModifier
End Property
Public Property Let TotalPrice(ByVal vdata As Single)
    mvarTotalPrice = vdata
End Property
Public Property Get TotalPrice() As Single
    TotalPrice = mvarTotalPrice
End Property
Public Property Let StructuralHealth(ByVal vdata As Long)
    mvarStructuralHealth = vdata
End Property
Public Property Get StructuralHealth() As Long
    StructuralHealth = mvarStructuralHealth
End Property
Public Property Let EmptyWeight(ByVal vdata As Double)
    mvarEmptyWeight = vdata
End Property
Public Property Get EmptyWeight() As Double
    EmptyWeight = mvarEmptyWeight
End Property
Public Property Let HSubmergedWeight(ByVal vdata As Double)
    mvarHSubmergedWeight = vdata
End Property
Public Property Get HSubmergedWeight() As Double
    HSubmergedWeight = mvarHSubmergedWeight
End Property
Public Property Let SubmergedWeight(ByVal vdata As Double)
    mvarSubmergedWeight = vdata
End Property
Public Property Get SubmergedWeight() As Double
    SubmergedWeight = mvarSubmergedWeight
End Property
Public Property Let HSubmergedMass(ByVal vdata As Double)
    mvarHSubmergedMass = vdata
End Property
Public Property Get HSubmergedMass() As Double
    HSubmergedMass = mvarHSubmergedMass
End Property
Public Property Let SubmergedMass(ByVal vdata As Double)
    mvarSubmergedMass = vdata
End Property
Public Property Get SubmergedMass() As Double
    SubmergedMass = mvarSubmergedMass
End Property
Public Property Let FloatationRating(ByVal vdata As Single)
    mvarFloatationRating = vdata
End Property
Public Property Get FloatationRating() As Single
    FloatationRating = mvarFloatationRating
End Property
Public Property Let HLoadedWeight(ByVal vdata As Double)
    mvarHLoadedWeight = vdata
End Property
Public Property Get HLoadedWeight() As Double
    HLoadedWeight = mvarHLoadedWeight
End Property
Public Property Let HLoadedMass(ByVal vdata As Double)
    mvarHLoadedMass = vdata
End Property
Public Property Get HLoadedMass() As Double
    HLoadedMass = mvarHLoadedMass
End Property
Public Property Let LoadedWeight(ByVal vdata As Double)
    mvarLoadedWeight = vdata
End Property
Public Property Get TotalVolume() As Double
    TotalVolume = mvarTotalVolume
End Property
Public Property Let TotalVolume(ByVal vdata As Double)
    mvarTotalVolume = vdata
End Property
Public Property Get LoadedWeight() As Double
    LoadedWeight = mvarLoadedWeight
End Property
Public Property Let LoadedMass(ByVal vdata As Double)
    mvarLoadedMass = vdata
End Property
Public Property Get LoadedMass() As Double
    LoadedMass = mvarLoadedMass
End Property
Public Property Let TotalContinuousPower(ByVal vdata As Single)
    mvarTotalContinuousPower = vdata
End Property
Public Property Get TotalContinuousPower() As Single
    TotalContinuousPower = mvarTotalContinuousPower
End Property
Public Property Let TotalStoredPower(ByVal vdata As Single)
    mvarTotalStoredPower = vdata
End Property
Public Property Get TotalStoredPower() As Single
    TotalStoredPower = mvarTotalStoredPower
End Property
Public Property Let TotalContinuousPowerConsumption(ByVal vdata As Single)
    mvarTotalContinuousPowerConsumption = vdata
End Property
Public Property Get TotalContinuousPowerConsumption() As Single
    TotalContinuousPowerConsumption = mvarTotalContinuousPowerConsumption
End Property
Public Property Let TotalStoredPowerConsumption(ByVal vdata As Single)
    mvarTotalStoredPowerConsumption = vdata
End Property
Public Property Get TotalStoredPowerConsumption() As Single
    TotalStoredPowerConsumption = mvarTotalStoredPowerConsumption
End Property


'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////
' Statistics Calculation Functions
'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////
Sub Update()
    ' see pages 25 - 26 of Vehicles
    Dim element As Object
    Dim lngStart As Long
    Dim lngStop As Long
    On Error Resume Next
    
    '//first, recalc every single component's stats
    For Each element In Veh.Components
        element.StatsUpdate
    Next
    'calc total structural stats and have each sub update
    'its own surface area, volume and weight stats
    'this routine saves the Vehicle's Structural Surface area and Weight
    Call Me.CalcStructuralStats
    
    'calc weight and cost for all options and surface features
    If Veh.crew.UseRecommendedCrew Then
        Veh.crew.StatsUpdate
    End If
    
    Veh.Options.CalcOptionsWeightandCost
    Veh.surface.Update
    
    DoEvents
    'calc vehicles total empty weight, loadedweight, etc
    Call CalcWeight
    
    
    'calc total volume of vehicle
    mvarTotalVolume = Me.CalcTotalVolume
    mvarTotalPrice = Me.CalcPrice ' calc final price of vehicle
    mvarSizeModifier = Me.CalcSizeModifier(mvarTotalVolume)
    mvarTotalContinuousPower = Me.CalcTotalGeneratedPower
    mvarTotalContinuousPowerConsumption = Me.CalcTotalContinuousPowerConsumption
    
    ' If vehicle is submersible, calculate the submerged weight and mass
    If surface.Submersible Then
        mvarSubmergedWeight = Me.CalcSubmergedWeight(mvarLoadedWeight, mvarTotalVolume)
        mvarSubmergedMass = mvarSubmergedWeight / 2000
    Else
        mvarSubmergedWeight = 0
        mvarSubmergedMass = 0
    End If
    
    mvarStructuralHealth = Me.CalcHealth(Veh.Components(BODY_KEY).HitPoints, mvarHLoadedWeight)
    
    ' re-calc vehicle performance figures
    For Each element In Veh.PerformanceProfiles
        element.CalcPerformance
    Next
        
End Sub

Public Function CalcFloatationRating() As Single
'todo: single or double return value?
Dim Floatation As Single ' the floatation multiplier

' Calculate the floatation rating (must be done after volume)

' to have a floatation rating, it must have floatation enabled.
' Note that subermisble hulls require the floatation option
If (Veh.surface.Floatation) Then
    Floatation = 62.5
    If Not (Veh.surface.Submersible) Then
        ' submersible hulls will have 62.5 rating irregardless of line type
        ' so if its not a submersible hull, we check the lines to reduce the rating
        Select Case Veh.surface.HydrodynamicLines
            Case "none"
                Floatation = 62.5
            Case "mediocre"
                Floatation = 57
            Case "average"
                Floatation = 52
            Case "submarine"
                Floatation = 62.5
            Case "fine"
                Floatation = 48
            Case "very fine"
                Floatation = 45
        End Select
    End If
Else
    Floatation = 0
End If

'todo: the floatation rating calc below is correct for a floating ship, but not for a submersible one! since
'      those require the volume of all supers and turrets as well as body!!!
'      This stat should be MOVED to the stats class anyways!  Its not a body stat!

mvarFloatationRating = Round(Floatation * mvarVolume, 2)
End Function
Public Sub CalcStructuralStats()
    Dim temparea As Single
    Dim TempTotalArea As Single
    Dim TempCost As Single
    Dim TempWeight As Single
    Dim element As Object
    ' Calculate the Total Structural and Total Surface Area Weight of the Vehicle
    'NOTE: This is NOT supposed to use components, ONLY other SUB ASSEMBLIES
    'todo: optimize to only check on "subassemblies" or depending on changes to code
    ' architecture, "container" type objects sans "group object"
    For Each element In Veh.Components
        Select Case element.Datatype
            Case Body, Skid, Wheel, Track, Leg, Wing, AutogyroRotor, _
                TTRotor, CARotor, MMRotor, Hydrofoil, Hovercraft, Superstructure, _
                Pod, Turret, Popturret, Arm
                'element.StatsUpdate 'Note this call must be made
                '(NOTE: The above is commented out since im now updating ALL STATS
                'before hand in CalcStats)
                'since Volumes for most of the above Subassemblies takes
                'the body's Volume into account.  Note that the Body must
                'statupdate FIRST as a result which is ensured here
                
                temparea = element.SurfaceArea + temparea
                TempWeight = element.Weight + TempWeight
                TempCost = element.Cost + TempCost
            Case Mast, Gasbag, OpenMount
                TempTotalArea = element.SurfaceArea + TempTotalArea
        End Select
    Next
    
    
    mvarStructuralSurfaceArea = temparea
    mvarTotalSurfaceArea = temparea + TempTotalArea
    mvarStructuralCost = TempCost
    mvarStructuralWeight = TempWeight

End Sub

Sub CalcWeight()
    Dim FuelWeight As Single
    Dim AmmoWeight As Single
    Dim ProvisionsWeight As Single
    Dim GunCarriagesWeight As Single
    Dim AuxVehiclesWeight As Single
    Dim HardPointWeight As Single
    Dim CargoWeight As Single
    Dim lngObjectsConnectedtoHardpoints As Long
    Dim ComponentsWeight As Single
    
    Dim element As Object ' any object in the collection
    ' add up the weight of all componenets, structures, features except for Fuel, Ammunition
    ' provisions, any carried vehicles, and TL5- guns that require carriages.
    ' the resulting sum is the Empty Weight
    On Error GoTo errorhandler
    
    For Each element In Veh.Components
        
        If TypeOf element Is clsBody Then
            With element
                ComponentsWeight = ComponentsWeight + element.Weight
            End With
        
        'ElseIf TypeOf element Is clsWeaponLink Then '07/09/02 MPJ OBSOLETE. Weaponlinks no longer stored in Components collection
    
        ElseIf TypeOf element Is clsCargo Then
            CargoWeight = CargoWeight + element.CargoWeight '//we get cargo weight in the Usual internal Payload
            ComponentsWeight = ComponentsWeight + element.Weight
        
         ElseIf TypeOf element Is clsHardPoint Then
                If Veh.Options.UseHardpointMountedWeights = False Then
                    HardPointWeight = HardPointWeight + (element.Quantity * element.LoadCapacity)
                End If
                
                ComponentsWeight = ComponentsWeight + element.Weight
                '//note equipment pods can only attach to hardpoints so thats where
                '//we get their weights
        ElseIf TypeOf element Is clsSoftware Then
    
        ElseIf TypeOf element Is clsLiftingGas Then
       
        ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsHardPoint Then
        
            If Veh.Options.UseHardpointMountedWeights Then
                HardPointWeight = HardPointWeight + element.Weight
            End If
            '//count up how many objects are actually attached to hardpoints so
            '//we can use these in drag calculations later
            If Veh.Components(element.LogicalParent).Datatype = HardPoint Then
                lngObjectsConnectedtoHardpoints = lngObjectsConnectedtoHardpoints + 1
            End If
        ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsEquipmentPod Then
            If Veh.Options.UseHardpointMountedWeights Then
                HardPointWeight = HardPointWeight + element.Weight
            End If
            '//count up how many objects are actually attached to hardpoints so
            '//we can use these in drag calculations later
            If Veh.Components(element.LogicalParent).Datatype = HardPoint Then
                lngObjectsConnectedtoHardpoints = lngObjectsConnectedtoHardpoints + 1
            End If
        ElseIf TypeOf element Is clsFuelTank Then
            FuelWeight = FuelWeight + element.FuelWeight
            ComponentsWeight = ComponentsWeight + element.Weight
            
        ElseIf TypeOf element Is clsProvisions Then
            ProvisionsWeight = ProvisionsWeight + element.Weight
    
        ElseIf TypeOf element Is clsVehicleStorage Then
            AuxVehiclesWeight = AuxVehiclesWeight + element.CraftWeight
            ComponentsWeight = ComponentsWeight + element.Weight
    
        ElseIf TypeOf element Is clsWeaponAmmunition Then 'check for ammunition
            AmmoWeight = AmmoWeight + element.Weight
            
        ElseIf TypeOf element Is clsWeaponGun Then  'check for guns with carriages
            If element.Carriage Then
                GunCarriagesWeight = GunCarriagesWeight + element.Weight
            Else
                ComponentsWeight = ComponentsWeight + element.Weight
            End If
        
        Else
            ComponentsWeight = ComponentsWeight + element.Weight
        End If
    Next
    
    'save these overall vehicle stats
    
    mvarUsualInternalPayload = (Veh.Options.PerPersonWeight * Veh.crew.TotalNumberCrewPassengers) + _
                            CargoWeight
    mvarEmptyWeight = ComponentsWeight + Veh.Options.OptionsWeight
    mvarLoadedWeight = mvarEmptyWeight + mvarUsualInternalPayload + _
                    Veh.Options.RollStabilizersWeight + FuelWeight + AmmoWeight + _
                    ProvisionsWeight + AuxVehiclesWeight + _
                    Veh.Options.TotalVariableSweepWeight + _
                    Veh.Options.TotalFoldingWingWeight + _
                    Veh.Options.TotalCompartmentalizationWeight
    mvarLoadedMass = mvarLoadedWeight / 2000
    mvarHLoadedWeight = mvarLoadedWeight + HardPointWeight
    mvarHLoadedMass = mvarHLoadedWeight / 2000
    mvarTotalHardPointConnections = lngObjectsConnectedtoHardpoints
    
    Exit Sub

errorhandler:
    Debug.Print "CalcWeight: " & err.Description
End Sub

Function CalcSubmergedWeight(Lweight As Single, TVolume As Single) As Single
    Dim TempSweight As Single
    Const Multiplier = 62.5
    
    TempSweight = Multiplier * TVolume
    
    If TempSweight < Lweight Then
        TempSweight = Lweight
    End If
    
    CalcSubmergedWeight = TempSweight
    'TODO: Do i need to do a submerged weight with hardpoints?
    'i've already got a HSubmergedWeight and HSubmergedMass properties in the clsBody
End Function

Function CalcTotalVolume() As Single

    Dim TempVolume As Single ' holds the volume in progress
    Dim element As Object
    
    For Each element In Veh.Components
        Select Case element.Datatype 'Only subassemblies are computed in totalvolume
            Case Mast, OpenMount, Gasbag, Body, Skid, Wheel, Track, Leg, Wing, AutogyroRotor, TTRotor, CARotor, MMRotor, Hydrofoil, Hovercraft, Superstructure, Pod, Turret, Popturret, Arm
                TempVolume = TempVolume + element.Volume
                'todo: are masts to be included here? gasbags too?
        End Select
    Next
    CalcTotalVolume = TempVolume
End Function

Function CalcPrice() As Single
    ' Add up total cost of everything built into the vehicle except for ammunition
    ' and fuel to find out how much it costs
    Dim TempPrice As Single ' holds the volume in progress
    Dim element As Object
    Dim sQuality As String
    Dim QMod As Single
    
    
    For Each element In Veh.Components
        If Not TypeOf element Is clsWeaponAmmunition Then
            TempPrice = TempPrice + element.Cost
        End If
    Next
    TempPrice = TempPrice
    
    'Caculate the Vehicle Quality Modifiers
    sQuality = Veh.Options.Quality
    If sQuality = "standard" Then
        QMod = 1
    ElseIf sQuality = "cheap" Then
        QMod = 0.5
    ElseIf sQuality = "fine" Then
        QMod = 4
    ElseIf sQuality = "very fine" Then
        QMod = 20
    End If
    
    CalcPrice = (TempPrice + Veh.Options.OptionsCost) * QMod
End Function

Function CalcHealth(BodyHitPoints As Double, LoadedWeight As Single) As Long
    Dim TempHealth As Double
    Dim HMod As Integer
    Dim sQuality As String
    
    ' structural HT = (200 * BodyHitPoints / LoadedWeight) + 5
    ' If the vehicle has Hardpoints, always use the weight WITH hardpoints loaded -
    ' do NOT use two different values
    ' Round HT to the nearest whole number.
    If LoadedWeight = 0 Then
        TempHealth = 1
    Else
        TempHealth = Round((200 * BodyHitPoints / LoadedWeight) + 5, 0)
    End If
    
    'Note: im assume maximum HT of 12 STILL applies. If users disagree
    'i can move the below HMods below the Max HT check below
    'Caculate the Vehicle Quality Modifiers
    sQuality = Veh.Options.Quality
    If sQuality = "standard" Then
        HMod = 0
    ElseIf sQuality = "cheap" Then
        HMod = -1
    ElseIf sQuality = "fine" Then
        HMod = 1
    ElseIf sQuality = "very fine" Then
        HMod = 2
    End If
    
    TempHealth = TempHealth + HMod
    
    ' The maximum allowed structural HT is 12 or the vehicle's TL, whichever is greater.
    If TempHealth > 12 Then
        If gVehicleTL > 12 Then
            TempHealth = Val(gVehicleTL)
        Else
            TempHealth = 12
        End If
    End If
    
    CalcHealth = TempHealth
End Function

Function CalcSizeModifier(ByVal Volume As Double) As Long
    ' todo: will try to optimize later when i finish. Maybe.
    Dim TempModifier As Long
    Dim TVolume As Double
    Dim y As Single
    Dim i As Long
    
    TempModifier = -4
    TVolume = 0.1
    y = 0.7
    i = 0
    
    If Volume <= TVolume Then
        CalcSizeModifier = -4
        Exit Function
    End If
    
    Do
        TVolume = TVolume * 3
        TempModifier = TempModifier + 1
        If Volume <= TVolume Then
            Exit Do
        End If
        TVolume = TVolume + (y * 10 ^ i)
        TempModifier = TempModifier + 1
        If Volume <= TVolume Then
            Exit Do
        End If
        i = i + 1
    Loop
    CalcSizeModifier = TempModifier
End Function

Function CalcTotalGeneratedPower() As Double
    'add all the outputs of all the Power Systems from the Vehicle
    Dim TempOutput As Double
    Dim dblStored As Double
    Dim Keys() As String
    Dim i As Long
    
    On Error GoTo errorhandler
    Keys = Veh.KeyManager.GetCurrentPowerSystemKeys
    
    For i = 1 To UBound(Keys)
        If TypeOf Veh.Components(Keys(i)) Is clsEnergyBank Then
            dblStored = dlbstored + Veh.Components(Keys(i)).Output
        Else
            TempOutput = TempOutput + Veh.Components(Keys(i)).Output
        End If
    Next
    
    mvarTotalStoredPower = dblStored
    CalcTotalGeneratedPower = TempOutput
    Exit Function
    
errorhandler:
        If err.Number = 9 Then 'array not dimensioned properly.  Return a value of 0
            CalcTotalGeneratedPower = 0
            Exit Function
        End If
End Function

Function CalcTotalContinuousPowerConsumption() As Double
    'add all the outputs of all the Power Systems from the Vehicle
    Dim TempPower As Double
    Dim Keys() As String
    Dim i As Long
    
    'todo: we cant determine total power consumed FROM STORED POWER suppliers
    ' unless we use a "Primary" power config setting so that we know which
    ' configuration to base this stat on.
    On Error GoTo errorhandler
    Keys = Veh.KeyManager.GetCurrentPowerConsumptionKeys
    
    For i = 1 To UBound(Keys)
        TempPower = TempPower + Veh.Components(Keys(i)).PowerReqt
    Next
    
    CalcTotalContinuousPowerConsumption = TempPower
    Exit Function
errorhandler:
        If err.Number = 9 Then 'array not dimensioned properly.  Return a value of 0
            CalcTotalContinuousPowerConsumption = 0
            Exit Function
        End If
End Function

Private Sub Class_Initialize()
    mvarTotalSurfaceArea = 0
    mvarStructuralSurfaceArea = 0
    
    mvarEmptyWeight = 0
End Sub

Sub CalcOptionsWeightandCost()

    Dim element As Object
    Dim TotalArea As Single
    Dim BodyArea As Single
    Dim BodyHits As Long
    Dim TrimmedArea As Single
    Dim IgnoredArea As Single
    
    Dim compartmentalizationWeight As Single
    Dim compartmentalizationCost As Single
    Dim ControlledInstabilityCost As Single
    Dim VariableSweepCost As Single
    Dim VariableSweepWeight As Single
    Dim foldingwingsWeight As Single
    Dim foldingwingsCost As Single
    Dim ImprovedSuspensionCost As Single
    Dim SnowTiresCost As Single
    Dim RacingTiresCost As Single
    Dim PunctureResistantCost As Single
    Dim WheelBladesCost As Single
    Dim WheelBladesWeight As Single
    Dim OtherWheelCosts As Single
    
    
    
    For Each element In Veh.Components
        'update the weight and cost for the armor
        If TypeOf element Is clsArmor Then
            element.StatsUpdate
        'this is not a misprint.  See page92 under the table for rules
        'which say you do not include surfaceare for Skids, Gasbag and Masts
        ElseIf TypeOf element Is clsMast Then
            IgnoredArea = IgnoredArea + element.SurfaceArea
        ElseIf TypeOf element Is clsSkid Then
            IgnoredArea = IgnoredArea + element.SurfaceArea
        ElseIf TypeOf element Is clsGasbag Then
            IgnoredArea = IgnoredArea + element.SurfaceArea
        
        ElseIf TypeOf element Is clsBody Then
            With element
                .CalcCompartmentalizationStats
                compartmentalizationWeight = compartmentalizationWeight + .compartmentalizationWeight
                compartmentalizationCost = compartmentalizationCost + .compartmentalizationCost
                ImprovedSuspensionCost = ImprovedSuspensionCost + .ImprovedSuspensionCost
            End With
        ElseIf TypeOf element Is clsWheel Then
            With element
                ImprovedSuspensionCost = ImprovedSuspensionCost + .ImprovedSuspensionCost
                SnowTiresCost = SnowTiresCost + .SnowTiresCost
                RacingTiresCost = RacingTiresCost + .RacingTiresCost
                PunctureResistantCost = PunctureResistantCost + .PunctureResistantCost
                OtherWheelCosts = OtherWheelCosts + .ImprovedBrakesCost + .AllWheelSteeringCost + .SmartWheelsCost
                WheelBladesCost = WheelBladesCost + .WheelBladesCost
                WheelBladesWeight = WheelBladesWeight + .WheelBladesWeight
                
            End With
        ElseIf (TypeOf element Is clsSkid) Or (TypeOf element Is clsLeg) Or (TypeOf element Is clsTrack) Then
            ImprovedSuspensionCost = ImprovedSuspensionCost + element.ImprovedSuspensionCost
        ElseIf (TypeOf element Is clsSuperStructure) Or (TypeOf element Is clsTurret) Or (TypeOf element Is clsPopTurret) Then
            With element
                .CalcCompartmentalizationStats
                compartmentalizationWeight = compartmentalizationWeight + .compartmentalizationWeight
                compartmentalizationCost = compartmentalizationCost + .compartmentalizationCost
            End With
        ElseIf TypeOf element Is clsWing Then
            With element
                .CalcWingRotorOptionWeightsAndCosts
                foldingwingsWeight = foldingwingsWeight + .FoldingWeight
                VariableSweepWeight = VariableSweepWeight + .VariableSweepWeight
                VariableSweepCost = VariableSweepCost + .VariableSweepCost
                ControlledInstabilityCost = ControlledInstabilityCost + .ControlledInstabilityCost
            End With
        ElseIf TypeOf element Is clsRotor Then
            With element
                .CalcWingRotorOptionWeightsAndCosts
                foldingwingsWeight = foldingwingsWeight + .FoldingWeight
                foldingwingsCost = foldingwingsCost + .FoldingCost
                ControlledInstabilityCost = ControlledInstabilityCost + .ControlledInstabilityCost
            End With
        End If
    Next
    

    ' Get weight and cost for Rams
    If mvarRam Then
        mvarRamWeight = 1 * BodyArea
        mvarRamCost = 2 * BodyArea
    Else
        mvarRamWeight = 0
        mvarRamCost = 0
    End If
    ' Get weight and cost for Bulldozers
    If mvarBulldozer Then
        mvarBullDozerWeight = 2 * BodyArea
        mvarBullDozerCost = 4 * BodyArea
    Else
        mvarBullDozerWeight = 0
        mvarBullDozerCost = 0
    End If
    ' Get wieght and cost for Plows
    If mvarPlow Then
        mvarPlowWeight = 2 * BodyArea
        mvarPlowCost = 4 * BodyArea
    Else
        mvarPlowWeight = 0
        mvarPlowCost = 0
    End If
    ' Get weight and cost for Convertible
    If mvarConvertible = "none" Then
        mvarConvertibleCost = 0
        mvarConvertibleWeight = 0
    Else
    End If
    ' Get weight and cost for Hitch
    If mvarHitch Then
        mvarHitchCost = 0.1 * BodyHits
        mvarHitchWeight = mvarHitchCost
    Else
        mvarHitchCost = 0
        mvarHitchWeight = 0
    End If
    ' Get weight and cost for Pin
    If mvarPin <> "none" Then
        mvarPinCost = 0.05 * BodyHits
        mvarPinWeight = 0.1 * BodyHits
        If mvarPin = "explosive" Then
            mvarPinCost = mvarPinCost * 5
        End If
    Else
        mvarPinCost = 0
        mvarPinWeight = 0
    End If
    'Get weight and cost for Roll Stabilizers
    If mvarRollStabilizers = False Then
        mvarRollStabilizersCost = 0
        mvarRollStabilizersWeight = 0
    Else
        'do divide by zero check
        If Veh.Stats.StructuralSurfaceArea = 0 Then
            mvarRollStabilizersCost = 0
            mvarRollStabilizersWeight = 0
        Else
            mvarRollStabilizersCost = 0.1 * (BodyArea / Veh.Stats.StructuralSurfaceArea) * Veh.Stats.StructuralCost
            mvarRollStabilizersWeight = 0.05 * (BodyArea / Veh.Stats.StructuralSurfaceArea) * Veh.Stats.StructuralWeight
        End If
    End If
           
    mvarOptionsWeight = mvarRollStabilizersWeight + mvarConvertibleWeight + _
    mvarPinWeight + mvarHitchWeight + mvarPlowWeight + mvarBullDozerWeight + _
    mvarRamWeight + WheelBladesWeight + compartmentalizationWeight + _
    VariableSweepWeight + foldingwingsWeight
    
    'note the MagicLevitationEnergyCost is not in dollars but in units of Mana or Energy right? thats why its not added to OptionsCost
    mvarOptionsCost = mvarRollStabilizersCost + mvarConvertibleCost + _
    mvarPinCost + mvarHitchCost + mvarPlowCost + mvarBullDozerCost + mvarRamCost + _
    VariableSweepCost + ControlledInstabilityCost + _
    foldingwingsCost + ImprovedSuspensionCost + _
    compartmentalizationCost + mvarTotalSnowTiresCost + _
    RacingTiresCost + PuntureResistantCost + WheelBladesCost + _
    OtherWheelCosts
End Sub

