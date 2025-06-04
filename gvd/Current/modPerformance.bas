Attribute VB_Name = "modPerformance"
Option Explicit

'//////////////////////////////////////////////////////////////////////
' modPerformance.base  - Michael P. Joseph
' Created - 11/18/98
' Helper Functions used by the clsPerformance.cls
'//////////////////////////////////////////////////////////////////////

Public Function GetMotiveAssemblyKey(ByVal PerformanceType As Long) As String
    ' code originally butchered from frmNewProfile (that form is obsolete)
    Dim KeyChain() As String
    Dim element As Object
    Dim i As Long
    Dim Datatype As Long
    On Error GoTo errorhandler

    KeyChain = Veh.KeyManager.GetCurrentSubAssembliesKeys
    
    If PerformanceType = PERFORMANCEWHEEL Then
        PerformanceType = Wheel
    ElseIf PerformanceType = PERFORMANCESKID Then
        PerformanceType = Skid
    ElseIf PerformanceType = PERFORMANCETRACK Then
        PerformanceType = Track
    ElseIf PerformanceType = PERFORMANCELEG Then
        PerformanceType = Leg
    ElseIf PerformanceType = PERFORMANCEFLEX Then
        PerformanceType = Body
    End If
    
    If UBound(KeyChain) = 1 And KeyChain(1) = "" Then
        GoTo errorhandler
    Else
        For i = 1 To UBound(KeyChain)
            Datatype = Veh.Components(KeyChain(i)).Datatype
            If Datatype = PerformanceType Then
                GetMotiveAssemblyKey = KeyChain(i)
                Exit Function
            End If
        Next
    End If
    '
    'For Each element In Veh.Components
    '    Select Case element.Datatype
    '        Case Wheel, Skid, Track, Leg, FlexibodyDrivetrain
    '            GetMotiveAssembly = element.Key
    '            Exit Function
    '    End Select
    'Next

errorhandler:
        GetMotiveAssemblyKey = ""
End Function
        

Public Sub GetVehicleWeight(ByVal lngPerformanceType As Long, _
                            ByVal sngPercentAuxVehicleWeight As Long, _
                            ByVal sngPercentCargoWeight As Long, _
                            ByVal sngPercentAmmunitionWeight As Long, _
                            ByVal sngPercentHardpointWeight As Long, _
                            ByVal sngPercentFuelWeight As Long, _
                            ByVal sngPercentProvisionWeight As Long, _
                            ByRef dblWeight As Double, ByRef dblMass As Double)
                            
    Dim HardPointWeight As Single
    Dim FuelWeight As Single
    Dim ProvisionsWeight As Single
    Dim AmmoWeight As Single
    Dim GunCarriagesWeight As Single
    Dim CargoWeight As Single
    Dim AuxVehiclesWeight As Single
    Dim element As Object
  
    
    On Error Resume Next
    '//based on the options, this routine sets the m_VWeight and m_VMass variables
    If lngPerformanceType = PERFORMANCESUB Then
        dblWeight = Veh.Stats.SubmergedWeight
    Else
        dblWeight = Veh.Stats.HLoadedWeight 'MPJ 07/25/2000 Was using Loaded Weight instead of HardpointLoadedWeight
                                        ' the user can chose to use a % of hardpoint weight (0 to 100%) in the performance
                                        ' profile Edit dialog
    End If
    
    HardPointWeight = Veh.Stats.HLoadedWeight - Veh.Stats.LoadedWeight
    If HardPointWeight < 0 Then HardPointWeight = 0
    
    For Each element In Veh.Components
        If TypeOf element Is clsFuelTank Then
            FuelWeight = element.FuelWeight
        ElseIf TypeOf element Is clsCargo Then
            CargoWeight = CargoWeight + element.CargoWeight
        
        ElseIf TypeOf element Is clsProvisions Then
            ProvisionsWeight = ProvisionsWeight + element.Weight
        ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsVehicleStorage Then
            AuxVehiclesWeight = AuxVehiclesWeight + element.CraftWeight
        ElseIf TypeName(element) = "clsWeaponAmmunition" Then 'check for ammunition
            AmmoWeight = AmmoWeight + element.Weight
        'ElseIf TypeName(element) = "clsWeaponGun" Then 'check for guns with carriages
        '    If element.Carriage Then
        '        GunCarriagesWeight = GunCarriagesWeight + element.Weight
        '    End If
        End If
    Next
    
    '//now lets subtract all of these from our weight
    dblWeight = dblWeight - CargoWeight - HardPointWeight - FuelWeight - ProvisionsWeight - AmmoWeight
    
    '//now add the percentages of the weight
    dblWeight = dblWeight + (AuxVehiclesWeight * sngPercentAuxVehicleWeight) + _
                (CargoWeight * sngPercentCargoWeight) + (AmmoWeight * sngPercentAmmunitionWeight) + _
                (HardPointWeight * sngPercentHardpointWeight) + (FuelWeight * sngPercentFuelWeight) + _
                (ProvisionsWeight * sngPercentProvisionWeight)
        
    dblMass = dblWeight / 2000 'get our mass
    
End Sub

Public Function GetSlowestAnimalSpeed(KeyChain As Variant) As Single
    Dim i As Long
    Dim sKey As String
    Dim dType As Integer
    Dim Slowest As Single
    On Error Resume Next
    
    For i = 1 To UBound(KeyChain)
        sKey = KeyChain(i)
        If sKey = "" Then Exit For
        dType = Veh.Components(sKey).Datatype
        'check for animals.  Max vehicle ground speed cant exceed slowest animal
        If (dType = RopeHarness) Or (dType = YokeandPoleHarness) Or (dType = ShaftandCollarHarness) Or (dType = WhiffletreeHarness) Then
            ' set the first animal to the slowest
            If Slowest = 0 Then Slowest = Veh.Components(sKey).Speed
            ' check if this animal is slower than the current slowest
            If Veh.Components(sKey).Speed < Slowest Then Slowest = Veh.Components(sKey).Speed
        End If
    Next
    
    GetSlowestAnimalSpeed = Slowest
End Function


Public Function GetTiltRotorForwardThrust(ByVal sngPercentThrust As Single) As Single
    '//if the user has checked the TreatTiltRotorsAsPropellers, then
    '//we use their thrust to for forward thrust and NOT for lift
    Dim retval As Single
    Dim element As Object
    
    For Each element In Veh.Components
        If TypeOf element Is clsHelicopterDrivetrain Then
            Select Case element.Datatype
                Case CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain
                    If element.TiltRotor Then
                       retval = retval + (3.5 * element.MotivePower * sngPercentThrust)
                    End If
            End Select
        End If
    Next
    
    GetTiltRotorForwardThrust = retval

End Function

Public Function CalcTotalStaticLift(ByRef KeyChain As Variant, _
                                    ByVal bTreatTiltRotorsAsPropellers As Boolean, _
                                    ByVal bSEVSidewalls As Boolean, _
                                    ByVal bGEVSkirt As Boolean, _
                                    ByVal lngPerformanceType As Long, _
                                    ByVal sngPercentThrust As Single, _
                                    ByVal dblWeight As Double) As Single

    Dim i As Integer
    Dim Templift As Single
    Dim TempFanLift As Single
    Dim sKey As String
    Dim sngVehicleEmptyWeight As Single
    Dim dType As Long
    Dim element As Object
    
    '//search the vehicle for liftin gas
    For Each element In Veh.Components
        If TypeOf element Is clsLiftingGas Then
              Templift = Templift + element.Lift
        End If
    Next
                    
    'next get the lift for Levitation (page 41)
    With Veh.Surface
        sngVehicleEmptyWeight = dblWeight
        If .bMagicLevitation Then
            Templift = Templift + sngVehicleEmptyWeight
        End If
        If .bAntigravityCoating Then
            Templift = Templift + sngVehicleEmptyWeight
        End If
        If .bSuperScienceCoating Then
            Templift = Templift + sngVehicleEmptyWeight
        End If
    End With
    
    'exit if there are no propulsion systems in the keychain
    If KeyChain(1) = "" Then
        CalcTotalStaticLift = Templift
        Exit Function
    End If
    
    'cycle through each item in the array to find all static lift components
    For i = 1 To UBound(KeyChain)
        sKey = KeyChain(i)
         dType = Veh.Components(sKey).Datatype
        
        With Veh.Components(sKey)
            Select Case dType
                Case CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain
                    If bTreatTiltRotorsAsPropellers Then
                        If .TiltRotor Then
                            '//we dont add the lift of a tilt rotor when its in forward flight mode
                        Else
                            Templift = Templift + .Lift * sngPercentThrust
                        End If
                    Else
                        Templift = Templift + .Lift * sngPercentThrust
                    End If
                Case OrnithopterDrivetrain
                    Templift = Templift + .Lift * sngPercentThrust
                Case ContraGravGenerator
                    Templift = Templift + .Lift '* sngPercentThrust
                Case DuctedFan
                    'MPJ 06/30/2000  Fixed.  HoverFan option was not being run at all
                    ' since the old code was If .Liftengine then if .Hoveran rather than Elseif
                    TempFanLift = .MotiveThrust * sngPercentThrust
                    If .LiftEngine Then
                        Templift = TempFanLift ' no multiplier.  LiftEngines cant be used for forward propulsion though
                        
                        '//if its a GEV skirt hovercraft then
                        '//hoverfan lift is increased by 5
                        '//for SEV sidewalls its increased by 4
                        '//else its increased by 2
                    ElseIf .HoverFan Then
                        If lngPerformanceType = PERFORMANCEHOVER Then
                            If bSEVSidewalls Then
                                Templift = Templift + (4 * TempFanLift)
                            ElseIf bGEVSkirt Then
                                Templift = Templift + (5 * TempFanLift)
                            Else
                                Templift = Templift + (2 * TempFanLift)
                            End If
                        End If
                    End If
                Case Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam
                    If .LiftEngine Then
                        If .Afterburner Then 'PPP this should only include AB if its checked!
                            Templift = Templift + .ABThrust
                        Else
                            Templift = Templift + .MotiveThrust * sngPercentThrust
                        End If
                    End If
                Case LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, _
                    OptimizedFusion, AntimatterThermal, AntimatterPion
                    If .LiftEngine Then
                        Templift = Templift + .MotiveThrust * sngPercentThrust
                    End If
                Case OrionEngine
                    If .LiftEngine Then
                        Templift = Templift + .MotiveThrust * sngPercentThrust
                    End If
               Case StandardThruster, SuperThruster, MegaThruster
                    If .LiftEngine Then
                        Templift = Templift + .MotiveThrust * sngPercentThrust
                    End If
            End Select
        End With
    Next
    
    CalcTotalStaticLift = Templift
End Function

Function NoPriorPeriscopeChildren(ByVal parentkey As String) As Boolean
    Dim element As Object
    Dim i As Integer

    i = 0
    For Each element In Veh.Components
        If element.LogicalParent = parentkey Then
            i = i + 1
            If i >= 2 Then
                NoPriorPeriscopeChildren = False
                Exit Function
            End If
        End If
    Next
    ' if it makes it out of the loop, no children found and fuction returns TRUE
    NoPriorPeriscopeChildren = True
End Function

Function GetTotalHitPoints(ByVal Classname As String) As Double
    'returns the total number of hit points for all objects of a given class
    Dim element As Object
    Dim TempHitPoints As Double
    
    For Each element In Veh.Components
        If TypeName(element) = Classname Then
            TempHitPoints = TempHitPoints + element.HitPoints
        End If
    Next
    GetTotalHitPoints = TempHitPoints
End Function

'Function GetLowestDR() As Long
'' this is used to find max speed restrictions for aerial performance
''only DR from metal, composite or laminate armor counts
''JAW 2000.06.06
''
' Dim element As Object
' Dim lngRetval As Long
' Dim lngTemp As Long
' Dim bDRSet As Boolean
' Dim lngArmorCount As Long
' On Error Resume Next
' '//must initialize the DR to 21 and from there we check for less
' lngRetval = 21
'
'    For Each element In Veh
'        If TypeOf element Is clsArmor Then
'            InfoPrint 1, "Armor"
'            If TypeOf Veh.Components(element.LogicalParent) Is clsPopTurret Then
'                    'skip these DR's
'                    InfoPrint 1, " - PopTurret"
'            ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsSkid Then
'                    InfoPrint 1, " - Skid"
'                    If element.Retractable Then
'                        'skip retractables
'                    Else
'                        lngTemp = element.GetLowestDR
'                        bDRSet = True
'                    End If
'            ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsWheel Then
'                    InfoPrint 1, " - Wheel"
'                    If Veh.Components(element.LogicalParent).SubType = "retractable" Then
'                        'skip retractables
'                    Else
'                        lngTemp = element.GetLowestDR
'                        bDRSet = True
'                    End If
'            Else
'                InfoPrint 1, " - other?"
'                lngTemp = element.GetLowestDR
'                bDRSet = True
'            End If
'        End If
'        If bDRSet Then
'            lngRetval = Minimum(lngRetval, lngTemp)
'        End If
'    Next
'
'    '//if armor was never set, then we no we have no armor
'    If bDRSet = False Then
'        lngRetval = 0
'    End If
'
'    GetLowestDR = lngRetval
'End Function

Function GetLowestDR() As Long
' this is used to find max speed restrictions for aerial performance
'only DR from metal, composite or laminate armor counts
'JAW 2000.06.19
'refined to count overall and armor-by-location together properly

 Dim element As Object
 Dim lngRetval As Long
 Dim lngTemp As Long
 Dim bDRSet As Boolean
 Dim lngArmorCount As Long
    Dim MinDRX As Long
 Dim OverallDR As Long
 Dim DR1 As Long
 Dim DR2 As Long
 Dim DR3 As Long
 Dim DR4 As Long
 Dim DR5 As Long
 Dim DR6 As Long
 
 On Error Resume Next
 '//must initialize the DR to 21 and from there we check for less
     lngRetval = 999999999
        MinDRX = 999999999
    For Each element In Veh
        If TypeOf element Is clsArmor Then
            If TypeOf Veh.Components(element.LogicalParent) Is clsPopTurret Then
                    'skip these DR's
            ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsSkid Then
                    If element.Retractable Then
                        'skip retractables
                    Else
                        lngTemp = element.GetLowestDR
                        bDRSet = True
                    End If
            ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsWheel Then
                    If Veh.Components(element.LogicalParent).SubType = "retractable" Then
                        'skip retractables
                    Else
                        lngTemp = element.GetLowestDR
                        bDRSet = True
                    End If
            Else
                AccumulateDR OverallDR, DR1, DR2, DR3, DR4, DR5, DR6, element
                'lngTemp = element.GetLowestDR
                'bDRSet = True
            End If
        End If
        If bDRSet Then
            lngRetval = Minimum(lngRetval, lngTemp)
        End If
    Next
    
    
    MinDRX = Minimum(DR1, MinDRX)
    MinDRX = Minimum(DR2, MinDRX)
    MinDRX = Minimum(DR3, MinDRX)
    MinDRX = Minimum(DR4, MinDRX)
    MinDRX = Minimum(DR5, MinDRX)
    MinDRX = Minimum(DR6, MinDRX)
        
    
    'if bDRSet is true, then exposed components were found.
    If bDRSet = True Then
        lngRetval = Minimum(lngRetval, MinDRX)
    Else
        lngRetval = MinDRX
    End If
    
    GetLowestDR = OverallDR + lngRetval
End Function

Sub AccumulateDR(ByRef AccumDR As Long, ByRef AccumDR1 As Long, ByRef AccumDR2 As Long, _
        ByRef AccumDR3 As Long, ByRef AccumDR4 As Long, ByRef AccumDR5 As Long, _
        ByRef AccumDR6 As Long, ByRef ArmorChunk As clsArmor)
    'JAW 2000.06.19
    'mostly cribbed from the GetLowestDR member function of clsArmor
    
    On Error Resume Next
    'Dim lngRetval As Long
    With ArmorChunk
    
    Select Case .Datatype
       Case ArmorComplexFacing
            Select Case .Material1
                Case "metal", "composite", "laminate"
                    AccumDR1 = AccumDR1 + .DR1
                    Select Case .Material2
                        Case "metal", "composite", "laminate"
                            AccumDR2 = AccumDR2 + .DR2
                            Select Case .Material3
                                Case "metal", "composite", "laminate"
                                    AccumDR3 = AccumDR3 + .DR3
                                    Select Case .Material4
                                        Case "metal", "composite", "laminate"
                                            AccumDR4 = AccumDR4 + .DR4
                                            Select Case .Material5
                                                Case "metal", "composite", "laminate"
                                                    AccumDR5 = AccumDR5 + .DR5
                                                    Select Case .Material6
                                                        Case "metal", "composite", "laminate"
                                                            AccumDR6 = AccumDR6 + .DR6
                                                        Case Else
                                                            'lngRetval = 0
                                                    End Select
                                                Case Else
                                                    'lngRetval = 0
                                            End Select
                                        Case Else
                                            'lngRetval = 0
                                    End Select
                                Case Else
                                    'lngRetval = 0
                            End Select
                        Case Else
                            'lngRetval = 0
                    End Select
                Case Else
                    'lngRetval = 0
            End Select
           
       Case ArmorBasicFacing
           Select Case .Material
                Case "metal", "composite", "laminate"
                    AccumDR1 = AccumDR1 + .DR1
                    AccumDR2 = AccumDR2 + .DR2
                    AccumDR3 = AccumDR3 + .DR3
                    AccumDR4 = AccumDR4 + .DR4
                    AccumDR5 = AccumDR5 + .DR5
                    AccumDR6 = AccumDR6 + .DR6
                
                Case Else
                    'lngRetval = 0
            End Select
       
       Case ArmorLocation, ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
            Select Case .Material
                Case "metal", "composite", "laminate"
                    AccumDR = AccumDR + .DR
                Case Else
                    'lngRetval = 0
            End Select
          
    End Select

    End With
    'GetLowestDR = lngRetval
End Sub

Function GetLowestTL(ByVal Classname As String) As Long
    'returns the Tech Level of the lowest Tech Level objects of a given class
    Dim element As Object
    Dim TempTL As Long
    Dim NewTempTL As Long
    
    For Each element In Veh.Components
        If TypeName(element) = Classname Then
            NewTempTL = element.TL
        End If
        
        If TempTL = 0 Then
            TempTL = NewTempTL
        Else
            TempTL = Minimum(TempTL, NewTempTL)
        End If
    Next
    GetLowestTL = TempTL
End Function

Function MinimumNonZero(ByVal x As Variant, ByVal y As Variant) As Variant
    'compares x and y and returns the minimum non zero number
    
    If x = 0 Then
        MinimumNonZero = y
        Exit Function
    ElseIf y = 0 Then
        MinimumNonZero = x
        Exit Function
    End If
    
    If x < y Then MinimumNonZero = x Else MinimumNonZero = y

End Function


Function Minimum(ByVal x As Variant, ByVal y As Variant) As Variant
    'compares x and y and returns the minimum of the two
    If x < y Then Minimum = x Else Minimum = y
End Function

Function Maximum(ByVal x As Variant, ByVal y As Variant) As Variant
    'compares x and y and returns the maximum of the two
    If x > y Then Maximum = x Else Maximum = y
End Function

Function GetHovercraftType() As String
    'returns whether the vehicle has hovercraft with sev sidewalls
    Dim element As Object
    
    For Each element In Veh.Components
        If TypeOf element Is clsHovercraft Then
            If element.SubType = "SEV Sidewalls" Then
                GetHovercraftType = "SEV"
                Exit Function
            Else
                GetHovercraftType = "GEV"
                Exit Function
            End If
        End If
    Next
    GetHovercraftType = "none"
End Function

Function VehicleHasResponsiveStruct() As Boolean
    'returns true if the vehicle has a responsive structure
    If Veh.Components(BODY_KEY).Responsive Then
        VehicleHasResponsiveStruct = True
    Else
        VehicleHasResponsiveStruct = False
    End If
End Function

Function VehicleHasWings() As Boolean
    'returns true if the vehicle has wings
    Dim element As Object

    For Each element In Veh.Components
        If TypeOf element Is clsWing Then
            VehicleHasWings = True
            Exit Function
        End If
    Next

    VehicleHasWings = False

End Function

Function VehicleHasFlarecraftWings() As Boolean
    'returns true if the vehilce has flarecraft wings on any wing assembly
    Dim element As Object

    For Each element In Veh.Components
        If TypeOf element Is clsWing Then
            If element.SubType = "flarecraft" Then
                VehicleHasFlarecraftWings = True
                Exit Function
            End If
        End If
    Next
    VehicleHasFlarecraftWings = False
End Function

Function VehicleHasRotors() As Boolean
    'returns true if the vehicle has rotors subassembly
    Dim element As Object

    For Each element In Veh.Components
        If TypeOf element Is clsRotor Then
            VehicleHasRotors = True
            Exit Function
        End If
    Next
    VehicleHasRotors = False
End Function


Function VehicleHasNonTiltRotors() As Boolean
    'returns true if the vehicle has rotors subassembly
    Dim element As Object

    For Each element In Veh.Components
        If TypeOf element Is clsHelicopterDrivetrain Then
            If element.TiltRotor = False Then
                VehicleHasNonTiltRotors = True
                Exit Function
            End If
        End If
    Next
    VehicleHasNonTiltRotors = False
End Function

Function VehicleHasCoaxialRotors() As Boolean
    'returns true if the vehicle has coaxial rotors installs
    Dim element As Object
    For Each element In Veh.Components
        If TypeOf element Is clsRotor Then
            If element.Datatype = CARotor Then
                VehicleHasCoaxialRotors = True
                Exit Function
            End If
        End If
    Next
    VehicleHasCoaxialRotors = False
End Function

Function VehiclehasElectORCompcontrols() As Boolean
    'returns true if the vehicle has EITHER electronic or computerized controls
    Dim element As Object
    
    For Each element In Veh.Components
        Select Case element.Datatype
        Case ElectronicDivingControl, ComputerizedDivingControl, _
            ElectronicManeuverControl, ComputerizedManeuverControl
            
            VehiclehasElectORCompcontrols = True
            Exit Function
        End Select
    Next
    VehiclehasElectORCompcontrols = False
End Function

Function VehicleHasCompControls() As Boolean
    Dim element As Object
    
    For Each element In Veh.Components
        Select Case element.Datatype
        Case ComputerizedDivingControl, ComputerizedManeuverControl
            
            VehicleHasCompControls = True
            Exit Function
        End Select
    Next
    VehicleHasCompControls = False
End Function

Function AllWingsAreHighAgility() As Boolean
    Dim element As Object
    Dim All As Boolean
    
    For Each element In Veh.Components
        If TypeOf element Is clsWing Then
            If element.SubType = "high agility" Then
                All = True
            Else
                All = False
                Exit For
            End If
        End If
    Next
    AllWingsAreHighAgility = All
End Function
 
Function AllWingsAreVariableSweep() As Boolean
    Dim element As Object
    Dim All As Boolean
    
    For Each element In Veh.Components
        If TypeOf element Is clsWing Then
            If element.VariableSweep <> "none" Then
                All = True
            Else
                All = False
                Exit For
            End If
        End If
    Next
    AllWingsAreVariableSweep = All
End Function
    
Function AllWingsRotorsControlledInstability() As Boolean
    Dim element As Object
    Dim component As Integer
    Dim All As Boolean
    'determines if all wings and or rotors have controlled instability set
    
    All = False 'init the flag
    
    For Each element In Veh.Components
        component = element.Datatype
        If component = Wing Then
            If element.ControlledInstability Then
                All = True
            Else
                All = False
                Exit For
            End If
        ElseIf (component = AutogyroRotor) Or (component = CARotor) Or (component = MMRotor) Or (component = TTRotor) Then
            If element.ControlledInstability Then
                All = True
            Else
                All = False
                Exit For
            End If
        End If
    Next
    AllWingsRotorsControlledInstability = All
End Function

Function VehicleHasMMRRotors() As Boolean
    Dim element As Object
    
    For Each element In Veh.Components
        If element.Datatype = MMRotor Then
            VehicleHasMMRRotors = True
        End If
    Next
    VehicleHasMMRRotors = False
End Function

Function VehicleHasBipeorTripWings()
    Dim element As Object
    Dim All As Boolean
    
    For Each element In Veh.Components
        If element.Datatype = Wing Then
            If (element.SubType = "biplane") Or (element.SubType = "triplane") Then
                All = True
            Else
                All = False
                Exit For
            End If
        End If
    Next
    VehicleHasBipeorTripWings = All
End Function

Function VehicleHasOnlyStubWings()
    Dim element As Object
    Dim All As Boolean
    
    For Each element In Veh.Components
        If TypeOf element Is clsWing Then
            If element.SubType = "stub" Then
                All = True
            Else
                All = False
                Exit For
            End If
        End If
    Next
    VehicleHasOnlyStubWings = All
End Function
