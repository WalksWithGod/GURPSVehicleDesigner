VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarKey As String
Private mvarPerformanceType As String
Private mvaraAcceleration As Integer 'local copy
Private mvaraDeceleration As Integer 'local copy
Private mvaraDrag As Integer 'local copy
Private mvaraManeuverability As Integer 'local copy
Private mvaraMotiveThrust As Integer 'local copy
Private mvaraStability As Integer 'local copy
Private mvaraStallSpeed As Integer 'local copy
Private mvaraTopSpeed As Integer 'local copy
Private mvargSpeedFactor As Integer 'local copy
Private mvargAcceleration As Integer 'local copy
Private mvargDeceleration As Integer 'local copy
Private mvargManeuverability As Integer 'local copy
Private mvargOffRoad As Integer 'local copy
Private mvargPressure As Integer 'local copy
Private mvargStability As Integer 'local copy
Private mvargTopSpeed As Integer 'local copy
Private mvarhAcceleration As Integer 'local copy
Private mvarhDeceleration As Integer 'local copy
Private mvarsFTLSpeed As Integer 'local copy
Private mvarsManeuverability As Integer 'local copy
Private mvarsStability As Integer 'local copy
Private mvarsTopSpeed As Integer 'local copy
Private mvarsuAcceleration As Integer 'local copy
Private mvarsuCrushDepth As Integer 'local copy
Private mvarsuDeceleration As Integer 'local copy
Private mvarsuIDeceleration As Integer
Private mvarsuHydroDrag As Integer 'local copy
Private mvarsuManeuverability As Integer 'local copy
Private mvarsuStability As Integer 'local copy
Private mvarsuTopSpeed As Integer 'local copy
Private mvarsuTotalAquaticThrust As Double
Private mvarwAcceleration As Integer 'local copy
Private mvarwDeceleration As Integer 'local copy
Private mvarwIDeceleration As Double 'the increased deceleration
Private mvarwDraft As Integer 'local copy
Private mvarwHydroDrag As Integer 'local copy
Private mvarwManeuverability As Integer 'local copy
Private mvarwStability As Integer 'local copy
Private mvarwTopSpeed As Long 'local copy
Private mvarwPlaningSpeed As Long 'note: planing speed is in addition to topspeed (see page 131 top right column)
Private mvarwHydrofoilSpeed As Long 'note:hydrofoil speed is in addition to topspeed(see page 131 top right column
Private mvarwTotalAquaticThrust As Double

Private arrThruster() As udtThruster

Public Sub Add(Key As String, ClassName As String)
Dim Size As Integer
Dim Exist As Boolean
'init the Flag
Exist = False
'make sure the array has some data it in
If arrThruster(1).Key = "" Then
    'add the item to the array
    arrThruster(1).ClassName = ClassName
    arrThruster(1).Key = Key
End If
'determine if the item already exists in the array or not
For i = 1 To UBound(arrThruster)
    If arrThruster(i).Key = Key Then
        Exist = True
    End If
Next
If Exist = False Then
    ' increase the size of the array by one
    Size = UBound(arrThruster) + 1
    ReDim Preserve arrThruster(Size)
    'add the item to the array
    arrThruster(Size).ClassName = ClassName
    arrThruster(Size).Key = Key
End If
End Sub


Public Function Checked(Key As String) As Boolean
' checks to see if the item is checked or not in the array
For i = 1 To UBound(arrThruster)
    If arrThruster(i).Key = Key Then
        If arrThruster(i).Checked = True Then
            Checked = True
            Exit Function
        Else
            Checked = False
            Exit Function
        End If
    End If
Next
End Function


Public Sub CheckItem(Key As String)
' check the item in the array
For i = 1 To UBound(arrThruster)
    If arrThruster(i).Key = Key Then
        arrThruster(i).Checked = True
        Exit Sub
    End If
Next
End Sub

Public Sub UnCheckItem(Key As String)
' UNcheck the item in the array
For i = 1 To UBound(arrThruster)
    If arrThruster(i).Key = Key Then
        arrThruster(i).Checked = False
        Exit Sub
    End If
Next
End Sub
Public Sub CalcPerformance()
' TODO first it must determine which type of performance it is then
' call appropriate routines to just calculate relevant performance

'gSpeedFactor = CalcGroundSpeedFactor
'gTopSpeed = CalcGroundSpeed

'water performance
    wHydroDrag = CalcHydroDrag
    wTopSpeed = CalcWaterSpeed
    wAcceleration = CalcWaterAcceleration
    CalcWaterMRandSR
    CalcWaterDeceleration 'this does both Deceleration and Increased Deceleration
'submerged performance
    suHydroDrag = CalcSuHydroDrag
    suTopSpeed = CalcSuTopSpeed
    suAcceleration = CalcSuAcceleration
    suDeceleration = wDeceleration
    suIDeceleration = wIDeceleration
    suManeuverabiltiy = wManeuverability
    suStability = wStability
    suDraft = ((colVehicle("1_").SubmergedWeight) ^ 1 / 3) / 3
    suCrushDepth = CalcCrushDepth
'aerial performance
    aStallSpeed = CalcAStallSpeed
    'aTakeOffRun = CalcTakeOffRun TODO need to do these 3
    'aLandingRun = CalcLandingRun
    'aCanFly = CalcACanFly
    aMotiveThrust = CalcAMotiveThrust
    aTopSpeed = CalcATopSpeed
    aAcceleration = CalcAAcceleration
    aManeuverability = CalcAManeuverability
    aStability = CalcAStability
    aDeceleration = CalcADeceleration
    
End Sub
Function CalcAStallSpeed() As Double
Dim SL As Single
Dim RS As Single
Dim TempSpeed As Double
Dim LiftArea As Double
Dim Element As Object
'determine total lift area
For Each Element In colVehicle
    If TypeOf Element Is clsWing Then
        If Element.SubType = "STOL wings" Then
            LiftArea = LiftArea + (1.5 * Element.SurfaceArea)
        ElseIf Element.SubType = "Flarecraft wings" Then
            LiftArea = LiftArea + (3 * Element.SurfaceArea)
        Else
            LiftArea = LiftArea + Element.SurfaceArea
        End If
    If TypeOf Element Is clsRotar Then
        If Element.SubType = "Flarecraft rotars" Then
            LiftArea = LiftArea + (3 * Element.SubType)
        Else
            LiftArea = LiftArea + Element.SurfaceArea
        End If
    End If
Next
If colVehicle("1_").LiftingBody Then
    LiftArea = LiftArea + (0.3 * colVehicle("1_").SurfaceArea)
Else
    LiftArea = LiftArea + (0.1 * colVehicle("1_").SurfaceArea)
End If

'determine the StreamLining modifier
With colVehicle("1_")
    If .StreamLining = "None" Then SL = 1
    ElseIf .StreamLining = "Fair" Then SL = 1
    ElseIf .StreamLining = "Good" Then SL = 1.05
    ElseIf .StreamLining = "Very Good" Then SL = 1.1
    ElseIf .StreamLining = "Superior" Then SL = 1.15
    ElseIf .StreamLining = "Excellent" Then SL = 1.2
    ElseIf .StreamLining = "Radical" Then SL = 1.3
    End If
End With

'determine Responsive structure modififer
If colVehicle.Item(PerformanceType).Responsive Then ' this might choke if it hits FlexibodyDrivetrain since its not an assembly
    RS = 1.5
Else: RS = 2
End If

'do final calculation
TempSpeed = ((colVehicle("1_").LoadedWeight - CalcTotalStaticLift) / LiftArea) * SL * RS
'do final rounding
CalcAStallSpeed = CLng(TempSpeed / 5) * 5 'round to nearest 5mph
End Function

Function CalcAMotiveThrust() As Double


End Function

Function CalcATopSpeed() As Double

End Function

Function CalcAAcceleration() As Double

End Function

Function CalcAManeuverability() As Double

End Function

Function CalcAStability() As Double


End Function

Function CalcADeceleration() As Double


End Function
Function CalcTotalStaticLift() As Double
' TODO should i round the final product?
Dim i As Integer
Dim TempLift As Double

'cycle through each item in the array to find all static lift components
For i = 1 To UBound(arrThruster)
    If arrThruster(i).ClassName = "clsLiftingGas" Then
        TempLift = TempLift + colVehicle.Item(arrThruster(i).Key).Lift
    ElseIf arrThruster(i).ClassName = "clsContraGravGenerator" Then
        TempLift = TempLift + colVehicle.Item(arrThruster(i).Key).Lift
    ElseIf arrThruster(i).ClassName = "clsDuctedFan" Then
        If colVehicle.Item(arrThruster(i).Key).LiftEngine Then
            TempLift = TempLift + colVehicle.Item(arrThruster(i).Key).Lift
        End If
    ElseIf arrThruster(i).ClassName = "clsJetEngine" Then
        If colVehicle.Item(arrThruster(i).Key).LiftEngine Then
            TempLift = TempLift + colVehicle.Item(arrThruster(i).Key).Lift
        End If
    ElseIf arrThruster(i).ClassName = "clsRocketEngine" Then
        If colVehicle.Item(arrThruster(i).Key).LiftEngine Then
            TempLift = TempLift + colVehicle.Item(arrThruster(i).Key).Lift
        End If
    ElseIf arrThruster(i).ClassName = "clsOrionEngine" Then
        If colVehicle.Item(arrThruster(i).Key).LiftEngine Then
            TempLift = TempLift + colVehicle.Item(arrThruster(i).Key).Lift
        End If
    ElseIf arrThruster(i).ClassName = "clsReactionlessThruster" Then
        If colVehicle.Item(arrThruster(i).Key).LiftEngine Then
            TempLift = TempLift + colVehicle.Item(arrThruster(i).Key).Lift
        End If
    End If
Next
CalcTotalStaticLift = TempLift
End Function
Function CalcSuHyrdroDrag() As Double
Dim Ls As Integer
Dim TempLift As Double
Dim MinWeight As Double

TempLift = CalcTotalContragravLift
Hl = GetHl
With colVehicle("1_")
    If .SubmarineLines Then Ls = 10
    If .HydrodynamicLines = "Very Fine" Then Ls = 6
    If .HydrodynamicLines = "Fine" Then Ls = 4
    If .HydrodynamicLines = "Average" Then Ls = 3
    If .HydrodynamicLines = "Mediocre" Then Ls = 2
    If .HydrodynamicLines = "None" Then Ls = 1
    TempWeight = .SubmergedWeight
End With
    
Tempdrag = ((TempWeight ^ 1 / 3) ^ 2) / Ls
CalcSuHydroDrag = CLng(Tempdrag)

End Function
Function CalcSuTopSpeed() As Double
Dim Animal As Boolean
Dim SlowestAnimal As Single
Dim TempSpeed As Double
'init Animal flag
Animal = False
'init the SlowestAnimal
SlowestAnimal = 0

For i = 1 To UBound(arrThruster) - 1
    'check for animals.  Max vehicle ground speed cant exceed slowest animal
    If arrThruster(i).ClassName = "clsAnimal" Then
        Animal = True 'TODO i dont think i need this flag.  Check CalcGroundSpeed also!!
        ' set the first animal to the slowest
        If SlowestAnimal = 0 Then SlowestAnimal = colVehicle(arrThruster(i).Key).Speed
        ' check if this animal is slower than the current slowest
        If colVehicle(arrThruster(i).Key).Speed < SlowestAnimal Then SlowestAnimal = colVehicle(arrThruster(i).Key).Speed
    End If
Next
    ' add motive powers for all other types of thrust components
    Select Case arrThruster(i).ClassName
        Case "clsLegDrivetrain"
            TotalMotivePower = TotalMotivePower + (colVehicle(arrThruster(i).Key).MotivePower * 2)
        Case "clsFlexibodyDrivetrain"
            TotalMotivePower = TotalMotivePower + (colVehicle(arrThruster(i).Key).MotivePower * 5)
        'Case "" ' debug Need to check if putting this hack in is ok
            'do nothing
        Case Else
            TotalMotivePower = TotalMotivePower + colVehicle(arrThruster(i).Key).MotivePower
    End Select
suTotalAquaticThrust = TotalMotivePower 'save the submerged totalaquaticthrust
TempSpeed = ((TotalMotivePower / suHydroDrag) ^ 1 / 3) * 6

' Check to make sure we dont exceed speed of slowest animal(if applicable)
If SlowestAnimal = 0 Then ' do nothing
ElseIf TempSpeed < SlowestAnimal Then ' do nothing
Else ' cant exceed slowest animal!
    TempSpeed = SlowestAnimal
End If
' Do the final rounding
If TempSpeed >= 20 Then
    ' round to nearest 5mph
    TempSpeed = CLng(TempSpeed / 5) * 5
Else ' round to nearest whole number
    TempSpeed = CLng(TempSpeed)
End If

'return the function's value
CalcSuTopSpeed = TempSpeed
End Function
Function CalcSuAcceleration() As Double
Dim TempAcceleration As Double

TempAcceleration = suTotalAquaticThrust / colVehicle("1_").SubmergedWeight * 20
' round to nearest 5 miles per hour unless under 5 then just nearest whole number
If TempAcceleration < 1 Then
    TempAcceleration = (CInt(TempAcceleration * 10)) / 10 'round to 1 decimal place
ElseIf TempAcceleration < 5 Then ' round to nearest 1mph
    TempAcceleration = CInt(TempAcceleration)
Else ' round to nearest 5mph
    TempAcceleration = CLng(TempAcceleration * 5) \ 5
End If
CalcSuAcceleration = TempAcceleration
End Function

Function CalcCrushDepth() As Double
Dim Element As Object
Dim LowestDR As Long
Dim LowObject As Object
Dim TempCrush As Double
Dim SMod As Integer
LowestDR = 0 'init

For Each Element In colVehicle
    If TypeOf Element Is clsEnvironmentalSystem Then
        If Element.subytpe = "Full Life System" Then
            If LowestDR = 0 Then
                LowestDR = GetParentDR(Element.Parent)
                Set LowObject = Element
            ElseIf LowestDR > GetParentDR(Element.Parent) Then
                LowestDR = colVehicle(Element.Parent).DR
                Set LowObject = Element
            End If
        ElseIf Element.SubType = "Limited Life System" Then
            If LowestDR = 0 Then
                LowestDR = GetParentDR(Element.Parent)
                Set LowObject = Element
            ElseIf LowestDR > GetParentDR(Element.Parent) Then
                LowestDR = colVehicle(Element.Parent).DR
                Set LowObject = Element
            End If
        ElseIf Element.SubType = "Total Life System" Then
            If LowestDR = 0 Then
                LowestDR = GetParentDR(Element.Parent)
                Set LowObject = Element
            ElseIf LowestDR > GetParentDR(Element.Parent) Then
                LowestDR = colVehicle(Element.Parent).DR
                Set LowObject = Element
            End If
        End If
    End If
Next
        
'get the structure modifier
With colVehicle.Item(PerformanceType)
    If .structure = "Extra-light" Then SMod = 0.25
    ElseIf .structure = "Light" Then SMod = 0.5
    ElseIf .structure = "Medium" Then SMod = 1
    ElseIf .structure = "Heavy" Then SMod = 2
    ElseIf .structure = "Extra-heavy" Then SMod = 4
    End If
End With
'do final calculations
TempCrush = (LowestDR + 10) * SMod
If colVehicle("1_").Submersible = False Then
    TempCrush = TempCrush / 2
End If
CalcCrushDepth = CLng(TempCrush) 'round to nearest whole number
Set LowObject = Nothing ' make sure we destroy the object
End Function
Function GetParentDR(Parent As String) As Long
GetParentDR = colVehicle(Parent).DR
End Function

Function CalcHydroDrag() As Double
Dim Hl As Integer
Dim TempLift As Double
Dim MinWeight As Double

TempLift = CalcTotalContragravLift
Hl = GetHl
With colVehicle("1_")
    If .Trimaran = True Then Hl = Hl + (Hl * 0.1)
    If .Catamaran = True Then Hl = Hl + (Hl * 0.2)
    TempWeight = .LoadedWeight
End With
    
MinWeight = TempWeight * 0.1
TempWeight = TempWeight - CalcTotalContragravLift
If TempWeight < MinWeight Then TempWeight = MinWeight
Tempdrag = ((TempWeight ^ 1 / 3) ^ 2) / Hl
CalcHydroDrag = CLng(Tempdrag)

End Function

Function GetHl() As Integer
Dim Lines As String
Dim Hl As Integer

Lines = colVehicle("1_").HydrodynamicLines
If Lines = "Very Fine" Then
    Hl = 20
ElseIf Lines = "Fine" Then
    Hl = 15
ElseIf Lines = "Average" Then
    Hl = 10
ElseIf Lines = "Mediocre" Then
    Hl = 5
ElseIf Lines = "None" Then
    Hl = 1
End If

If colVehicle("1_").SubmarineLines = True Then Hl = 5
GetHl = Hl
End Function

Function CalcWaterSpeed() As Double
Dim Animal As Boolean
Dim SlowestAnimal As Single
Dim TempSpeed As Double
'init Animal flag
Animal = False
'init the SlowestAnimal
SlowestAnimal = 0

For i = 1 To UBound(arrThruster) - 1
    'check for animals.  Max vehicle ground speed cant exceed slowest animal
    If arrThruster(i).ClassName = "clsAnimal" Then
        Animal = True 'TODO i dont think i need this flag.  Check CalcGroundSpeed also!!
        ' set the first animal to the slowest
        If SlowestAnimal = 0 Then SlowestAnimal = colVehicle(arrThruster(i).Key).Speed
        ' check if this animal is slower than the current slowest
        If colVehicle(arrThruster(i).Key).Speed < SlowestAnimal Then SlowestAnimal = colVehicle(arrThruster(i).Key).Speed
    End If
Next
    ' add motive powers for all other types of thrust components
    Select Case arrThruster(i).ClassName
        Case "clsTrackedDrivetrain", "clsLegDrivetrain", "clsWheeledDrivetrain"
            TotalMotivePower = TotalMotivePower + (colVehicle(arrThruster(i).Key).MotivePower * 2)
        Case "clsFlexibodyDrivetrain"
            TotalMotivePower = TotalMotivePower + (colVehicle(arrThruster(i).Key).MotivePower * 5)
        'Case "" ' debug Need to check if putting this hack in is ok
            'do nothing
        Case Else
            TotalMotivePower = TotalMotivePower + colVehicle(arrThruster(i).Key).MotivePower
    End Select
wTotalAquaticThrust = TotalMotivePower 'save the totalaquaticthrust
TempSpeed = ((TotalMotivePower / wHydroDrag) ^ 1 / 3) * 6

' add in the streamlining effects
With colVehicle("1_")
If TempSpeed > 50 Then
    If .StreamLining = "None" Then
        If TempSpeed > 150 Then TempSpeed = 150
    ElseIf .StreamLining = "Fair" Then
        TempSpeed = TempSpeed + (TempSpeed * 0.05)
        If TempSpeed > 150 Then TempSpeed = 150
    ElseIf .StreamLining = "Good" Then
        TempSpeed = TempSpeed + (TempSpeed * 0.1)
        If TempSpeed > 150 Then TempSpeed = 150
    Else
        TempSpeed = TempSpeed + (TempSpeed * 0.1)
    End If
End If
End With
' Check to make sure we dont exceed speed of slowest animal(if applicable)
If SlowestAnimal = 0 Then ' do nothing
ElseIf TempSpeed < SlowestAnimal Then ' do nothing
Else ' cant exceed slowest animal!
    TempSpeed = SlowestAnimal
End If
' Do the final rounding
If TempSpeed >= 20 Then
    ' round to nearest 5mph
    TempSpeed = CLng(TempSpeed / 5) * 5
Else ' round to nearest whole number
    TempSpeed = CLng(TempSpeed)
End If

'TODO: check if hydrofoil and planing top speed calculations are done BEFORE or AFTER
' Streamlining effects adjustments below (see page 131 top right column) If they are "before" then
' i need to move Planing and Hydrofoil speed calcs above the Streamlinging effects section
'TODO: check if the errata "hydrofoil modifier" it speaks of is indeed PlaningSpeed * 1.5

'Get the Planing and Hydrofoil Top Speed
Dim NeededHydrofoilSpeed As Double
Dim NeededPlaningSpeed As Double
Dim TempPlaningSpeed As Double
Dim TempHydrofoilSpeed As Double

NeededHydrofoilSpeed = 20 + (colVehicle("1_").SurfaceArea / 100)
NeededPlaningSpeed = ((GetHl * 5) + 5)
If NeededPlaningSpeed >= colVehicle("1_").LoadedWeight * 0.05 Then
    TempPlaningSpeed = TempSpeed * 2
End If
For Each Element In colVehicle
    If TypeOf Element Is clsHydrofoil Then
        If TempPlaningSpeed > 0 Then
            TempHydrofoilSpeed = TempPlaningSpeed * 1.5
        Else
            TempHydrofoilSpeed = TempSpeed * 1.5
        End If
        Exit For 'exit this loop after we find a hydrofoil
    End If
Next
wPlaningSpeed = TempPlaningSpeed 'save the calculations
wHyrdrofoilSpeed = TempHydrofoilSpeed 'save the calculations
'return the function's value
CalcWaterSpeed = TempSpeed
End Function

Function CalcWaterAcceleration() As Double
Dim TempAcceleration As Double

TempAcceleration = (mvartotalaquaticthrust / colVehicle("1_").LoadedWeight) * 20
' round to nearest 5 miles per hour unless under 5 then just nearest whole number
If TempAcceleration < 1 Then
    TempAcceleration = (CInt(TempAcceleration * 10)) / 10 'round to 1 decimal place
ElseIf TempAcceleration < 5 Then ' round to nearest 1mph
    TempAcceleration = CInt(TempAcceleration)
Else ' round to nearest 5mph
    TempAcceleration = CLng(TempAcceleration * 5) \ 5
End If
CalcWaterAcceleration = TempAcceleration
End Function

Sub CalcWaterDeceleration()
Dim Temp As Double

Temp = 100 * (wMR / GetHl)
If Temp > 10 Then Temp = 10
wDeceleration = Temp ' this is for unpowered/drifting deceleration
wIDeceleration = (wAcceleration / 2) + Temp ' this is for powered deceleration
End Sub

Sub CalcDraft()
Dim Hl As Single
Dim TempDraft As Double

With colVehicle("1_")
    TempWeight = .LoadedWeight
    If .HydrodynamicLines = "None" Then Hl = 1
    ElseIf .HydrodynamicLines = "Mediocre" Then Hl = 1.1
    ElseIf .HydrodynamicLines = "Average" Then Hl = 1.2
    ElseIf .HydrodynamicLines = "Fine" Then Hl = 1.3
    ElseIf .HydrodynamicLines = "Very Fine" Then Hl = 1.4
    ElseIf .SubmarineLines Then Hl = 2
    End If
End With
TempDraft = (((TempWeight) ^ 1 / 3) / 15) * Hl
wDraft = (CLng(TempDraft * 10)) / 10 'round to one decimal place

End Sub

Sub CalcWaterMRandSR()
Dim Category As Integer
Dim TempMR As Double
Dim TempSR As Single
Dim TempType As udtWaterStability
Dim TempTech As Integer
Dim MRBonus As Single
Dim SRBonus As Integer

Dim i As Integer
MRBonus = 0 'init the bonus value
SRBonus = 0
With colVehicle("1_")
    TempTech = .StructuralTL
    If .Volume <= 100 Then
        Category = 1
    ElseIf .Volume <= 1000 Then Category = 2
    ElseIf .Volume <= 10000 Then Category = 3
    ElseIf .Volume <= 100000 Then Category = 4
    ElseIf .Volume <= 1000000 Then Category = 5
    Else: Category = 6
    End If
End With
'get the initial values
TempType = WaterStabMatrix(TempTech, Category)
TempSR = TempType.SR ' need to save this early since we will inadvertantly modify it

'make the MR category adjustments
If colVehicle("1_").Responsive = True Then
    If Category = 1 Then
        MRBonus = MRBonus + 0.25
    Else
        Category = Category - 1
    End If
End If
For i = 1 To UBound(arrThruster) - 1
    'check for Flexibody drivetrain modifier
    If arrThruster(i).ClassName = "clsFlexibodyDrivetrain" Then
        If Category = 1 Then
            MRBonus = MRBonus + 0.25
        Else
            Category = Category - 1
        End If
    End If
Next
For Each Element In colVehicle
    If TypeOf Element Is clsManeuverControl Then
        If Element.SubType = "Mechanical" Then
        Else 'add modifier for Electric or comptuerized controls
            If Category = 1 Then
                MRBonus = MRBonus + 0.25
            Else
                Category = Category - 1
            End If
        SRBonus = SRBonus + 1 ' one of the SR modifiers
        End If
    End If
Next

'set the final value for maneuverability
TempType = WaterStabMatrix(TempTech, Category)
wManeuverability = TempType.MR + MRBonus

'make the SR adjustments
With colVehicle("1_")
If .RollStabilizers Then SRBonus = SRBonus + 1
If .HydrodynamicLines = "Average" Then SRBonus = SRBonus - 1
If .HydrodynamicLines = "Fine" Then SRBonus = SRBonus - 2
If .HydrodynamicLines = "Very Fine" Then SRBonus = SRBonus - 2
If .SubmarineLines Then SRBonus = SRBonus - 2
If .Catamaran Then SRBonus = SRBonus + 2
If .Trimaran Then SRBonus = SRBonus + 2
End With

'set the final value for stability
TempSR = TempSR + SRBonus
If TempSR < 1 Then TempSR = 1
wStability = TempSR
End Sub

Function CalcGroundSpeed() As Double
Dim Animal As Boolean
Dim SlowestAnimal As Single
Dim TempSpeed As Double
'init Animal flag
Animal = False
'init the SlowestAnimal
SlowestAnimal = 0
Select Case TypeName(colVehicle.Item(PerformanceType))
    Case "clsWheel"
        For i = 1 To UBound(arrThruster) - 1
            'check for animals.  Max vehicle ground speed cant exceed slowest animal
            If arrThruster(i).ClassName = "clsAnimal" Then
                Animal = True
                ' set the first animal to the slowest
                If SlowestAnimal = 0 Then SlowestAnimal = colVehicle(arrThruster(i).Key).Speed
                ' check if this animal is slower than the current slowest
                If colVehicle(arrThruster(i).Key).Speed < SlowestAnimal Then SlowestAnimal = colVehicle(arrThruster(i).Key).Speed
            End If
        Next
            ' add motive powers for all other types of thrust components
            Select Case arrThruster(i).ClassName
                Case "clsWheeledDrivetrain", "clsAnimal"
                    TotalMotivePower = TotalMotivePower + colVehicle(arrThruster(i).Key).MotivePower
                Case Else
                    TotalMotivePower = TotalMotivePower + (colVehicle(arrThruster(i).Key).MotivePower / 4)
            End Select
        If colVehicle("1_").LoadedMass = 0 Then
            TempSpeed = 0
        Else
            TempSpeed = Sqr(TotalMotivePower / colVehicle("1_").LoadedMass)
        End If
        TempSpeed = TempSpeed * gSpeedFactor
    Case "clsSkid"
        For i = 1 To UBound(arrThruster)
            'check for animals.  Max vehicle ground speed cant exceed slowest animal
            If arrThruster(i).ClassName = "clsAnimal" Then
                Animal = True
                ' set the first animal to the slowest
                If SlowestAnimal = 0 Then SlowestAnimal = colVehicle(arrThruster(i).Key).Speed
                ' check if this animal is slower than the current slowest
                If colVehicle(arrThruster(i).Key).Speed < SlowestAnimal Then SlowestAnimal = colVehicle(arrThruster(i).Key).Speed
            End If
        Next
            ' add motive powers for all other types of thrust components
            Select Case arrThruster(i).ClassName
                Case "clsAnimal"
                    TotalMotivePower = TotalMotivePower + colVehicle(arrThruster(i).Key).MotivePower
                Case Else
                    TotalMotivePower = TotalMotivePower + (colVehicle(arrThruster(i).Key).MotivePower / 4)
            End Select
        TempSpeed = Sqr(TotalMotivePower / colVehicle("1_").LoadedMass)
        TempSpeed = TempSpeed * gSpeedFactor
    Case "clsTrack"
        For i = 1 To UBound(arrThruster)
            ' get the trackeddrivetrain's motive power
            If arrThruster(i).ClassName = "clsTrackedDrivetrain" Then
                TempSpeed = Sqr(colVehicle(arrThruster(i).Key).MotivePower / colVehicle("1_").LoadedMass)
                TempSpeed = TempSpeed * gSpeedFactor
            End If
        Next
    Case "clsLeg"
        For i = 1 To UBound(arrThruster)
            ' get the trackeddrivetrain's motive power
            If arrThruster(i).ClassName = "clsLegDrivetrain" Then
                TempSpeed = Sqr(colVehicle(arrThruster(i).Key).MotivePower / colVehicle("1_").LoadedMass)
                TempSpeed = TempSpeed * gSpeedFactor
            End If
        Next
    Case "Flexibody__"
        For i = 1 To UBound(arrThruster)
            ' get the trackeddrivetrain's motive power
            If arrThruster(i).ClassName = "clsFlexibodyDrivetrain" Then
                TempSpeed = Sqr(colVehicle(arrThruster(i).Key).MotivePower / colVehicle("1_").LoadedMass)
                TempSpeed = TempSpeed * gSpeedFactor
            End If
        Next
End Select
' add in the streamlining effects
With colVehicle("1_")
If TempSpeed > 50 Then
    If .StreamLining = "None" Then
        If TempSpeed > 600 Then TempSpeed = 600
    ElseIf .StreamLining = "Fair" Then
        TempSpeed = TempSpeed + (TempSpeed * 0.05)
        If TempSpeed > 600 Then TempSpeed = 600
    Else
        TempSpeed = TempSpeed + (TempSpeed * 0.1)
    End If
End If
End With
' Check to make sure we dont exceed speed of slowest animal(if applicable)
If SlowestAnimal = 0 Then ' do nothing
ElseIf TempSpeed < SlowestAnimal Then ' do nothing
Else ' cant exceed slowest animal!
    TempSpeed = SlowestAnimal
End If
' Do the final rounding
If TempSpeed >= 20 Then
    ' round to nearest 5mph
    CalcGroundSpeed = CLng(TempSpeed / 5) * 5
Else ' round to nearest whole number
    CalcGroundSpeed = CLng(TempSpeed)
End If
End Function

Function CalcGroundSpeedFactor()
Dim TempKey As String
Dim Bonus As Integer
Dim RailBonus As Boolean
Dim LegCount As Integer
Dim Element As Object

TempKey = PerformanceType

'find the improved suspension bonus
With colVehicle.Item(TempKey)
    If .ImprovedSuspension = True Then
        If TypeName(colVehicle.Item(TempKey)) = "clsWheel" Then
            Bonus = 2
        Else
            Bonus = 1
        End If
    End If

'Determine if there is a railway bonus
    If TypeName(colVehicle.Item(TempKey)) = "clsWheel" Then
        If .SubType = "Railway Wheels" Then
            If gVehicleTL >= 5 Then
                RailBonus = True
            Else
                RailBonus = False
            End If
        End If
    End If
' get the Final total speed factor
Select Case TypeName(colVehicle.Item(TempKey))
    Case "clsSkid"
        CalcGroundSpeedFactor = 6 + Bonus
        Exit Function
    Case "clsWheel"
        If gVehicleTL <= 4 Then
            CalcGroundSpeedFactor = 8 + Bonus
            Exit Function
        ElseIf gVehicleTL = 5 Then
            If RailBonus = True Then
                CalcGroundSpeedFactor = 2 * (12 + Bonus)
                Exit Function
            Else
                CalcGroundSpeedFactor = 12 + Bonus
                Exit Function
            End If
        ElseIf gVehicleTL >= 6 Then
            If RailBonus = True Then
                CalcGroundSpeedFactor = 2 * (16 + Bonus)
                Exit Function
            Else
                CalcGroundSpeedFactor = 16 + Bonus
                Exit Function
            End If
        End If
    Case "clsTrack"
        'TODO make sure the SubTypes are correct!!!!
        If .SubType = "Tracks" Then
            If gVehicleTL = 6 Then
                CalcGroundSpeedFactor = 10 + Bonus
                Exit Function
            Else
                CalcGroundSpeedFactor = 12 + Bonus
                Exit Function
            End If
        ElseIf .SubType = "Skitracks" Then
            If gVehicleTL = 6 Then
                CalcGroundSpeedFactor = 8 + Bonus
                Exit Function
            Else
                CalcGroundSpeedFactor = 10 + Bonus
                Exit Function
            End If
        ElseIf .SubType = "Halftracks" Then
            If gVehicleTL = 6 Then
                CalcGroundSpeedFactor = 12 + Bonus
                Exit Function
            Else
                CalcGroundSpeedFactor = 14 + Bonus
                Exit Function
            End If
        Else
            MsgBox "Subtypes wrong in clsPerformance.CalcGroundSpeedFactor"
        End If
    Case "clsLeg"
        ' determine how many legs are on the vehicle
        LegCount = 0 'init the counter
        For Each Element In colVehicle
            If TypeName(Element) = "clsLeg" Then
                LegCount = LegCount + 1
            End If
        Next
        If LegCount = 1 Then
            MsgBox "user hasnt added at least 2 legs!! Using _1_ for Ground Speed Factor"
            CalcGroundSpeedFactor = 1
            Exit Function
        ElseIf LegCount = 2 Then
            CalcGroundSpeedFactor = 8 + Bonus
            Exit Function
        ElseIf LegCount = 3 Then
            CalcGroundSpeedFactor = 10 + Bonus
            Exit Function
        ElseIf LegCount >= 4 Then
            CalcGroundSpeedFactor = 12 + Bonus
            Exit Function
        End If
    End Select
End With
        
End Function

Function CalcGroundAcceleration(Speedfactor As Integer, TopSpeed As Long)
Dim TempAcceleration As Long
' calculates maximum acceleration
TempAcceleration = (TopSpeed * Speedfactor) * 0.8
' round to nearest 5 miles per hour unless under 5 then just nearest whole number
If TempAcceleration < 5 Then ' round to nearest 1mph
    TempAcceleration = CInt(TempAcceleration)
Else ' round to nearest 5mph
    TempAcceleration = CLng(TempAcceleration * 5) \ 5
End If
' TODO there is nasty exception involving Vehicles with Legs
CalcGroundAcceleration = TempAcceleration
End Function

Public Property Let PerformanceType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PerformanceType = 5
    mvarPerformanceType = vData
End Property


Public Property Get PerformanceType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PerformanceType
    PerformanceType = mvarPerformanceType
End Property

Public Property Let Key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property
Public Property Let wTopSpeed(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wTopSpeed = 5
    mvarwTopSpeed = vData
End Property


Public Property Get wTopSpeed() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wTopSpeed
    wTopSpeed = mvarwTopSpeed
End Property

Public Property Let wHydrofoilSpeed(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wHydrofoilSpeed = 5
    mvarwHydrofoilSpeed = vData
End Property


Public Property Get wHydrofoilSpeed() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wHydrofoilSpeed
    wHydrofoilSpeed = mvarwHydrofoilSpeed
End Property

Public Property Let wPlaningSpeed(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wPlaningSpeed = 5
    mvarwPlaningSpeed = vData
End Property


Public Property Get wPlaningSpeed() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wPlaningSpeed
    wPlaningSpeed = mvarwPlaningSpeed
End Property


Public Property Let wStability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wStability = 5
    mvarwStability = vData
End Property

Public Property Get wStability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wStability
    wStability = mvarwStability
End Property

Public Property Let wManeuverability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wManeuverability = 5
    mvarwManeuverability = vData
End Property


Public Property Get wManeuverability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wManeuverability
    wManeuverability = mvarwManeuverability
End Property

Public Property Let wHydroDrag(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wHydroDrag = 5
    mvarwHydroDrag = vData
End Property

Public Property Get wHydroDrag() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wHydroDrag
    wHydroDrag = mvarwHydroDrag
End Property

Public Property Let wDraft(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wDraft = 5
    mvarwDraft = vData
End Property

Public Property Get wDraft() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wDraft
    wDraft = mvarwDraft
End Property

Public Property Let wDeceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wDeceleration = 5
    mvarwDeceleration = vData
End Property

Public Property Get wDeceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wDeceleration
    wDeceleration = mvarwDeceleration
End Property
Public Property Let wIDeceleration(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wIDeceleration = 5
    mvarwIDeceleration = vData
End Property

Public Property Get wIDeceleration() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wIDeceleration
    wIDeceleration = mvarwIDeceleration
End Property
Public Property Let wAcceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wAcceleration = 5
    mvarwAcceleration = vData
End Property

Public Property Get wAcceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wAcceleration
    wAcceleration = mvarwAcceleration
End Property

Public Property Let wTotalAquaticThrust(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wTotalAquaticThrust = 5
    mvarwTotalAquaticThrust = vData
End Property

Public Property Get wTotalAquaticThrust() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wTotalAquaticThrust
    wTotalAquaticThrust = mvarwTotalAquaticThrust
End Property
Public Property Let suTopSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suTopSpeed = 5
    mvarsuTopSpeed = vData
End Property

Public Property Get suTopSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suTopSpeed
    suTopSpeed = mvarsuTopSpeed
End Property

Public Property Let suTotalAquaticThrust(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suTotalAquaticThrust = 5
    mvarsuTotalAquaticThrust = vData
End Property

Public Property Get suTotalAquaticThrust() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suDeceleration
    suTotalAquaticThrust = mvarsuTotalAquaticThrust
End Property

Public Property Let suStability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suStability = 5
    mvarsuStability = vData
End Property

Public Property Get suStability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suStability
    suStability = mvarsuStability
End Property


Public Property Let suManeuverability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suManeuverability = 5
    mvarsuManeuverability = vData
End Property

Public Property Get suManeuverability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suManeuverability
    suManeuverability = mvarsuManeuverability
End Property

Public Property Let suHydroDrag(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suHydroDrag = 5
    mvarsuHydroDrag = vData
End Property

Public Property Get suHydroDrag() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suHydroDrag
    suHydroDrag = mvarsuHydroDrag
End Property

Public Property Let suDeceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suDeceleration = 5
    mvarsuDeceleration = vData
End Property

Public Property Get suDeceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suDeceleration
    suDeceleration = mvarsuDeceleration
End Property

Public Property Let suIDeceleration(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suIDeceleration = 5
    mvarsuIDeceleration = vData
End Property

Public Property Get suIDeceleration() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suIDeceleration
    suIDeceleration = mvarsuIDeceleration
End Property

Public Property Let suCrushDepth(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suCrushDepth = 5
    mvarsuCrushDepth = vData
End Property

Public Property Get suCrushDepth() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suCrushDepth
    suCrushDepth = mvarsuCrushDepth
End Property

Public Property Let suAcceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suAcceleration = 5
    mvarsuAcceleration = vData
End Property

Public Property Get suAcceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suAcceleration
    suAcceleration = mvarsuAcceleration
End Property

Public Property Let sTopSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTopSpeed = 5
    mvarsTopSpeed = vData
End Property

Public Property Get sTopSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTopSpeed
    sTopSpeed = mvarsTopSpeed
End Property

Public Property Let sStability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sStability = 5
    mvarsStability = vData
End Property

Public Property Get sStability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sStability
    sStability = mvarsStability
End Property

Public Property Let sManeuverability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sManeuverability = 5
    mvarsManeuverability = vData
End Property

Public Property Get sManeuverability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sManeuverability
    sManeuverability = mvarsManeuverability
End Property

Public Property Let sFTLSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFTLSpeed = 5
    mvarsFTLSpeed = vData
End Property

Public Property Get sFTLSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFTLSpeed
    sFTLSpeed = mvarsFTLSpeed
End Property

Public Property Let hDeceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hDeceleration = 5
    mvarhDeceleration = vData
End Property

Public Property Get hDeceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hDeceleration
    hDeceleration = mvarhDeceleration
End Property

Public Property Let hAcceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hAcceleration = 5
    mvarhAcceleration = vData
End Property

Public Property Get hAcceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hAcceleration
    hAcceleration = mvarhAcceleration
End Property

Public Property Let gSpeedFactor(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gSpeedFactor = 5
Dim Bonus As Single
Dim LegCount As Integer
    mvargSpeedFactor = vData
'///////////////////////////////////////////////////
' Get the Ground Acceleration but first check
' for the Leg Exception rule
If TypeName(colVehicle.Item(PerformanceType)) = "clsLeg" Then
   ' determine how many legs are on the vehicle
    LegCount = 0 'init the counter
    For Each Element In colVehicle
        If TypeName(Element) = "clsLeg" Then
            LegCount = LegCount + 1
        End If
        If LegCount = 2 Then
            Bonus = 12
        ElseIf LegCount = 3 Then
            Bonus = 9.6
        ElseIf LegCount >= 4 Then
            Bonus = 8
        End If
    Next
    TempAcceleration = ((mvargTopSpeed / mvargSpeedFactor) * 0.8) + Bonus
Else ' there are no legs. Use standard Ground Acceleration formula
    TempAcceleration = (mvargTopSpeed / mvargSpeedFactor) * 0.8
End If
' Do the rounding
If TempAcceleration > 5 Then
    mvargAcceleration = CLng(TempAcceleration / 5) * 5 'to nearest 5mph
Else
    mvargAcceleration = CLng(TempAcceleration) 'to nearest 1mph
End If
'//////////////////////////////////////////////////////
'now get the Ground Deceleration
If TypeName(colVehicle.Item(PerformanceType)) = "clsTrack" Then
    'check for improved breaks with smartwheels
    If colVehicle.Item(PerformanceType).ImprovedBreaks = True Then
        mvargDeceleration = 15
        If colVehicle.Item(PerformanceType).Smartwheels = True Then
            mvargDeceleration = 20
        End If
    'check for just smartwheels
    ElseIf colVehicle.Item(PerformanceType).Smartwheels = True Then
        mvargDeceleration = 15
    Else ' neither smartwheels or improved found
        mvargDeceleration = 10
    End If
ElseIf TypeName(colVehicle.Item(PerformanceType)) = "clsTrack" Then
    If colVehicle.Item(PerformanceType).SubType = "Skitracks" Then
        mvargDeceleration = 15
    Else
        mvargDeceleration = 20
    End If
ElseIf TypeName(colVehicle.Item(PerformanceType)) = "clsSkid" Then
    mvargDeceleration = 5
Else: mvargDeceleration = 20 'covers legs and flexibody
End If

'///////////////////////////////////////////////////////////
'now get the GroundStability and Manuever Ratings

Dim TempVolume As Double
TempVolume = colVehicle("1_").Volume
' Determine which Motive System to use
'1 = skid
'2 = Wheels 1
'3 = Wheels 2
'4 = Wheels 3
'5 = wheels 4-7
'6 = Wheels 8+
'7 = Tracks
'8 = Skitracks
'9 = Halftracks
'10 = 2 legs
'11 = 3 legs
'12 = 4+ legs
'13 = Flexibody
If TypeName(colVehicle.Item(PerformanceType)) = "clsSkid" Then
    MotiveSystem = 1
ElseIf TypeName(colVehicle.Item(PerformanceType)) = "clsWheel" Then
    With colVehicle.Item(PerformanceType)
        If .Number = 1 Then
            MotiveSystem = 2
        ElseIf .Number = 2 Then
            MotiveSystem = 3
        ElseIf .Number = 3 Then
            MotiveSystem = 4
        ElseIf .Number <= 6 Then
            MotiveSystem = 5
        Else: MotiveSystem = 6
        End If
    End With
ElseIf TypeName(colVehicle.Item(PerformanceType)) = "clsTrack" Then
    With colVehicle.Item(PerformanceType)
        If .SubType = "Tracks" Then
            MotiveSystem = 7
        ElseIf .SubType = "Skitracks" Then
            MotiveSystem = 8
        Else: MotiveSystem = 9
        End If
    End With
ElseIf TypeName(colVehicle.Item(PerformanceType)) = "clsLeg" Then
    LegCount = 0 'init the legcount
    'TODOD i might be able to delete this legcount and use the above one
    'but only IF i keep all ground performance calcs in one routine
    For Each Element In colVehicle
        If TypeName(Element) = "clsLeg" Then
            LegCount = LegCount + 1
        End If
    Next
    If LegCount = 2 Then
        MotiveSystem = 10
    ElseIf LegCount = 3 Then
        MotiveSystem = 11
    ElseIf LegCount >= 4 Then
        MotiveSystem = 12
    End If
Else: MotiveSystem = 13 'flexibody
End If
' get the actual values from the table
i = MotiveSystem
If TempVolume <= 30 Then
    TempMR = GroundStabMatrix(i).M1
    TempSR = GroundStabMatrix(i).S1
ElseIf TempVolume <= 100 Then
    TempMR = GroundStabMatrix(i).M2
    TempSR = GroundStabMatrix(i).S2
ElseIf TempVolume <= 300 Then
    TempMR = GroundStabMatrix(i).M3
    TempSR = GroundStabMatrix(i).S3
ElseIf TempVolume <= 3000 Then
    TempMR = GroundStabMatrix(i).M4
    TempSR = GroundStabMatrix(i).S4
Else
    TempMR = GroundStabMatrix(i).M5
    TempSR = GroundStabMatrix(i).S5
End If
'//////////////////////////////////////////////////////
' apply the gMR and gSR modifiers
If colVehicle.Item(PerformanceType).ImprovedSuspension = True Then
    TempSR = TempSR + 1 'increase the gSR
    If TempMR = 0.125 Then
        TempMR = 0.25
    Else
        TempMR = TempMR + 0.25
    End If
End If

For Each Element In colVehicle
        If TypeName(Element) = "clsManeuverControl" Then
            If Element.SubType = "Electronic" Then
                If TempMR = 0.125 Then
                    TempMR = 0.25
                Else
                    TempMR = TempMR + 0.25
                End If
            ElseIf Element.SubType = "Computerized" Then
                If TempMR = 0.125 Then
                    TempMR = 0.25
                Else
                    TempMR = TempMR + 0.25
                End If
            End If
        Exit For 'stop after first maneuver control is found
        End If
Next
'add modifier for responsive structure.  NOTE: if one subassembly has
'responsive structure, they all must!
If colVehicle.Item(PerformanceType).Responsive = True Then
    If TempMR = 0.125 Then
       TempMR = 0.25
    Else
        TempMR = TempMR + 0.25
    End If
End If
'add modifiers for smartwheels and all wheel steering
With colVehicle.Item(PerformanceType)
If TypeName(colVehicle.Item(PerformanceType)) = "clsWheel" Then
    If .Smartwheels = True Then
        If TempMR = 0.125 Then
            TempMR = 0.25
        Else
            TempMR = TempMR + 0.25
        End If
    End If
    If .AllwheelSteering = True Then
        If TempMR = 0.125 Then
            TempMR = 0.25
        Else
            TempMR = TempMR + 0.25
        End If
    End If
    If .SubType = "Small Wheels" Then
        TempSR = TempSR + 1
        If TempMR > 0.5 Then TempMR = 0.5 ' vehicle with small wheels cant exceed .5 MR
    ElseIf .SubType = "Railway Wheels" Then
        If TempMR > 0.5 Then TempMR = 0.5 ' vehicle with railway wheels cant exceed .5 MR
    End If
' add modifiers for legs built at or prior to TL 7
ElseIf TypeName(colVehicle.Item(PerformanceType)) = "clsLeg" Then
    If gVehicleTL <= 7 Then TempSR = TempSR - 1
End If
End With
' conduct final exception checks for harnessed animals, sails or non-folded wings and rotars
For Each Element In colVehicle
    If TypeName(Element) = "clsSail" Then
        If TempMR > 0.5 Then TempMR = 0.5 'vehicle with sails cant exceed .5 MR
    ElseIf TypeName(Element) = "clsAnimal" Then
        If TempMR > 0.5 Then TempMR = 0.5 'vehicle with harnessed animals cant exceed .5MR
    ElseIf TypeName(Element) = "clsWing" Then
        If Element.Folding <> True Then
            If TempMR > 0.5 Then TempMR = 0.5 'vehicles with non folded wings limited to .5MR
        End If
    ElseIf TypeName(Element) = "clsRotar" Then
        If Element.Folding <> True Then
            If TempMR > 0.5 Then TempMR = 0.5 'vehicles with non folded rotars limited to .5MR
        End If
    ElseIf TypeName(Element) = "clsCrewStation" Then
        If Element.SubType = "Harness" Then
            TempSR = TempSR - 1 'vehicles worn as a harness are SR - 1
        End If
    End If
Next
'////////////////////////////////////////////////////////////
'' get the contact area
Dim ContactArea As Double
Dim TempContactArea As Double

If TypeName(colVehicle.Item(PerformanceType)) = "clsFlexibodyDrivetrain" Then
    ContactArea = colVehicle("1_").SurfaceArea / 6
ElseIf TypeName(colVehicle.Item(PerformanceType)) = "clsLeg" Then
    For Each Element In colVehicle
        If TypeOf Element Is clsLeg Then
            TempContactArea = TempContactArea + Element.SurfaceArea
        End If
    Next
    ContactArea = TempContactArea / 12.5
ElseIf TypeName(colVehicle.Item(PerformanceType)) = "clsSkid" Then
    ContactArea = colVehicle.Item(PerformanceType).SurfaceArea / 10
ElseIf TypeName(colVehicle.Item(PerformanceType)) = "clsTrack" Then
    If colVehicle.Item(PerformanceType).SubType = "Skitracks" Then
        ContactArea = colVehicle.Item(PerformanceType).SurfaceArea / 10
    ElseIf colVehicle.Item(PerformanceType).SubType = "Halftracks" Then
        ContactArea = colVehicle.Item(PerformanceType).SurfaceArea / 20
    ElseIf colVehicle.Item(PerformanceType).SubType = "Tracks" Then
        ContactArea = colVehicle.Item(PerformanceType).SurfaceArea / 5
    End If
ElseIf TypeName(colVehicle.Item(PerformanceType)) = "clsWheel" Then
    If colVehicle.Item(PerformanceType).SubType = "Railway Wheels" Then
        ContactArea = colVehicle.Item(PerformanceType).SurfaceArea / 66
    ElseIf colVehicle.Item(PerformanceType).SubType = "Off-road Wheels" Then
        ContactArea = colVehicle.Item(PerformanceType).SurfaceArea / 33
    Else
        If gVehicleTL >= 6 Then
            ContactArea = colVehicle.Item(PerformanceType).SurfaceArea / 50
        Else
            ContactArea = colVehicle.Item(PerformanceType).SurfaceArea / 66
        End If
    End If
End If

' get the ground pressure
Dim TempWeight As Double
Dim MinWeight As Double ' contragrav cant reduce weight to less than 10% of original
Dim sDescription As String
Dim TempPressure As Double
Dim First As Integer
Dim Second As Integer
Dim arrPT(7, 28)
'fill the pressure table
arrPT(1, 1) = 100
arrPT(1, 2) = 100
arrPT(1, 3) = 4 / 5
arrPT(1, 4) = 2 / 3
arrPT(2, 1) = 100
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

TempWeight = colVehicle("1_").HLoadedWeight
MinWeight = TempWeight * 0.1
TempWeight = TempWeight - CalcTotalContragravLift
If TempWeight < MinWeight Then TempWeight = MinWeight
gPressure = TempWeight

If TempWeight <= 150 Then
    First = 1
ElseIf TempWeight <= 900 Then First = 2
ElseIf TempWeight <= 1800 Then First = 3
ElseIf TempWeight <= 2700 Then First = 4
ElseIf TempWeight <= 7500 Then First = 5
ElseIf TempWeight <= 15000 Then First = 6
Else: First = 7
End If

Select Case TypeName(colVehicle.Item(PerformanceType))
    Case "clsLeg", "clsFlexibodyDrivetrain"
        Second = 1
    Case "clsTrack" ' but not Skitracks
        If colVehicle.Item(PerformanceType).SubType = "Tracks" Then
            Second = 2
        Else: Second = 3
        End If
    Case "clsSkid"
        Second = 4
    Case "clsWheel"
        'TODO check ruling about vehicles with small wheels. See page 130!
        'if small wheels do count, then the below code is complete.
        'or if small wheels count but have reduced speed, i must add that code
        'TODO make sure property allwheeldrive is being flagged properly in the clsWheel by the
        ' clsWheeleddrivtrain child
        If colVehicle.Item(PerformanceType).AllWheelDrive = True Then
            Second = 3
        Else: Second = 4
        End If
End Select
' get off road speed
If Second = 5 Then
    gOffRoad = 0 ' if it has railway wheels
Else
    gOffRoad = gTopSpeed * arrPT(First, Second)
End If

End Property

Public Property Get gSpeedFactor() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gSpeedFactor
    gSpeedFactor = mvargSpeedFactor
End Property

Public Property Let gTopSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gTopSpeed = 5
    mvargTopSpeed = vData
End Property

Public Property Get gTopSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gTopSpeed
    gTopSpeed = mvargTopSpeed
End Property

Public Property Let gStability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gStability = 5
    mvargStability = vData
End Property

Public Property Get gStability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gStability
    gStability = mvargStability
End Property

Public Property Let gPressure(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gPressure = 5
    mvargPressure = vData
End Property


Public Property Get gPressure() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gPressure
    gPressure = mvargPressure
End Property



Public Property Let gOffRoad(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gOffRoad = 5
    mvargOffRoad = vData
End Property


Public Property Get gOffRoad() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gOffRoad
    gOffRoad = mvargOffRoad
End Property



Public Property Let gManeuverability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gManeuverability = 5
    mvargManeuverability = vData
End Property


Public Property Get gManeuverability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gManeuverability
    gManeuverability = mvargManeuverability
End Property



Public Property Let gDeceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gDeceleration = 5
    mvargDeceleration = vData
End Property


Public Property Get gDeceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gDeceleration
    gDeceleration = mvargDeceleration
End Property



Public Property Let gAcceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gAcceleration = 5
    mvargAcceleration = vData
End Property


Public Property Get gAcceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gAcceleration
    gAcceleration = mvargAcceleration
End Property



Public Property Let aTopSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aTopSpeed = 5
    mvaraTopSpeed = vData
End Property


Public Property Get aTopSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aTopSpeed
    aTopSpeed = mvaraTopSpeed
End Property



Public Property Let aStallSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aStallSpeed = 5
    mvaraStallSpeed = vData
End Property


Public Property Get aStallSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aStallSpeed
    aStallSpeed = mvaraStallSpeed
End Property



Public Property Let aStability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aStability = 5
    mvaraStability = vData
End Property


Public Property Get aStability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aStability
    aStability = mvaraStability
End Property



Public Property Let aMotiveThrust(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aMotiveThrust = 5
    mvaraMotiveThrust = vData
End Property


Public Property Get aMotiveThrust() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aMotiveThrust
    aMotiveThrust = mvaraMotiveThrust
End Property



Public Property Let aManeuverability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aManeuverability = 5
    mvaraManeuverability = vData
End Property


Public Property Get aManeuverability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aManeuverability
    aManeuverability = mvaraManeuverability
End Property



Public Property Let aDrag(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aDrag = 5
    mvaraDrag = vData
End Property


Public Property Get aDrag() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aDrag
    aDrag = mvaraDrag
End Property



Public Property Let aDeceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aDeceleration = 5
    mvaraDeceleration = vData
End Property


Public Property Get aDeceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aDeceleration
    aDeceleration = mvaraDeceleration
End Property



Public Property Let aAcceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aAcceleration = 5
    mvaraAcceleration = vData
End Property


Public Property Get aAcceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aAcceleration
    aAcceleration = mvaraAcceleration
End Property



Private Sub Class_Initialize()
ReDim arrThruster(1)

End Sub
