Attribute VB_Name = "modStatistics"
Option Explicit
' This module contains all the math functions needed to calculate vehicle statistics

Function CalcBodyVolume() As Double

Dim SlopeMultiplier As Single ' vehicles total slope Volume multiplier
Dim RetractsCost As Single ' Rectractable Wheels or Skids Volume Multiplier
Dim StreamLiningCost As Single ' Streamlined Hull Volume Multiplier
Dim HydrodynamicHullCost As Single ' total of all Special Structure Cost Modifiers
Dim TotalOtherCost As Single ' total of all Other Volume Modifiers
Dim objBody As clsBody ' holds the body class for the vehicle
Dim sStreamlining As String ' vehicle body's streamlining property
Dim sHydrodynamicHull As String ' vehicle body's hydrodynamichulllines property
Dim sRetracts As String ' vehicle body's rectract property
Dim bSubmersible As Boolean ' vehicle's submersible property
Dim sCatorTrimaran As String ' vheicles hydrodynamic hull option type
Const CatorTrimaranCost = 1.3 ' Catamaran or Trimaran Hull Volume Multiplier
Const SubmersibleCost = 1.25 ' Submersible Hull Volume Multiplier

Set objBody = colVehicle.Item("1_")
With objBody
    sStreamlining = .StreamLining
    sRetracts = .RetractableWheelsSkids
    sHydrodynamicHull = .HydrodynamicLines
    bSubmersible = .Submersible
    If .Trimaran = True Then
        sCatorTrimaran = "Trimaran"
    ElseIf .Catamaran = True Then
        sCatorTrimaran = "Catamaran"
    Else
        sCatorTrimaran = "None"
    End If
End With
objBody.CombinedComponentVolume = 0#  ' must reset the variable back to 0
' Calculate the combined component volume
objBody.CombinedComponentVolume = CalcCombinedVolume("1_")
' Add any turret rotationspace if it exists
objBody.CombinedComponentVolume = objBody.CombinedComponentVolume + CalcRotationSpace("1_")
' Determine if the minimum volume (due to masts) is achieved
If objBody.MinimumVolume = 0 Then
Else
    If objBody.MinimumVolume > objBody.CombinedComponentVolume Then
        MsgBox "Warning:  Body total comoponent volume must be greater than or equal to " & objBody.MinimumVolume & " [Tallest Mast Height /4) cubed].  You must add at least " & objBody.MinimumVolume - objBody.CombinedComponentVolume
    End If
End If
    
' Calculate the Real Body Volume
Select Case sStreamlining
    Case "None"
        StreamLiningCost = 1
    Case "Fair"
        StreamLiningCost = 1.1
    Case "Good"
        StreamLiningCost = 1.2
    Case "Very Good"
        StreamLiningCost = 1.25
    Case "Superior"
        StreamLiningCost = 1.3
    Case "Excellent"
        StreamLiningCost = 1.35
    Case "Radical"
        StreamLiningCost = 1.4
End Select

Select Case sHydrodynamicHull
    Case "None"
        HydrodynamicHullCost = 1
    Case "Mediocre"
        HydrodynamicHullCost = 1.1
    Case "Average"
        HydrodynamicHullCost = 1.2
    Case "Submarine"
        HydrodynamicHullCost = 1.2
    Case "Fine"
        HydrodynamicHullCost = 1.3
    Case "Very Fine"
        HydrodynamicHullCost = 1.3
End Select

Select Case sRetracts
    Case "None"
        RetractsCost = 1
    Case "Body"
        RetractsCost = 1.075
    Case "Body & Wings"
        RetractsCost = 1.025
End Select

' calculate the Multiplier needed for the slopes applied to the Assembly
SlopeMultiplier = CalcSlopeMultiplier("1_")
' Calculate Volume modifier value for Other Modifiers
TotalOtherCost = 1 ' initialize variable
If bSubmersible = True Then TotalOtherCost = SubmersibleCost
If sCatorTrimaran <> "None" Then TotalOtherCost = TotalOtherCost * CatorTrimaranCost
' Calculate Final Volume for the Body
CalcBodyVolume = objBody.CombinedComponentVolume * TotalOtherCost * StreamLiningCost * HydrodynamicHullCost * SlopeMultiplier * RetractsCost
End Function
Function CalcCombinedVolume(ObjectKey As String) As Double
Dim i As Integer
Dim NumObjects As Integer
NumObjects = colVehicle.Count
' add up volumes of all objects attached to the subassembly
For i = 1 To NumObjects
        If TypeName(colVehicle.Item(i)) = "clsPerformance" Then
        ElseIf colVehicle.Item(i).Parent = ObjectKey Then
            If TypeOf colVehicle.Item(i) Is clsPopTurret Then
            CalcCombinedVolume = CalcCombinedVolume + colVehicle.Item(i).RotationSpace
            ElseIf TypeOf colVehicle.Item(i) Is clsTurret Then
            CalcCombinedVolume = CalcCombinedVolume + colVehicle.Item(i).RotationSpace
            Else
            CalcCombinedVolume = CalcCombinedVolume + colVehicle.Item(i).Volume
            End If
        End If
Next
End Function
Function CalcRotationSpace(sKey As String) As Double
Dim i As Integer
Dim NumObjects As Integer
NumObjects = colVehicle.Count
' add up volumes of all objects attached to the subassembly
For i = 1 To NumObjects
        If TypeName(colVehicle.Item(i)) = "clsPerformance" Then
        ElseIf colVehicle.Item(i).Parent = sKey Then
            If TypeOf colVehicle.Item(i) Is clsTurret Then
                CalcRotationSpace = CalcRotationSpace + Val(colVehicle.Item(i).RotationSpace)
            ElseIf TypeOf colVehicle.Item(i) Is clsPopTurret Then
                CalcRotationSpace = CalcRotationSpace + Val(colVehicle.Item(i).RotationSpace)
            End If
        End If
Next
End Function

Function CalcSlopeMultiplier(sKey As String) As Single
Dim Temp As Integer

Temp = Val(colVehicle(sKey).SlopeF) + Val(colVehicle(sKey).SlopeB) + Val(colVehicle(sKey).SlopeL) + Val(colVehicle(sKey).SlopeR)

Select Case Temp
    Case 0
        CalcSlopeMultiplier = 1
    Case 30
        CalcSlopeMultiplier = 1.1
    Case 60
        CalcSlopeMultiplier = 1.25
    Case 90
        CalcSlopeMultiplier = 1.4
    Case 120
        CalcSlopeMultiplier = 1.6
    Case 150
        CalcSlopeMultiplier = 2
    Case 180
        CalcSlopeMultiplier = 2.5
    Case 210
        CalcSlopeMultiplier = 3.3
    Case 240
        CalcSlopeMultiplier = 5
End Select
    
End Function

Function CalcSurfaceArea(Volume As Double) As Double
CalcSurfaceArea = ((Volume ^ (1 / 3)) ^ 2) * 6
End Function

Function TechLevelModifier(TechLevel As Integer)

Select Case TechLevel
    Case Is <= 5
        TechLevelModifier = 12
    Case 6
        TechLevelModifier = 8
    Case 7
        TechLevelModifier = 6
    Case 8
        TechLevelModifier = 4
    Case 9
        TechLevelModifier = 3
    Case 10
        TechLevelModifier = 2
    Case 11
        TechLevelModifier = 1.5
    Case Is >= 12
        TechLevelModifier = 1
End Select
End Function

Sub CalcStructuralStats()
Dim TempArea As Double
Dim TempTotalArea As Double
Dim TempCost As Double
Dim TempWeight As Double
Dim element As Object
' Calculate the Total Structural and Total Surface Area Weight of the Vehicle (not just the body!!)
For Each element In colVehicle
    Select Case TypeName(element)
        Case "clsBody"
            element.StatsUpdate
            TempArea = element.SurfaceArea + TempArea
            TempWeight = element.Weight + TempWeight
            TempCost = element.Cost + TempCost
        Case "clsMast", "clsGasbag", "clsOpenMount"
            TempTotalArea = element.SurfaceArea + TempTotalArea
    End Select
Next
With colVehicle("1_")
    .StructuralSurfaceArea = TempArea
    .TotalSurfaceArea = TempArea + TempTotalArea
    .StructuralCost = TempCost
    .StructuralWeight = TempWeight
End With
End Sub


Function BasicDesignCost(SurfaceArea As Double, StructureTL As Integer, sStrength As String, sMaterials As String, bResponsive As Boolean, bRobotic As Boolean, bBiomechanical As Boolean, bLivingMetal As Boolean) As Double
'This produces structural COST and not BASIC cost!!!
'TODO might rename this function so that its not misleading later if i have time
Dim StructureCost As Integer ' Basic Design Cost
Dim StrengthCost As Single ' Frame Strength Cost Multiplier
Dim MaterialsCost As Single ' Materials Cost Multiplier
Dim StreamLiningCost As Single ' Streamlined Structure Cost Multiplier
Dim TotalSpecialCost As Single ' total of all Special Structure Cost Modifiers
Dim TotalOtherCost As Single ' total of all Other Cost Modifiers
Dim objBody As clsBody ' holds the body class for the vehicle
Dim sStreamlining As String ' vehicle body's streamlining property
Dim bSubmersible As Boolean ' vehicle body's submersible property
Dim bWingsorRotars As Boolean ' vehicle body's wingsorrotars property
Dim bLiftingBody As Boolean ' vehicle body's LiftingBody property
Dim bFlexibody As Boolean ' vehcicle body's Flexibody Drivetrain property

Const ResponsiveCost = 1.5 ' Responsive Structure Cost Multiplier
Const RoboticCost = 2 ' Robotic Structure Cost Multiplier
Const BiomechanicalCost = 1.5 ' Biomechanical Structure Cost Multiplier
Const LivingMetalCost = 2 ' Living Metal Structure Cost Multiplier
Const SubmersibleCost = 2 ' Submersible Structure Cost Multiplier
Const WingsorRotorsCost = 10 ' Wings or Rotors Cost Multiplier
Const LiftingBodyCost = 1.2 ' Lifting Body Cost Multiplier
Const FlexibodyDriveCost = 2 ' Flexibody Drive Train Cost Multiplier


Set objBody = colVehicle.Item("1_")

sStreamlining = objBody.StreamLining
bSubmersible = objBody.Submersible
bWingsorRotars = objBody.WingsorRotors
bLiftingBody = objBody.LiftingBody
bFlexibody = objBody.FlexibodyDrivetrain

Select Case StructureTL
    Case 0, 1
        StructureCost = 5
    Case 2, 3, 4
        StructureCost = 5
    Case 5
        StructureCost = 5
    Case 6
        StructureCost = 10
    Case 7
        StructureCost = 50
    Case 8
        StructureCost = 50
    Case 9
        StructureCost = 50
    Case 10
        StructureCost = 50
    Case 11
        StructureCost = 50
    Case Is >= 12
        StructureCost = 50
    Case Else
        StructureCost = 1
End Select

Select Case sStrength
    Case "Super-Light"
        StrengthCost = 0.1
    Case "Extra-Light"
        StrengthCost = 0.25
    Case "Light"
        StrengthCost = 0.5
    Case "Medium"
        StrengthCost = 1
    Case "Heavy"
        StrengthCost = 2
    Case "Extra-Heavy"
        StrengthCost = 5
    Case Else
        StrengthCost = 1
End Select

Select Case sMaterials
    Case "Very Cheap"
        MaterialsCost = 0.2
    Case "Cheap"
        MaterialsCost = 0.5
    Case "Standard"
        MaterialsCost = 1
    Case "Expensive"
        MaterialsCost = 2
    Case "Very Expensive"
        MaterialsCost = 5
    Case "Advanced"
        MaterialsCost = 10
    Case Else
        MaterialsCost = 1
End Select

Select Case sStreamlining
    Case "None"
        StreamLiningCost = 1
    Case "Fair"
        StreamLiningCost = 1.2
    Case "Good"
        StreamLiningCost = 1.5
    Case "Very Good"
        StreamLiningCost = 2
    Case "Superior"
        StreamLiningCost = 3
    Case "Excellent"
        StreamLiningCost = 5
    Case "Radical"
        StreamLiningCost = 10
    Case Else
        StreamLiningCost = 1
End Select

' Calculate total cost modifier value for Special Structures
TotalSpecialCost = 1 ' initialize variable
If bResponsive = True Then TotalSpecialCost = ResponsiveCost
If bRobotic = True Then TotalSpecialCost = TotalSpecialCost * RoboticCost
If bBiomechanical = True Then TotalSpecialCost = TotalSpecialCost * BiomechanicalCost
If bLivingMetal = True Then TotalSpecialCost = TotalSpecialCost * LivingMetalCost
    
' Calculate total cost modifier value for Other Modifiers
TotalOtherCost = 1 ' initialize variable
If bSubmersible = True Then TotalOtherCost = 2
If bWingsorRotars = True Then TotalOtherCost = TotalOtherCost * 10
If bLiftingBody = True Then TotalOtherCost = TotalOtherCost * 1.2
If bFlexibody = True Then TotalOtherCost = TotalOtherCost * 2

' Calculate Final Structural Cost
BasicDesignCost = SurfaceArea * StructureCost * StrengthCost * MaterialsCost * StreamLiningCost * TotalSpecialCost * TotalOtherCost
End Function

Function BasicDesignWeight(SurfaceArea As Double, StructureTL As Integer, sStrength As String, sMaterials As String) As Double
'This produces the STRUCTURAL WEIGHT and NOT the BASIC DESIGN WEIGHT
' TODO might rename this function later

Dim StructureWeight As Integer ' Basic Design Weight
Dim StrengthWeight As Single ' Frame Strength Weight Multiplier
Dim MaterialsWeight As Single ' Materials Weight Multiplier
Dim TotalOtherWeight As Single ' Total value of Other Weight Multipliers
Dim objBody As clsBody ' holds the body class of the vehicle
Dim bSubmersible As Boolean
Dim bFlexibody As Boolean

Const SubmersibleCost = 2 ' Submerisble Structure Weight Multiplier
Const FlexibodyCost = 2 ' Flexibody Drivetrain Weight Multiplier

Set objBody = colVehicle.Item("1_")

bSubmersible = objBody.Submersible
bFlexibody = objBody.FlexibodyDrivetrain

Select Case StructureTL
    Case 0, 1
        StructureWeight = 20
    Case 2, 3, 4
        StructureWeight = 18
    Case 5
        StructureWeight = 12
    Case 6
        StructureWeight = 8
    Case 7
        StructureWeight = 6
    Case 8
        StructureWeight = 4
    Case 9
        StructureWeight = 3
    Case 10
        StructureWeight = 2
    Case 11
        StructureWeight = 1.5
    Case Is >= 12
        StructureWeight = 1
End Select

Select Case sStrength
    Case "Super-Light"
        StrengthWeight = 0.1
    Case "Extra-Light"
        StrengthWeight = 0.25
    Case "Light"
        StrengthWeight = 0.5
    Case "Medium"
        StrengthWeight = 1
    Case "Heavy"
        StrengthWeight = 1.5
    Case "Extra-Heavy"
        StrengthWeight = 2
End Select

Select Case sMaterials
    Case "Very Cheap"
        MaterialsWeight = 2
    Case "Cheap"
        MaterialsWeight = 1.5
    Case "Standard"
        MaterialsWeight = 1
    Case "Expensive"
        MaterialsWeight = 0.75
    Case "Very Expensive"
        MaterialsWeight = 0.5
    Case "Advanced"
        MaterialsWeight = 0.375
End Select
   
' Calculate total weight modifier value for Other Modifiers
TotalOtherWeight = 1 ' initialize variable
If bSubmersible = True Then TotalOtherWeight = SubmersibleCost
If bFlexibody = True Then TotalOtherWeight = TotalOtherWeight * FlexibodyCost

' Calculate Final Structural Weight
BasicDesignWeight = SurfaceArea * StructureWeight * StrengthWeight * MaterialsWeight * TotalOtherWeight
End Function


Function CalcHitPoints(SubAssembly As String, FrameStrength As String, Area As Double, ParamArray NumberofWheelsTracksSkids()) As Integer
Dim TempHitPoints As Double

Select Case SubAssembly
Case "clsArm", "clsRotar"
    TempHitPoints = Area * 3
Case "clsBody", "clsSuperStructure", "clsTurret", "clsPopTurret", "clsPod", "clsLeg", "clsWing"
    TempHitPoints = Area * 1.5
Case "clsSkid", "clsTrack"
    TempHitPoints = (Area * 1.5) / (NumberofWheelsTracksSkids(0))
Case "clsGasbag"
    TempHitPoints = Area * 0.01
Case "clsMast", "clsOpenMount"
    TempHitPoints = Area * 2
Case "clsWheel"
    TempHitPoints = (Area * 3) / NumberofWheelsTracksSkids(0)
Case "clsHovercraft"
    TempHitPoints = Area * 1.5
Case "clsHydrofoil"
    TempHitPoints = Area * 1.5
End Select
' /////////////////////////
' I removed the gasbag, open mount and mast sections and placed them directly in the
' classes for those subassemblies.
' /////////////////////////
' Continue to calculate hitpoints for all other types of subassemblies
Select Case FrameStrength
Case "Super-Light"
    TempHitPoints = TempHitPoints / 10
Case "Extra-Light"
    TempHitPoints = TempHitPoints / 4
Case "Light"
    TempHitPoints = TempHitPoints / 2
Case "Heavy"
    TempHitPoints = TempHitPoints * 2
Case "Extra-Heavy"
    TempHitPoints = TempHitPoints * 4
End Select
' produce a rounded (whole) number
CalcHitPoints = CInt(TempHitPoints)
' make sure that we always have at least 1 hit point... can't have 0
If CalcHitPoints < 1 Then
    CalcHitPoints = 1
End If

End Function



'//////////////////////////////////////////////////////////////////////
' Statistics Calculation Functions
'//////////////////////////////////////////////////////////////////////
Sub CalcWeight()
Dim i As Integer ' counter
Dim TempWeight As Double ' holds the weight in progress
Dim TempWeight2 As Double
Dim element As Object ' any object in the collection
' add up the weight of all componenets, structures, features except for Fuel, Ammunition
' provisions, any carried vehicles, and TL5- guns that require carriages.
' the resulting sum is the Empty Weight
TempWeight = 0 ' initialize
TempWeight2 = 0 ' init
For Each element In colVehicle
    ' Skip all the carried vehicles.  This is done by checking for any object that is
    ' attached to a clsVehicleStorage
    If TypeName(element) = "clsPerformance" Then
        GoTo Skip
    ElseIf element.Key = "1_" Then
        GoTo Skip
    ElseIf TypeName(colVehicle.Item(element.Parent)) = "clsVehicleStorage" Then
        TempWeight2 = element.Weight + TempWeight2
        GoTo Skip ' skip to the next component
    End If
    
    ' check for fuel, ammunition, etc.
    Select Case TypeName(element)
        Case "clsFuel", "clsAmmunition", "clsProvisions"
            TempWeight2 = element.Weight + TempWeight2
            GoTo Skip ' skip to the next component
        Case "clsArtilleryandGun"
            If element.Carriage = True Then
                TempWeight2 = element.Weight + TempWeight2
                GoTo Skip ' skip to the next component
            End If
    End Select
    TempWeight = TempWeight + element.Weight
Skip:
Next
With colVehicle.Item("1_")
    .EmptyWeight = TempWeight
    .LoadedWeight = TempWeight + .UsualInternalPayload + TempWeight2
    .HLoadedWeight = TempWeight + .UsualInternalPayload + TempWeight2 + CalcHardPointWeight
    ' note that the loadedmass's are calculated automatically inside the clsBody
End With
End Sub

Function CalcUsualPayload()
' User can decide on normal weight of occupants and cargo for Usual Payload
' We suggest a "generic" usual payload of 200lbs per crew member and passenger
' then add 20lbs per cf. of cargo space to represent a normal load of cargo


End Function

Function CalcHardPointWeight()
' WARNING: Make sure i return at leat a 0 for CalcHardPointWeight if the user
'for example has not installed any hardpoints on the vehicle
Dim TempWeight As Double 'holds weight during calc in progress
Dim i, j As Integer ' counters
Dim element As Object
TempWeight = 0 ' initialize
For Each element In colVehicle
' check for fuel, ammunition, etc.
Select Case TypeName(element)
    Case "clsHardPoint"
        For j = 1 To colVehicle.Count
            If colVehicle(j).Parent = colVehicle.Item(i).Key Then
                TempWeight = TempWeight + colVehicle.Item(j).Weight
            End If
        Next
End Select
Next
CalcHardPointWeight = TempWeight
End Function


Function CalcTotalVolume() As Double
Dim i As Integer ' counter
Dim TempVolume As Double ' holds the volume in progress
Dim element As Object

For Each element In colVehicle
Select Case TypeName(element)
    Case "clsArtilleryandGun", "clsWheel", "clsEngine", "clsBody", "clsSkid", "clsTrack", "clsLeg", "clsWing", "clsRotar", "clsHydrofoil", "clsHovercraft", "clsSuperStructure", "clsOpenMount", "clsMast", "clsPod", "clsTurret", "clsPopTurret", "clsGasbag", "clsArm"
        TempVolume = TempVolume + element.Volume
End Select
Next
CalcTotalVolume = TempVolume
End Function


Function CalcPrice() As Double
' Add up total cost of everything built into the vehicle except for ammunition
' and fuel to find out how much it costs
Dim TempPrice As Double ' holds the volume in progress
Dim i As Integer 'counter
Dim element As Object

For Each element In colVehicle
    If TypeOf element Is clsPerformance Then GoTo Skip
    Select Case TypeName(element)
        Case "clsAmmunition", "clsFuel"
            GoTo Skip ' skip to the next component
    End Select
    TempPrice = TempPrice + element.Cost
Skip:
Next
TempPrice = TempPrice ' + the OptionsCost
CalcPrice = TempPrice
End Function



Function CalcHealth(BodyHitPoints As Long, LoadedWeight As Double) As Integer
Dim TempHealth As Double
' structural HT = (200 * BodyHitPoints / LoadedWeight) + 5
' If the vehicle has Hardpoints, always use the weight WITH hardpoints loaded -
' do NOT use two different values
' Round HT to the nearest whole number.
If LoadedWeight = 0 Then
    TempHealth = 1
Else
    TempHealth = CInt((200 * BodyHitPoints / LoadedWeight) + 5)
End If
' The maximum allowed structural HT is 12 or the vehicle's TL, whichever is greater.
If TempHealth > 12 Then
    If colVehicle.Item("1_").TL > 12 Then
        TempHealth = Val(colVehicle.Item("1_").TL)
    Else
        TempHealth = 12
    End If
End If
CalcHealth = TempHealth
End Function

Sub CalcStats()
' see pages 25 - 26 of Vehicles
Dim EmptyWeight As Double ' holds the vehicles empty weight
Dim OtherWeight As Double ' holds weight of fuel, ammunition, provisions, carried vehicle, and TL 5- gun carriages
Dim UsualPayLoad As Double ' holds the vehicles Usual Internal Payload
Dim LoadedWeight As Double ' holds the vehicles loaded weight W/O hardpoints
Dim HardPointWeight As Double 'holds the weight of just the hardpoints
Dim HLoadedWeight As Double ' holds the vehicles loaded weight W/ Hardpoints loaded
Dim LoadedMass As Double ' holds the loaded mass W/O hardpoints
Dim HLoadedMass As Double 'holds the loaded mass W Hardpoints loaded
Dim TotalVolume As Double ' holds entire vehicles volume
Dim SizeModifier As Integer ' holds the value of the size modifier
Dim SubmergedWeight As Double ' holds the value of the submerged weight
Dim SubmergedMass As Double ' holds the submerged mass in Tons
Dim TotalPrice As Double ' holds the final cost of the vehicle
Dim HitPoints As Long ' holds the Body's hit points
Dim StructuralHealth As Double

CalcWeight
CalcUsualPayload ' this function might be deleted as the UsualPayload will be
'done by the user at a dialog
CalcHardPointWeight

With colVehicle.Item("1_")
    EmptyWeight = .EmptyWeight
    UsualPayLoad = .UsualInternalPayload
    .TotalPrice = CalcPrice ' calc final price of vehicle
    .TotalVolume = CalcTotalVolume
    .SizeModifier = CalcSizeModifier(.TotalVolume)
    ' If vehicle is submersible, calculate the submerged weight and mass
    'If .Submersible = True Then
    '    .SubmergedWeight = CalcSubmergedWeight(.LoadedWeight, .TotalVolume)
    '    .SubmergedMass = .SubmergedWeight / 2000
    'End If
    
    .StructuralHealth = CalcHealth(.HitPoints, .HLoadedWeight)
End With
' Design Check:  Make sure vehicles Max Lift or Floatation rating has not been exceeded
' TODO If it does tell user to convert cargo space to empty space
' to reduce its weight or eliminate some armor

' TODO until i do this, my Performance functions will not work
End Sub


Function CalcSizeModifier(Volume As Double) As Integer
' this looks pretty heinous, but its the only thing i could think of
' will try to optimize later when i finish. Maybe.
Dim TempModifier As Integer

If Volume <= 0.1 Then
    TempModifier = -4
End If

If Volume > 0.1 Then
    If Volume > 0.3 Then
        If Volume > 1 Then
            If Volume > 3 Then
                If Volume > 10 Then
                    If Volume > 30 Then
                        If Volume > 100 Then
                            If Volume > 300 Then
                                If Volume > 1000 Then
                                    If Volume > 3000 Then
                                        If Volume > 10000 Then
                                            If Volume > 30000 Then
                                                If Volume > 100000 Then
                                                    If Volume > 300000 Then
                                                        If Volume > 1000000 Then
                                                            If Volume > 3000000 Then
                                                                If Volume > 10000000 Then
                                                                    If Volume > 30000000 Then
                                                                        If Volume > 100000000 Then
                                                                            If Volume > 300000000 Then
                                                                                If Volume > 100000000 Then
                                                                                   If Volume > 300000000 Then
                                                                                       If Volume > 1000000000 Then
                                                                                            If Volume > 3000000000# Then
                                                                                                If Volume > 10000000000# Then
                                                                                                   If Volume > 30000000000# Then
                                                                                                   MsgBox "Maximum Vehicle volume of 10 billion cubic feet has been exceeded.  Consult your GURPS Vehicles book if you wish to create a vehicle larger than this limit."
                                                                                                   TempModifier = 22
                                                                                                   Else
                                                                                                   TempModifier = 21
                                                                                                   End If
                                                                                                Else
                                                                                                TempModifier = 20
                                                                                                End If
                                                                                            Else
                                                                                            TempModifier = 19
                                                                                            End If
                                                                                       Else
                                                                                       TempModifier = 18
                                                                                       End If
                                                                                   Else
                                                                                   TempModifier = 17
                                                                                   End If
                                                                                Else
                                                                                TempModifier = 16
                                                                                End If
                                                                            Else
                                                                            TempModifier = 15
                                                                            End If
                                                                        Else
                                                                        TempModifier = 14
                                                                        End If
                                                                    Else
                                                                    TempModifier = 13
                                                                    End If
                                                                Else
                                                                TempModifier = 12
                                                                End If
                                                            Else
                                                            TempModifier = 11
                                                            End If
                                                        Else
                                                        TempModifier = 10
                                                        End If
                                                    Else
                                                    TempModifier = 9
                                                    End If
                                                Else
                                                TempModifier = 8
                                                End If
                                            Else
                                            TempModifier = 7
                                            End If
                                         Else
                                         TempModifier = 6
                                         End If
                                    Else
                                    TempModifier = 5
                                    End If
                                Else
                                TempModifier = 4
                                End If
                            Else
                            TempModifier = 3
                            End If
                        Else
                        TempModifier = 2
                        End If
                    Else
                    TempModifier = 1
                    End If
                Else
                TempModifier = 0
                End If
            Else
            TempModifier = -1
            End If
        Else
        TempModifier = -2
        End If
    Else
    TempModifier = -3
    End If
Else
TempModifier = -4
End If

CalcSizeModifier = TempModifier
End Function

