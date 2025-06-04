Attribute VB_Name = "modSurfaceFeatures"
Option Explicit

Sub SaveOptions()
With colVehicle.Item("1_")
    If frmOptions.chkSubmersible = 1 Then
        .Submersible = True
        .Sealed = False ' submerisbles are sealed for free
        .WaterProof = False ' submersibles are waterproof for free
    Else
        .Submersible = False
        .Sealed = frmOptions.chksealed.Value
        .WaterProof = frmOptions.chkwaterproof.Value
    End If
    If frmOptions.chksealed.Enabled = True Then
        If frmOptions.chksealed = 1 Then
            .Sealed = True
            .WaterProof = False ' sealed vehicles are waterproofed for free
        End If
    End If
    If frmOptions.chkForce = False Then
        .ForceScreen = "None"
    ElseIf frmOptions.optVariableScreen = True Then
        .ForceScreen = "Variable"
        .ForceDR = frmOptions.txtForceDR
    Else
        .ForceScreen = "Standard"
        .ForceDR = frmOptions.txtForceDR
    End If
    If frmOptions.chkPin = False Then
        .Pin = "None"
    ElseIf frmOptions.optExplosivePin = True Then
        .Pin = "Explosive"
    Else
        .Pin = "Standard"
    End If
      .SubmarineLines = frmOptions.chkSubmarine
      .FloatationHull = frmOptions.chkFloatation.Value
      .HydrodynamicLines = frmOptions.cboLines
      .Catamaran = frmOptions.optCatamaran.Value
      .Trimaran = frmOptions.optTrimaran.Value
      .StreamLining = frmOptions.cboStreamlining
      .LiftingBody = frmOptions.chkLiftingbody.Value
      .RollStabilizers = frmOptions.chkRollStabilizers.Value
      .Camouflage = frmOptions.chkCamouflage.Value
      .InfraredCloaking = frmOptions.cboInfraredCloaking.Text
      .EmissionCloaking = frmOptions.cboEmissionCloaking.Text
      .SoundBaffling = frmOptions.cboSoundBaffling.Text
      .Stealth = frmOptions.cboStealth.Text
      .LiquidCrystal = frmOptions.chkLiquidCrystal.Value
      .Chameleon = frmOptions.cboChameleon.Text
      .PsiShielding = frmOptions.chkPsiShielding.Value
      .Deflector = frmOptions.chkDeflector
      .Ram = frmOptions.chkRam
      .Bulldozer = frmOptions.chkBulldozer
      .Plow = frmOptions.chkPlow
      .Hitch = frmOptions.chkHitch
      .Convertible = frmOptions.cboConvertible.Text
    End With
End Sub


Sub CalcOptionsWeightandCost()
Dim element As Object
Dim OptionsWeight As Double
Dim OptionsCost As Double
Dim TotalArea As Double
Dim BodyArea As Double
Dim BodyHits As Long
Dim TrimmedArea As Double
Dim IgnoredArea As Double
Dim CamouflageCost As Double
Dim InfraredCost As Double
Dim InfraredWeight As Double
Dim EmissionCost As Double
Dim EmissionWeight As Double
Dim SoundCost As Double
Dim SoundWeight As Double
Dim StealthCost As Double
Dim StealthWeight As Double
Dim ChameleonCost As Double
Dim ChameleonWeight As Double
Dim LiquidCrystalCost As Double
Dim LiquidCrystalWeight As Double
Dim PsiShieldingCost As Double
Dim PsiShieldingWeight As Double
Dim RamWeight As Double
Dim RamCost As Double
Dim BullDozerWeight As Double
Dim BullDozerCost As Double
Dim PlowWeight As Double
Dim PlowCost As Double
Dim HitchCost As Double
Dim HitchWeight As Double
Dim PinCost As Double
Dim PinWeight As Double
Dim ForceCost As Double
Dim ForceWeight As Double
Dim DeflectorCost As Double
Dim DeflectorWeight As Double
Dim ConvertibleCost As Double
Dim ConvertibleWeight As Double
Dim RollStabilizerWeight As Double
Dim RollStabilizerCost As Double
Dim WaterProofCost As Double
Dim SealedCost As Double
Const BasicInfrared = 2
Const RadicalInfrared = 3
Const BasicEmission = 4
Const RadicalEmission = 5
Const BasicSound = 6
Const RadicalSound = 7
Const BasicStealth = 8
Const RadicalStealth = 9
Const BasicChameleon = 10
Const InstantChameleon = 11
Const IntruderChameleon = 12
Const LiquidCrystal = 13
Const PsiShielding = 14
Const Deflector = 18
Const ForceScreen = 19

For Each element In colVehicle
    If TypeOf element Is clsMast Then
        IgnoredArea = IgnoredArea + element.Area
    ElseIf TypeOf element Is clsSkid Then
        IgnoredArea = IgnoredArea + element.Area
    ElseIf TypeOf element Is clsGasbag Then
        IgnoredArea = IgnoredArea + element.Area
    End If
Next

With colVehicle("1_")

' Get the surface area of the body
BodyArea = .SurfaceArea
' Get the surface are of the entire vehicle
TotalArea = .SurfaceArea ' TODO this should reference .TotalSurfaceArea but i havent calced that yet
BodyHits = .HitPoints
' Get the area minus skids, masts and gas bags (for sound baffling)
TrimmedArea = TotalArea - IgnoredArea
' Get Cost for camouflage (note there is no weight)
    If .Camouflage = False Then
        CamouflageCost = 0
    Else
        CamouflageCost = 0.1 * TotalArea
    End If
' Get Cost and Weight for Infrared Cloaking
    If .InfraredCloaking = "None" Then
        InfraredCost = 0
        InfraredWeight = 0
    ElseIf .InfraredCloaking = "Basic" Then
        InfraredCost = GetSurfaceCost(BasicInfrared) * TotalArea
        InfraredWeight = GetSurfaceWeight(BasicInfrared) * TotalArea
    Else
        InfraredCost = GetSurfaceCost(RadicalInfrared) * TotalArea
        InfraredWeight = GetSurfaceWeight(RadicalInfrared) * TotalArea
    End If
' Get Cost and Weight for Emission Cloaking
    If .EmissionCloaking = "None" Then
        EmissionCost = 0
        EmissionWeight = 0
    ElseIf .EmissionCloaking = "Basic" Then
        EmissionCost = GetSurfaceCost(BasicEmission) * TotalArea
        EmissionWeight = GetSurfaceWeight(BasicEmission) * TotalArea
    Else
        EmissionCost = GetSurfaceCost(RadicalEmission) * TotalArea
        EmissionWeight = GetSurfaceWeight(RadicalEmission) * TotalArea
    End If
' Get cost and weight for Sound Baffling
    If .SoundBaffling = "None" Then
        SoundCost = 0
        SoundWeight = 0
    ElseIf .SoundBaffling = "Basic" Then
        SoundCost = GetSurfaceCost(BasicSound) * TrimmedArea
        SoundWeight = GetSurfaceWeight(BasicSound) * TrimmedArea
    Else
        SoundCost = GetSurfaceCost(RadicalSound) * TrimmedArea
        SoundWeight = GetSurfaceWeight(RadicalSound) * TrimmedArea
    End If
' Get cost and weight for Stealth
    If .Stealth = "None" Then
        StealthCost = 0
        StealthWeight = 0
    ElseIf .Stealth = "Basic" Then
        StealthCost = GetSurfaceCost(BasicStealth) * TotalArea
        StealthWeight = GetSurfaceWeight(BasicStealth) * TotalArea
    Else
        StealthCost = GetSurfaceCost(RadicalStealth) * TotalArea
        StealthWeight = GetSurfaceWeight(RadicalStealth) * TotalArea
    End If
'Get cost and weight for Chameleon system
    If .Chameleon = "None" Then
        ChameleonCost = 0
        ChameleonWeight = 0
    ElseIf .Chameleon = "Basic" Then
        ChameleonCost = GetSurfaceCost(BasicChameleon) * TotalArea
        ChameleonWeight = GetSurfaceWeight(BasicChameleon) * TotalArea
    ElseIf .Chameleon = "Instant" Then
        ChameleonCost = GetSurfaceCost(InstantChameleon) * TotalArea
        ChameleonWeight = GetSurfaceWeight(InstantChameleon) * TotalArea
    Else
        ChameleonCost = GetSurfaceCost(IntruderChameleon) * TotalArea
        ChameleonWeight = GetSurfaceWeight(IntruderChameleon) * TotalArea
    End If
'Get cost and weight for LiquidCrystal skin
    If .LiquidCrystal = False Then
        LiquidCrystalCost = 0
        LiquidCrystalWeight = 0
    Else
        LiquidCrystalCost = GetSurfaceCost(LiquidCrystal) * TotalArea
        LiquidCrystalWeight = GetSurfaceWeight(LiquidCrystal) * TotalArea
    End If
' Get cost and weight for PsiShielding
    If .PsiShielding = False Then
        PsiShieldingCost = 0
        PsiShieldingWeight = 0
    Else
        PsiShieldingCost = GetSurfaceCost(PsiShielding) * TotalArea
        PsiShieldingWeight = GetSurfaceWeight(PsiShielding) * TotalArea
    End If
' Get wieght and cost for Deflector Fields
If .Deflector = False Then
    DeflectorCost = 0
    DeflectorWeight = 0
Else
    DeflectorCost = GetSurfaceCost(Deflector) * TotalArea
    DeflectorWeight = GetSurfaceWeight(Deflector) * TotalArea
End If
' Get weight and cost for ForceScreens
If .ForceScreen = "None" Then
    ForceCost = 0
    ForceWeight = 0
Else
    ForceCost = GetSurfaceCost(ForceScreen) * TotalArea * .ForceDR
    ForceWeight = GetSurfaceWeight(ForceScreen) * TotalArea * .ForceDR
End If
' Get weight and cost for Rams
If .Ram = True Then
    RamWeight = 1 * BodyArea
    RamCost = 2 * BodyArea
Else
    RamWeight = 0
    RamCost = 0
End If
' Get weight and cost for Bulldozers
If .Bulldozer = True Then
    BullDozerWeight = 2 * BodyArea
    BullDozerCost = 4 * BodyArea
Else
    BullDozerWeight = 0
    BullDozerCost = 0
End If
' Get wieght and cost for Plows
If .Plow = True Then
    PlowWeight = 2 * BodyArea
    PlowCost = 4 * BodyArea
Else
    PlowWeight = 0
    PlowCost = 0
End If
' Get weight and cost for Convertible
If .Convertible = "None" Then
ConvertibleCost = 0
ConvertibleWeight = 0
Else
End If
' Get weight and cost for Hitch
If .Hitch = True Then
    HitchCost = 0.1 * BodyHits
    HitchWeight = HitchCost
Else
    HitchCost = 0
    HitchWeight = 0
End If
' Get weight and cost for Pin
If .Pin <> "None" Then
    PinCost = 0.05 * BodyHits
    PinWeight = 0.1 * BodyHits
    If .Pin = "Explosive" Then
        PinCost = PinCost * 5
    End If
Else
    PinCost = 0
    PinWeight = 0
End If
'Get weight and cost for Roll Stabilizers
If .RollStabilizers = False Then
    RollStabilizerCost = 0
    RollStabilizerWeight = 0
Else
    RollStabilizerCost = 0.1 * (BodyArea / .StructuralSurfaceArea) * .StructuralCost
    RollStabilizerWeight = 0.05 * (BodyArea / .StructuralSurfaceArea) * .StructuralWeight
End If
'Get Cost for waterproofing
If .WaterProof = True Then
    WaterProofCost = 2 * .StructuralSurfaceArea
Else
    WaterProofCost = 0
End If
'Get Cost for Sealed vehicle
If .Sealed = True Then
    If gVehicleTL <= 7 Then
        SealedCost = 40 * .StructuralSurfaceArea
    ElseIf gVehicleTL = 8 Then
        SealedCost = 20 * .StructuralSurfaceArea
    Else
        SealedCost = 10 * .StructuralSurfaceArea
    End If
Else
    SealedCost = 0
End If
End With
OptionsWeight = RollStabilizerWeight + ForceWeight + DeflectorWeight + ConvertibleWeight + PinWeight + HitchWeight + PlowWeight + BullDozerWeight + RamWeight + PsiShieldingWeight + LiquidCrystalWeight + ChameleonWeight + StealthWeight + SoundWeight + EmissionWeight + InfraredWeight
OptionsCost = SealedCost + WaterProofCost + RollStabilizerCost + ForceCost + DeflectorCost + ConvertibleCost + PinCost + HitchCost + PlowCost + BullDozerCost + RamCost + PsiShieldingCost + LiquidCrystalCost + ChameleonCost + StealthCost + SoundCost + InfraredCost + EmissionCost + CamouflageCost
frmOptions.StatusBar1.Panels(2).Text = "Added Weight: " & OptionsWeight & "lbs"
frmOptions.StatusBar1.Panels(1).Text = "Added Cost: $ " & OptionsCost

End Sub

Function GetSurfaceCost(FeatureID As Integer) As Single

' This routine calculates the Cost of a surface Feature
' IMPORTANT: This routine is optimized to only check valid techlevels!!
' If the user has somehow enabled a Surface Feature that is not allowed
' at the vehicles tech level, this routine will return 0!!!!  I must make
' sure to gray out features that cant be selected in the dialog

Dim CostModifier As Single
Dim i As Integer ' counter
Dim TempModifier As Single
Dim TempTech As Integer

'On Error GoTo TechLevelError
    ' init the two temporary variables
    TempModifier = 0
    TempTech = 0
   ' Get the Cost and Weight Modifiers
       For i = 1 To UBound(SurfaceMatrix)
           If SurfaceMatrix(i).FeatureType = FeatureID Then
               If SurfaceMatrix(i).TL = gVehicleTL Then
                    CostModifier = SurfaceMatrix(i).CostMod
                    GetSurfaceCost = CostModifier
                    Exit Function
               ElseIf SurfaceMatrix(i).TL < gVehicleTL Then
                    If SurfaceMatrix(i).TL > TempTech Then
                        CostModifier = SurfaceMatrix(i).CostMod
                        TempTech = SurfaceMatrix(i).TL
                    End If
               End If
           End If
       Next
GetSurfaceCost = CostModifier
'TechLevelError:
'MsgBox "Error In Function GetSurfaceCost:Unsupported TechLevel with feature ID # " & FeatureID
End Function

Function GetSurfaceWeight(FeatureID As Integer) As Single
' This routine calculates the Weight of a surface Feature
' IMPORTANT: This routine is optimized to only check valid techlevels!!
' If the user has somehow enabled a Surface Feature that is not allowed
' at the vehicles tech level, this routine will return 0!!!!

Dim WeightModifier As Single
Dim i As Integer ' counter
Dim TempTech As Integer
Dim TempModifier As Single

'On Error GoTo TechLevelError
    ' init the two temporary variables
    TempModifier = 0
    TempTech = 0
   ' Get the Weight Modifiers
       For i = 1 To UBound(SurfaceMatrix)
           If SurfaceMatrix(i).FeatureType = FeatureID Then
               If SurfaceMatrix(i).TL = gVehicleTL Then
                    WeightModifier = SurfaceMatrix(i).WeightMod
                    GetSurfaceWeight = WeightModifier
                    Exit Function
               ElseIf SurfaceMatrix(i).TL < gVehicleTL Then
                    If SurfaceMatrix(i).TL > TempTech Then
                        WeightModifier = SurfaceMatrix(i).WeightMod
                        TempTech = SurfaceMatrix(i).TL
                    End If
               End If
           End If
       Next
GetSurfaceWeight = WeightModifier
'TechLevelError:
'MsgBox "Error In Function GetSurfaceWeight:Unsupported TechLevel with feature ID # " & FeatureID
End Function
