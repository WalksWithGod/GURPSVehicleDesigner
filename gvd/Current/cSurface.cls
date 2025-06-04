VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements cIPersist
Implements cINode
Implements cIDisplay

' -------cINode interface variables
Private m_lngMaxChildren As Long
Private m_lngChildCount As Long
Private m_oChildren() As cINode
Private m_lngAttributes As Long
Private m_hParent As Long
Private m_hMe As Long
Private m_sName As String
Private m_sDescription As String
Private m_sImage As String

' -------cIDisplay interface variables
Private m_lngPropCount As Long
Private m_lngCurrentPropItem As Long
Private m_oProperties() As cPropertyItem



Private m_bLevitationCoating As Boolean
Private m_bMagicLevitation As Boolean
Private m_bAntigravityCoating As Boolean
Private m_bSuperScienceCoating As Boolean

Private m_sngMagicLevitationEnergyCost As Single
Private m_sngAntigravityCoatingCost As Single
Private m_sngSuperScienceCoatingCost As Single
Private m_sngMagicLevitationEnergyCostPerPound As Single
Private m_sngAntigravityCoatingCostPerSquareFoot As Single
Private m_sngSuperScienceCoatingCostPerSquareFoot As Single
Private m_byteAntigravityCoatingSurfaceAreaUseage As Byte
Private m_byteSuperScienceCoatingSurfaceAreaUseage As Byte



' totals
Private m_dblWeight As Double
Private m_dblCost As Double

Public Property Let Cost(ByVal dbl As Double)
    m_dblCost = dbl
End Property
Public Property Get Cost() As Double
    Cost = m_dblCost
End Property
Public Property Let Weight(ByVal dbl As Double)
    m_dblWeight = dbl
End Property
Public Property Get Weight() As Double
    Weight = m_dblWeight
End Property



Public Property Let MagicLevitationEnergyCostPerPound(ByVal sng As Single)
    mvarMagicLevitationEnergyCostPerPound = sng
End Property
Public Property Get MagicLevitationEnergyCostPerPound() As Single
    MagicLevitationEnergyCostPerPound = mvarMagicLevitationEnergyCostPerPound
End Property
Public Property Let MagicLevitationEnergyCost(ByVal sng As Single)
    mvarMagicLevitationEnergyCost = sng
End Property
Public Property Get MagicLevitationEnergyCost() As Single
    MagicLevitationEnergyCost = mvarMagicLevitationEnergyCost
End Property
Public Property Let AntigravityCoatingCostPerSquareFoot(ByVal sng As Single)
    mvarAntigravityCoatingCostPerSquareFoot = sng
End Property
Public Property Get AntigravityCoatingCostPerSquareFoot() As Single
    AntigravityCoatingCostPerSquareFoot = mvarAntigravityCoatingCostPerSquareFoot
End Property
Public Property Let AntigravityCoatingCost(ByVal sng As Single)
    mvarAntigravityCoatingCost = sng
End Property
Public Property Get AntigravityCoatingCost() As Single
    AntigravityCoatingCost = mvarAntigravityCoatingCost
End Property
Public Property Let SuperScienceCoatingCostPerSquareFoot(ByVal sng As Single)
    mvarSuperScienceCoatingCostPerSquareFoot = sng
End Property
Public Property Get SuperScienceCoatingCostPerSquareFoot() As Single
    SuperScienceCoatingCostPerSquareFoot = mvarSuperScienceCoatingCostPerSquareFoot
End Property
Public Property Let SuperScienceCoatingCost(ByVal sng As Single)
    mvarSuperScienceCoatingCost = sng
End Property
Public Property Get SuperScienceCoatingCost() As Single
    SuperScienceCoatingCost = mvarSuperScienceCoatingCost
End Property


Public Property Let bMagicLevitation(ByVal bln As Boolean)
    mvarbMagicLevitation = bln
End Property
Public Property Get bMagicLevitation() As Boolean
    bMagicLevitation = mvarbMagicLevitation
End Property
Public Property Let bAntigravityCoating(ByVal bln As Boolean)
    mvarbAntigravityCoating = bln
End Property
Public Property Get bAntigravityCoating() As Boolean
    bAntigravityCoating = mvarbAntigravityCoating
End Property
Public Property Let bSuperScienceCoating(ByVal bln As Boolean)
    mvarbSuperScienceCoating = bln
End Property
Public Property Get bSuperScienceCoating() As Boolean
    bSuperScienceCoating = mvarbSuperScienceCoating
End Property
Public Property Let AntigravityCoatingSurfaceAreaUseage(ByVal s As String)
    mvarAntigravityCoatingSurfaceAreaUseage = s
End Property
Public Property Get AntigravityCoatingSurfaceAreaUseage() As String
    AntigravityCoatingSurfaceAreaUseage = mvarAntigravityCoatingSurfaceAreaUseage
End Property
Public Property Let SuperScienceCoatingSurfaceAreaUseage(ByVal s As String)
    mvarSuperScienceCoatingSurfaceAreaUseage = s
End Property
Public Property Get SuperScienceCoatingSurfaceAreaUseage() As String
    SuperScienceCoatingSurfaceAreaUseage = mvarSuperScienceCoatingSurfaceAreaUseage
End Property



Private Sub Class_Initialize()
    
    m_bLevitationCoating = False
    m_bMagicLevitation = False
    m_bAntigravityCoating = False
    m_bSuperScienceCoating = False
    m_sngMagicLevitationEnergyCostPerPound = 700
    m_sngAntigravityCoatingCostPerSquareFoot = 10
    m_sngSuperScienceCoatingCostPerSquareFoot = 100
    m_byteAntigravityCoatingSurfaceAreaUseage = 1 '"Body"
    m_byteSuperScienceCoatingSurfaceAreaUseage = 1 '"Body"
End Sub


Sub Update()

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

    '----------
    Dim dblBodyArea As Double
    Dim dblTotalArea As Double
    Dim BodyHits As Single
    Dim dblTrimmedArea As Double
    Dim dblIgnoredArea As Double
    
    With Veh.Components(BODY_KEY)
        ' Get the surface area of the body
        dblBodyArea = .SurfaceArea
        ' Get the surface are of the entire vehicle
        dblTotalArea = Veh.Stats.totalSurfaceArea
        BodyHits = .HitPoints
    End With
    ' Get the area minus skids, masts and gas bags (for sound baffling rules on page 92)
    ' TODO: dblIgnoredArea is 0? wtf i think i forgot to get the skis,mast and bags area.
    ' And isnt that the same as structural surface area anyway?  I forget, look it up.
    dblTrimmedArea = dblTotalArea - dblIgnoredArea
    ' Get Cost for camouflage (note there is no weight)
    If m_bCamouflage = False Then
        m_sngCamouflageCost = 0
    Else
        m_sngCamouflageCost = 0.1 * dblTotalArea
    End If
    ' Get Cost and Weight for Infrared Cloaking
    If m_byteInfraredCloaking = None Then
        m_sngInfraredCost = 0
        m_sngInfraredWeight = 0
    ElseIf m_byteInfraredCloaking = BASIC Then
        m_sngInfraredCost = GetSurfaceCost(BasicInfrared) * dblTotalArea
        m_sngInfraredWeight = GetSurfaceWeight(BasicInfrared) * dblTotalArea
    Else
        Debug.Assert m_byteInfraredCloaking = RADICAL
        m_sngInfraredCost = GetSurfaceCost(RadicalInfrared) * dblTotalArea
        m_sngInfraredWeight = GetSurfaceWeight(RadicalInfrared) * dblTotalArea
    End If
    ' Get Cost and Weight for Emission Cloaking
    If m_byteEmissionCloaking = None Then
        mvarEmissionCost = 0
        mvarEmissionWeight = 0
    ElseIf m_byteEmissionCloaking = BASIC Then
        mvarEmissionCost = GetSurfaceCost(BasicEmission) * dblTotalArea
        mvarEmissionWeight = GetSurfaceWeight(BasicEmission) * dblTotalArea
    Else
        Debug.Assert m_byteEmissionCloaking = RADICAL
        mvarEmissionCost = GetSurfaceCost(RadicalEmission) * dblTotalArea
        mvarEmissionWeight = GetSurfaceWeight(RadicalEmission) * dblTotalArea
    End If
    ' Get cost and weight for Sound Baffling
    If m_byteSoundBaffling = None Then
        mvarSoundCost = 0
        mvarSoundWeight = 0
    ElseIf m_byteSoundBaffling = BASIC Then
        mvarSoundCost = GetSurfaceCost(BasicSound) * dblTrimmedArea
        mvarSoundWeight = GetSurfaceWeight(BasicSound) * dblTrimmedArea
    Else
        Debug.Assert m_byteSoundBaffling = RADICAL
        mvarSoundCost = GetSurfaceCost(RadicalSound) * dblTrimmedArea
        mvarSoundWeight = GetSurfaceWeight(RadicalSound) * dblTrimmedArea
    End If
    ' Get cost and weight for Stealth
    If mvarStealth = "none" Then
        mvarStealthCost = 0
        mvarStealthWeight = 0
    ElseIf mvarStealth = "basic" Then
        mvarStealthCost = GetSurfaceCost(BasicStealth) * dblTotalArea
        mvarStealthWeight = GetSurfaceWeight(BasicStealth) * dblTotalArea
    Else
        mvarStealthCost = GetSurfaceCost(RadicalStealth) * dblTotalArea
        mvarStealthWeight = GetSurfaceWeight(RadicalStealth) * dblTotalArea
    End If
    'Get cost and weight for Chameleon system
    If mvarChameleon = "none" Then
        mvarChameleonCost = 0
        mvarChameleonWeight = 0
    ElseIf mvarChameleon = "basic" Then
        mvarChameleonCost = GetSurfaceCost(BasicChameleon) * dblTotalArea
        mvarChameleonWeight = GetSurfaceWeight(BasicChameleon) * dblTotalArea
    ElseIf mvarChameleon = "instant" Then
        mvarChameleonCost = GetSurfaceCost(InstantChameleon) * dblTotalArea
        mvarChameleonWeight = GetSurfaceWeight(InstantChameleon) * dblTotalArea
    Else
        mvarChameleonCost = GetSurfaceCost(IntruderChameleon) * dblTotalArea
        mvarChameleonWeight = GetSurfaceWeight(IntruderChameleon) * dblTotalArea
    End If
    'Get cost and weight for LiquidCrystal skin
    If mvarLiquidCrystal = False Then
        mvarLiquidCrystalCost = 0
        mvarLiquidCrystalWeight = 0
    Else
        mvarLiquidCrystalCost = GetSurfaceCost(LiquidCrystal) * dblTotalArea
        mvarLiquidCrystalWeight = GetSurfaceWeight(LiquidCrystal) * dblTotalArea
    End If
    ' Get cost and weight for PsiShielding
    If mvarPsiShielding = False Then
        mvarPsiShieldingCost = 0
        mvarPsiShieldingWeight = 0
    Else
        mvarPsiShieldingCost = GetSurfaceCost(PsiShielding) * dblTotalArea
        mvarPsiShieldingWeight = GetSurfaceWeight(PsiShielding) * dblTotalArea
    End If
    ' Get cost for Levitation
    If mvarbMagicLevitation = True Then
        mvarMagicLevitationEnergyCost = mvarMagicLevitationEnergyCostPerPound * (dblTotalArea / 250)
    Else
        mvarMagicLevitationEnergyCost = 0
    End If
    
    If mvarbAntigravityCoating = True Then
        If mvarAntigravityCoatingSurfaceAreaUseage = "Vehicle" Then
            mvarAntigravityCoatingCost = mvarAntigravityCoatingCostPerSquareFoot * dblTotalArea
        Else
            mvarAntigravityCoatingCost = mvarAntigravityCoatingCostPerSquareFoot * dblBodyArea
        End If
    Else
        mvarAntigravityCoatingCost = 0
    End If
    
    If mvarbSuperScienceCoating = True Then
        If mvarSuperScienceCoatingSurfaceAreaUseage = "Vehicle" Then
            mvarSuperScienceCoatingCost = mvarSuperScienceCoatingCostPerSquareFoot * dblTotalArea
        Else
            mvarSuperScienceCoatingCost = mvarSuperScienceCoatingCostPerSquareFoot * dblBodyArea
        End If
    Else
        mvarSuperScienceCoatingCost = 0
    End If
    
    
    'Get Cost for waterproofing
    If mvarWaterProof Then
        mvarWaterProofCost = 2 * Veh.Stats.StructuralSurfaceArea
    Else
        mvarWaterProofCost = 0
    End If
    'Get Cost for Sealed vehicle
    If mvarSealed Then
        
        If gVehicleTL <= 7 Then
            mvarSealedCost = 40 * Veh.Stats.StructuralSurfaceArea
        ElseIf gVehicleTL = 8 Then
            mvarSealedCost = 20 * Veh.Stats.StructuralSurfaceArea
        Else
            mvarSealedCost = 10 * Veh.Stats.StructuralSurfaceArea
        End If
    Else
        mvarSealedCost = 0
    End If
    
    ' total surface feature weights
    mvarWeight = mvarPsiShieldingWeight + mvarLiquidCrystalWeight + _
        mvarChameleonWeight + mvarStealthWeight + mvarSoundWeight + mvarEmissionWeight + _
        mvarInfraredWeight
    
    ' total surface feature costs
    mvarCost = mvarAntigravityCoatingCost + mvarSuperScienceCoatingCost + _
        mvarSealedCost + mvarWaterProofCost + _
        mvarPsiShieldingCost + mvarLiquidCrystalCost + mvarChameleonCost + _
        mvarStealthCost + mvarSoundCost + mvarInfraredCost + mvarEmissionCost + _
        mvarCamouflageCost


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

'todo: note the inconsistancies here with these functions returning singles into doubles and what not.
' I need to make sure they are the same.
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



'///////////////////////////////////////////////////
'//cIDisplay Implemented Properties and Functions
Private Function cIDisplay_getFirstPropertyItem() As cPropertyItem
    On Error GoTo err:
    If Not m_oProperties(0) Is Nothing Then
        Set cIDisplay_getFirstPropertyItem = m_oProperties(0)
        m_lngCurrentPropItem = 0
    End If
    Exit Function
err:
    Debug.Print TypeName(Me) & ":cIDisplay_getFirstPropertyItem() -- no properties in m_oProperties() array for " & TypeName(Me)
End Function

Private Function cIDisplay_getNextPropertyItem() As cPropertyItem
    m_lngCurrentPropItem = m_lngCurrentPropItem + 1
    If m_lngCurrentPropItem <= m_lngPropCount - 1 Then
        If Not m_oProperties(m_lngCurrentPropItem) Is Nothing Then
            Set cIDisplay_getNextPropertyItem = m_oProperties(m_lngCurrentPropItem)
        End If
    Else
        m_lngCurrentPropItem = m_lngCurrentPropItem - 1
        Debug.Print TypeName(Me) & ":cIDisplay:getNextPropertyItem -- nextItem exceeds Property Count."
    End If
End Function
Private Function cIDisplay_getPropertyItemByIndex(ByVal iIndex As Long) As cPropertyItem
    On Error Resume Next
    Set cIDisplay_getPropertyItemByIndex = m_oProperties(iIndex)
End Function
'/////////////////////////////////////////////
'//Implemented cINode Properties and Functions
Private Function cINode_AddChild(oBase As cINode) As Boolean
    If m_lngMaxChildren = m_lngChildCount Then
        cINode_AddChild = False
    Else
        m_lngChildCount = m_lngChildCount + 1
        ReDim Preserve m_oChildren(m_lngChildCount - 1)
        Set m_oChildren(m_lngChildCount - 1) = oBase
        cINode_AddChild = True
    End If
End Function
Private Function cINode_getChildrenByClassName(ByRef Classname As String, hChilds() As Long) As Boolean
End Function
Private Function cINode_getChildIndexByHandle(ByVal h As Long) As Long
    Dim i As Long
    Dim lRet As Long
    lRet = -1
    For i = 0 To m_lngChildCount - 1
        If m_oChildren(i).Handle = h Then lRet = i: Exit For
    Next
    cINode_getChildIndexByHandle = lRet
End Function
Private Function cINode_getChild(ByVal lngIndex As Long) As cINode
    If (lngIndex >= 0) And (m_lngChildCount > 0) And (lngIndex <= m_lngChildCount - 1) Then
        Set cINode_getChild = m_oChildren(lngIndex)
    End If
End Function
Private Function cINode_removeChild(ByVal lngIndex As Long) As Boolean
    Dim i As Long
    If (lngIndex <= m_lngChildCount - 1) And (lngIndex >= 0) Then
        Set m_oChildren(lngIndex) = Nothing
        For i = lngIndex + 1 To m_lngChildCount - 1
            Set m_oChildren(i - 1) = m_oChildren(i)
        Next
    End If
    m_lngChildCount = m_lngChildCount - 1
    If m_lngChildCount > 0 Then
        ReDim Preserve m_oChildren(0 To m_lngChildCount - 1)
    Else
        Erase m_oChildren
    End If
End Function
Private Property Get cINode_childCount() As Long
    cINode_childCount = m_lngChildCount
End Property
Private Property Get cINode_ClassName() As String
    cINode_ClassName = TypeName(Me)
End Property
Private Property Get cINode_Attributes() As Long
    cINode_Attributes = m_lngAttributes
End Property
Private Property Let cINode_Handle(ByVal l As Long)
    m_hMe = l
End Property
Private Property Get cINode_Handle() As Long
    cINode_Handle = m_hMe
End Property
Private Property Let cINode_Parent(ByVal l As Long)
    m_hParent = l
End Property
Private Property Get cINode_Parent() As Long
    cINode_Parent = m_hParent
End Property
Private Property Let cINode_Name(ByVal s As String)
    m_sName = s
End Property
Private Property Get cINode_Name() As String
    cINode_Name = m_sName
End Property
Private Property Let cINode_Description(ByVal s As String)
    m_sDescription = s
End Property
Private Property Get cINode_Description() As String
    cINode_Description = m_sDescription
End Property
Private Property Let cINode_Image(ByVal s As String)
    m_sImage = s
End Property
Private Property Get cINode_Image() As String
    cINode_Image = m_sImage
End Property

'//cIPersist Interface
Private Property Get cIPersist_Classname() As String
    cIPersist_Classname = TypeName(Me) 'todo: did i ever test that typename me returns the base class and not the class of an interface in use?
End Property
Private Property Get cIPersist_GUID() As String
End Property
Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Long)
    Dim i As Long
    
    If iMode = cmp Then
        m_sName = op.Load(XML_NODE_NAME)
        m_sDescription = op.Load(XML_NODE_DESCRIPTION)
        m_lngChildCount = op.Load(XML_NODE_CHILDCOUNT)
        If m_lngChildCount > 0 Then
            ReDim m_oChildren(m_lngChildCount - 1)
            For i = 0 To m_lngChildCount - 1
                Set m_oChildren(i) = op.Load(XML_NODE_CHILD & i)
                m_oChildren(i).Parent = m_hMe
                Debug.Assert m_oChildren(i).Parent <> 0
            Next
        End If
    Else 'DEF
        m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
        m_sImage = op.Load(XML_NODE_IMAGE)
        
        ' load properties last, these will reference variables initialized above
        m_lngPropCount = op.Load(XML_NODE_PROPERTYCOUNT)
        If m_lngPropCount > 0 Then
            ReDim m_oProperties(m_lngPropCount - 1)
            For i = 0 To m_lngPropCount - 1
                Set m_oProperties(i) = op.Load(XML_NODE_PROPERTY & i)
            Next
        End If
    End If
End Sub
Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
End Sub



