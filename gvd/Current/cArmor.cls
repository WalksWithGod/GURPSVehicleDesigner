VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArmor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements cINode
Implements cIDisplay
Implements cIPersist

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


'Private m_Table() As Single 'todo: needed?


'local variable(s) to hold property value(s)
Private m_dblWeight As Double
Private m_dblCost As Double
Private m_dblSurfaceArea As Double
Private m_sngAverageDR As Single
Private m_sngAveragePD As Single


Private m_lngArmorType As Long ' e.g. overall, location, component, basic, complex, wheel,skirt,shield


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
    Debug.Print TypeName(Me) & ":cIDisplay:getFirstPropertyItem -- no properties in m_oProperties() array."
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

'//Implemented cIComponent Properties and Functions
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


'///////////////////////////////////////////
'//cIPersist Interface
Private Property Get cIPersist_Classname() As String
    cIPersist_Classname = TypeName(Me)
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
        m_lngPropCount = op.Load(XML_NODE_PROPERTYCOUNT)
        'm_Table = op.Load(XML_NODE_STATSTABLE)
        If m_lngPropCount > 0 Then
            ReDim m_oProperties(m_lngPropCount - 1)
            For i = 0 To m_lngPropCount - 1
                Set m_oProperties(i) = op.Load("property" & i)
            Next
        End If
    End If
End Sub
Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
End Sub
'//

Private Sub Class_Terminate()
    Dim i As Long
    For i = 0 To m_lngChildCount - 1
        Set m_oChildren(i) = Nothing
    Next
    
    For i = 0 To m_lngPropCount - 1
        Set m_oProperties(i) = Nothing
    Next
End Sub


Public Sub Init()


    Select Case mvarDatatype
        
          Case ArmorComplexFacing
            Material = "wood"
            Quality = "standard"
            DR = 1
            
            Material1 = "wood"
            Material2 = "wood"
            Material3 = "wood"
            Material4 = "wood"
            Material5 = "wood"
            Material6 = "wood"
            Quality1 = "standard"
            Quality2 = "standard"
            Quality3 = "standard"
            Quality4 = "standard"
            Quality5 = "standard"
            Quality6 = "standard"
            DR1 = 1
            DR2 = 1
            DR3 = 1
            DR4 = 1
            DR5 = 1
            If Veh.Components(mvarParent).Datatype = Body Then
                DR6 = 1
            Else
                DR6 = 0
            End If
        Case ArmorBasicFacing
           
            Material = "wood"
            Quality = "standard"
            DR1 = 1
            DR2 = 1
            DR3 = 1
            DR4 = 1
            DR5 = 1
            If Veh.Components(mvarParent).Datatype = Body Then
                DR6 = 1
            Else
                DR6 = 0
            End If
        Case ArmorOpenFrame
           
            Material = "wood"
            Quality = "standard"
            DR = 1
    
        Case ArmorGunShield
           
            Material = "wood"
            Quality = "standard"
            DR = 1
    
        Case ArmorLocation
           
            Material = "wood"
            Quality = "standard"
            DR = 1
    
        Case ArmorComponent
            
            Material = "wood"
            Quality = "standard"
            DR = 1
            
        Case ArmorOverall
            
            Material = "wood"
            Quality = "standard"
            DR = 1
    
        Case ArmorWheelGuard
            Material = "wood"
            Quality = "standard"
            DR = 1
            
    End Select
    
    mvarTL = Veh.Components(mvarParent).TL

End Sub



Public Sub StatsUpdate()
Dim component As Integer
Dim SlopeR As String
Dim SlopeL As String
Dim SlopeF As String
Dim SlopeB As String
Dim sCompare(1 To 6) As String
Dim sSides() As String
Dim i As Long
Dim j As Long
Dim count As Long

mvarZZInit = 1
mvarPrintOutput = "" ' reinit this var

mvarLocation = GetLocation

component = Veh.Components(mvarParent).Datatype

If (mvarDatatype = ArmorBasicFacing) Or (mvarDatatype = ArmorComplexFacing) Then
    SlopeR = Veh.Components(mvarParent).SlopeR
    SlopeL = Veh.Components(mvarParent).SlopeL
    SlopeF = Veh.Components(mvarParent).SlopeF
    SlopeB = Veh.Components(mvarParent).SlopeB
    
    
    CalcByFacingArmorWeightCost
    'get the PD
    ' Get the PD for each side (used for both Complex and Basic)
    mvarPD1 = CalcPD(mvarDR1, SlopeR, mvarMaterial1)
    mvarPD2 = CalcPD(mvarDR2, SlopeL, mvarMaterial2)
    mvarPD3 = CalcPD(mvarDR3, SlopeF, mvarMaterial3)
    mvarPD4 = CalcPD(mvarDR4, SlopeB, mvarMaterial4)
    mvarPD5 = CalcPD(mvarDR5, "none", mvarMaterial5) 'the top and underside dont have slope
    mvarPD6 = CalcPD(mvarDR6, "none", mvarMaterial6)
        
    'get the effective DR
    mvarEffectiveDR1 = CalcEffectiveDR(mvarDR1, SlopeR)
    mvarEffectiveDR2 = CalcEffectiveDR(mvarDR2, SlopeL)
    mvarEffectiveDR3 = CalcEffectiveDR(mvarDR3, SlopeF)
    mvarEffectiveDR4 = CalcEffectiveDR(mvarDR4, SlopeB)
    mvarEffectiveDR5 = mvarDR5
    mvarEffectiveDR6 = mvarDR6
    
        
ElseIf mvarDatatype = ArmorLocation Then
    
    CalcArmorWeightCost
    If component = Body Or component = Superstructure Or component = Turret Or component = Popturret Then
        
        SlopeR = Veh.Components(mvarParent).SlopeR
        SlopeL = Veh.Components(mvarParent).SlopeL
        SlopeF = Veh.Components(mvarParent).SlopeF
        SlopeB = Veh.Components(mvarParent).SlopeB
        ' Get the PD for each side (used for both Complex and Basic)
        mvarPD1 = CalcPD(mvarDR, SlopeR, mvarMaterial)
        mvarPD2 = CalcPD(mvarDR, SlopeL, mvarMaterial)
        mvarPD3 = CalcPD(mvarDR, SlopeF, mvarMaterial)
        mvarPD4 = CalcPD(mvarDR, SlopeB, mvarMaterial)
        mvarPD5 = CalcPD(mvarDR, "none", mvarMaterial)
        mvarPD6 = CalcPD(mvarDR, "none", mvarMaterial)
    Else
        mvarPD = CalcPD(mvarDR, "none", mvarMaterial)
    End If
Else
    CalcArmorWeightCost
    mvarPD = CalcPD(mvarDR, "none", mvarMaterial)
End If

'print output
If mvarRAP Then
End If
If mvarElectrified Then
End If
If mvarThermal Then
End If
If mvarRadiation Then
End If
If mvarCoating <> "none" Then
End If
If mvarPD Then
End If

Select Case mvarDatatype
    Case ArmorComplexFacing
        'can have different everything
        'get each side
        sCompare(1) = " PD " + Format(mvarPD1) + ", DR " + Format(mvarDR1) + " " + mvarQuality1 + " " + mvarMaterial1 + ". "
        sCompare(2) = " PD " + Format(mvarPD2) + ", DR " + Format(mvarDR2) + " " + mvarQuality2 + " " + mvarMaterial2 + ". "
        sCompare(3) = " PD " + Format(mvarPD3) + ", DR " + Format(mvarDR3) + " " + mvarQuality3 + " " + mvarMaterial3 + ". "
        sCompare(4) = " PD " + Format(mvarPD4) + ", DR " + Format(mvarDR4) + " " + mvarQuality4 + " " + mvarMaterial4 + ". "
        sCompare(5) = " PD " + Format(mvarPD5) + ", DR " + Format(mvarDR5) + " " + mvarQuality5 + " " + mvarMaterial5 + ". "
        sCompare(6) = " PD " + Format(mvarPD6) + ", DR " + Format(mvarDR6) + " " + mvarQuality6 + " " + mvarMaterial6 + ". "
        'move the first side (the Right) into the sSides array
        ReDim sSides(2, 1)
        sSides(1, 1) = sCompare(1)
        sSides(2, 1) = "R"
        count = 1
        'compare each side to see if it can be grouped with another
        For j = 2 To 6
            For i = 1 To count
                If sCompare(j) = sSides(1, i) Then
                    sSides(1, i) = sCompare(j)
                    sSides(2, i) = sSides(2, i) + "," + GetSideLetterFromNumber(j)
                ElseIf i = count Then
                    ReDim Preserve sSides(2, i + 1)
                    count = count + 1
                    sSides(1, i + 1) = sCompare(j)
                    sSides(2, i + 1) = GetSideLetterFromNumber(j)
                End If
            Next
        Next
        'get final string and include the surface options to the armor
        For i = 1 To count
            mvarPrintOutput = mvarPrintOutput + " " + sSides(2, i) + ": " + sSides(1, i)
        Next
        mvarPrintOutput = mvarPrintOutput + " (" + Format(mvarWeight, p_sFormat) + " lbs., $" + Format(mvarCost, p_sFormat) + ")."
        
    Case ArmorBasicFacing
        'same material and quality but different DR's and PD's
        sCompare(1) = " PD " + Format(mvarPD1) + ", DR " + Format(mvarDR1) + " "
        sCompare(2) = " PD " + Format(mvarPD2) + ", DR " + Format(mvarDR2) + " "
        sCompare(3) = " PD " + Format(mvarPD3) + ", DR " + Format(mvarDR3) + " "
        sCompare(4) = " PD " + Format(mvarPD4) + ", DR " + Format(mvarDR4) + " "
        sCompare(5) = " PD " + Format(mvarPD5) + ", DR " + Format(mvarDR5) + " "
        sCompare(6) = " PD " + Format(mvarPD6) + ", DR " + Format(mvarDR6) + " "
        'move the first side (the Right) into the sSides array
        ReDim sSides(2, 1)
        sSides(1, 1) = sCompare(1)
        sSides(2, 1) = "R"
        count = 1
        'compare each side to see if it can be grouped with another
        For j = 2 To 6
            For i = 1 To count
                If sCompare(j) = sSides(1, i) Then
                    sSides(1, i) = sCompare(j)
                    sSides(2, i) = sSides(2, i) + "," + GetSideLetterFromNumber(j)
                ElseIf i = count Then
                    ReDim Preserve sSides(2, i + 1)
                    count = count + 1
                    sSides(1, i + 1) = sCompare(j)
                    sSides(2, i + 1) = GetSideLetterFromNumber(j)
                End If
            Next
        Next
        'get final string and include the surface options to the armor
        For i = 1 To count
            mvarPrintOutput = mvarPrintOutput + " " + sSides(2, i) + ": " + sSides(1, i)
        Next
        mvarPrintOutput = mvarQuality + " " + mvarMaterial + mvarPrintOutput + " (" + Format(mvarWeight, p_sFormat) + " lbs., $" + Format(mvarCost, p_sFormat) + ")."
 
    
    Case ArmorLocation 'this can still have different PD's on Body and Turrets do to slope differences
        sCompare(1) = " PD " + Format(mvarPD1) + ", DR " + Format(mvarDR) + " "
        sCompare(2) = " PD " + Format(mvarPD2) + ", DR " + Format(mvarDR) + " "
        sCompare(3) = " PD " + Format(mvarPD3) + ", DR " + Format(mvarDR) + " "
        sCompare(4) = " PD " + Format(mvarPD4) + ", DR " + Format(mvarDR) + " "
        sCompare(5) = " PD " + Format(mvarPD5) + ", DR " + Format(mvarDR) + " "
        sCompare(6) = " PD " + Format(mvarPD6) + ", DR " + Format(mvarDR) + " "
        'move the first side (the Right) into the sSides array
        ReDim sSides(2, 1)
        sSides(1, 1) = sCompare(1)
        sSides(2, 1) = "R"
        count = 1
        'compare each side to see if it can be grouped with another
        For j = 2 To 6
            For i = 1 To count
                If sCompare(j) = sSides(1, i) Then
                    sSides(1, i) = sCompare(j)
                    sSides(2, i) = sSides(2, i) + "," + GetSideLetterFromNumber(j)
                ElseIf i = count Then
                    ReDim Preserve sSides(2, i + 1)
                    count = count + 1
                    sSides(1, i + 1) = sCompare(j)
                    sSides(2, i + 1) = GetSideLetterFromNumber(j)
                End If
            Next
        Next
        'get final string and include the surface options to the armor
        mvarPrintOutput = "DR " + Format(mvarDR) + " "
        For i = 1 To count
            mvarPrintOutput = mvarPrintOutput + " " + sSides(2, i) + ": " + sSides(1, i)
        Next
        mvarPrintOutput = mvarPrintOutput + mvarQuality + " " + mvarMaterial + " (" + Format(mvarWeight, p_sFormat) + " lbs., $" + Format(mvarCost, p_sFormat) + ")."
 
 
    Case ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
        mvarPrintOutput = "PD " + Format(mvarPD) + ", DR " + Format(mvarDR) + " " + mvarQuality + " " + mvarMaterial + " (" + Format(mvarWeight, p_sFormat) + " lbs., $" + Format(mvarCost, p_sFormat) + ")."
 
End Select

End Sub



Public Function FillMaterial() As String()
' populate the material combo
Dim materialarray() As String
ReDim materialarray(1)

    If mvarTL <= 6 Then
            materialarray = mAddKey(materialarray, "wood")
            materialarray = mAddKey(materialarray, "metal")
            materialarray = mAddKey(materialarray, "nonrigid")
    Else ' if its greater than or equal to 7
            materialarray = mAddKey(materialarray, "wood")
            materialarray = mAddKey(materialarray, "metal")
            materialarray = mAddKey(materialarray, "ablative")
            materialarray = mAddKey(materialarray, "fireproof ablative")
            materialarray = mAddKey(materialarray, "nonrigid")
            materialarray = mAddKey(materialarray, "composite")
            materialarray = mAddKey(materialarray, "laminate")
    End If

FillMaterial = materialarray
End Function

Public Function FillQuality(sMaterial As String) As String()
Dim MaterialCombo As ComboBox
Dim Selected As String ' holds the users selected Armor material
Dim arrQuality() As Integer 'holds the list of suitable quality
Dim iSelected As Integer 'holds converted Selected string
Dim element As Variant 'one element of the arrQuality array
Dim i As Integer 'counter
Dim count As Integer 'another counter
Dim qualityarray() As String
Dim TempTL As Integer

ReDim qualityarray(1)

Const Cheap = 1
Const Standard = 2
Const Expensive = 3
Const Advanced = 4

'get the type of armor that the user selected
Selected = sMaterial
'convert the Selected into an integer
Select Case Selected
    Case "wood"
        iSelected = 1
    Case "metal"
        iSelected = 2
    Case "ablative"
        iSelected = 3
    Case "fireproof ablative"
        iSelected = 4
    Case "nonrigid"
        iSelected = 5
    Case "composite"
        iSelected = 6
    Case "laminate"
        iSelected = 7
End Select

count = 1 ' init the counter
'given the tech level, produce list of quality types
TempTL = Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
TempTL = Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
For i = 1 To UBound(ArmorMatrix)
    If ArmorMatrix(i).TL = TempTL Then
        If ArmorMatrix(i).MaterialType = iSelected Then
            If ArmorMatrix(i).WeightMod <> 0 Then
                ReDim Preserve arrQuality(count)
                arrQuality(count) = ArmorMatrix(i).Quality
                count = count + 1
                If count > 5 Then Exit For Else
            End If
        End If
    End If
Next

'fill the Quality combo with the list of available items
For Each element In arrQuality
    Select Case element
        Case Cheap
            qualityarray = mAddKey(qualityarray, "cheap")
            'If TempText = "cheap" Then TextFlag = True
        Case Standard
            qualityarray = mAddKey(qualityarray, "standard")
            'If TempText = "standard" Then TextFlag = True
        Case Expensive
            qualityarray = mAddKey(qualityarray, "expensive")
            'If TempText = "expensive" Then TextFlag = True
        Case Advanced
            qualityarray = mAddKey(qualityarray, "advanced")
            'If TempText = "advanced" Then TextFlag = True
    End Select
Next

FillQuality = qualityarray

End Function

Sub CalcArmorWeightCost()
Dim TempTL As Integer
' This routine calculates the Cost and Weight of the armor.
    'These contstant values must match those in the module "modArmor" since
    'the armormatrix uses integers and not string names for the material and quality
    Const Cheap = 1
    Const Standard = 2
    Const Expensive = 3
    Const Advanced = 4
    
    Const Wood = 1
    Const Metal = 2
    Const Ablative = 3
    Const FireproofAblative = 4
    Const NonRigid = 5
    Const Composite = 6
    Const Laminate = 7

    Dim Area As Single 'holds the surface area
    Dim CostModifier As Single
    Dim WeightModifier As Single
    Dim SelectedMaterial As Integer
    Dim SelectedQuality As Integer
    Dim i As Integer ' counter
    
    ' Get the surface area based on the armor datatype being used
    If (mvarDatatype = ArmorWheelGuard) Then
        Area = Veh.Components(mvarParent).SurfaceArea / 2
    ElseIf (mvarDatatype = ArmorGunShield) Or (mvarDatatype = ArmorOpenFrame) Then
        Area = Veh.Components(mvarParent).SurfaceArea / 5
    ElseIf (mvarDatatype = ArmorLocation) Or (mvarDatatype = ArmorComponent) Then
        Area = Veh.Components(mvarParent).SurfaceArea
    ElseIf mvarDatatype = ArmorOverall Then
        Area = Veh.Stats.StructuralSurfaceArea
    End If
    
    'Determine Material
    If mvarMaterial = "wood" Then
        SelectedMaterial = Wood
    ElseIf mvarMaterial = "metal" Then SelectedMaterial = Metal
    ElseIf mvarMaterial = "ablative" Then SelectedMaterial = Ablative
    ElseIf mvarMaterial = "fireproof ablative" Then SelectedMaterial = FireproofAblative
    ElseIf mvarMaterial = "nonrigid" Then SelectedMaterial = NonRigid
    ElseIf mvarMaterial = "composite" Then SelectedMaterial = Composite
    ElseIf mvarMaterial = "laminate" Then SelectedMaterial = Laminate
    End If
        
    'Determine Quality
    If mvarQuality = "cheap" Then
        SelectedQuality = Cheap
    ElseIf mvarQuality = "standard" Then SelectedQuality = Standard
    ElseIf mvarQuality = "expensive" Then SelectedQuality = Expensive
    ElseIf mvarQuality = "advanced" Then SelectedQuality = Advanced
    End If

   ' Get the Cost and Weight Modifiers
       TempTL = Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
       TempTL = Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
       For i = 1 To UBound(ArmorMatrix)
           If ArmorMatrix(i).TL = TempTL Then
               If ArmorMatrix(i).MaterialType = SelectedMaterial Then
                   If ArmorMatrix(i).Quality = SelectedQuality Then
                       CostModifier = ArmorMatrix(i).Cost
                       WeightModifier = ArmorMatrix(i).WeightMod
                       Exit For
                    End If
                End If
           End If
       Next
          
       
    mvarWeight = mvarDR * Area * WeightModifier
    mvarCost = mvarWeight * CostModifier
    
    'get the final weight and cost by adding the cost/weight of the surface features
    CalcSurfaceFeaturesCostandWeight Area
End Sub

Sub CalcByFacingArmorWeightCost()
' This routine calculates the Cost and Weight of the armor
Dim Area As Single 'holds the surface area
Dim CostModifier(0 To 6) As Single
Dim WeightModifier(0 To 6) As Single
Dim iWeight(0 To 6) As Single
Dim iCost(0 To 6) As Single
Dim SelectedMaterial(0 To 6) As String
Dim SelectedQuality(0 To 6) As String
Dim iSelectedQuality(0 To 6) As Integer
Dim iSelectedMaterial(0 To 6) As Integer
Dim count As Integer
Dim i As Integer
Dim arrMaterial(0 To 6) As String
Dim arrQuality(0 To 6) As String
Dim arrDR(0 To 6) As Long
Dim TempCost As Single
Dim TempWeight As Single
Dim TempTL As Integer
'fill the arrMaterial array and arrQuality
arrMaterial(0) = mvarMaterial1
arrMaterial(1) = mvarMaterial2
arrMaterial(2) = mvarMaterial3
arrMaterial(3) = mvarMaterial4
arrMaterial(4) = mvarMaterial5
arrMaterial(5) = mvarMaterial6

arrQuality(0) = mvarQuality1
arrQuality(1) = mvarQuality2
arrQuality(2) = mvarQuality3
arrQuality(3) = mvarQuality4
arrQuality(4) = mvarQuality5
arrQuality(5) = mvarQuality6

'fill the arrDR aray
arrDR(0) = mvarDR1
arrDR(1) = mvarDR2
arrDR(2) = mvarDR3
arrDR(3) = mvarDR4
arrDR(4) = mvarDR5
arrDR(5) = mvarDR6

' re-init variables
TempCost = 0
TempWeight = 0

Area = Veh.Components(mvarParent).SurfaceArea
' There are just two paths, one for Complex and one for Basic
Select Case mvarDatatype
    Case ArmorComplexFacing
        For count = 0 To 5
            ' Get the quality and material of the armor
            SelectedMaterial(count) = arrMaterial(count)
            SelectedQuality(count) = arrQuality(count)
            'convert the Selected into an integer
            Select Case SelectedMaterial(count)
                Case "wood"
                    iSelectedMaterial(count) = 1
                Case "metal"
                    iSelectedMaterial(count) = 2
                Case "ablative"
                    iSelectedMaterial(count) = 3
                Case "fireproof ablative"
                    iSelectedMaterial(count) = 4
                Case "nonrigid"
                    iSelectedMaterial(count) = 5
                Case "composite"
                    iSelectedMaterial(count) = 6
                Case "laminate"
                    iSelectedMaterial(count) = 7
            End Select
            Select Case SelectedQuality(count)
                Case "cheap"
                    iSelectedQuality(count) = 1
                Case "standard"
                    iSelectedQuality(count) = 2
                Case "expensive"
                    iSelectedQuality(count) = 3
                Case "advanced"
                    iSelectedQuality(count) = 4
            End Select
            ' Get the Cost and Weight Modifiers
                TempTL = Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
                TempTL = Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
                For i = 1 To UBound(ArmorMatrix)
                    If ArmorMatrix(i).TL = TempTL Then
                        If ArmorMatrix(i).MaterialType = iSelectedMaterial(count) Then
                            If ArmorMatrix(i).Quality = iSelectedQuality(count) Then
                                CostModifier(count) = ArmorMatrix(i).Cost
                                WeightModifier(count) = ArmorMatrix(i).WeightMod
                                Exit For
                             End If
                         End If
                    End If
                Next
                'get the average DR
                CalcAverageDr
                ' Get the Cost and weight of each face
                If mvarParent = "1_" Then
                    iWeight(count) = Val(arrDR(count)) * (Area / 6) * WeightModifier(count)
                Else
                    iWeight(count) = Val(arrDR(count)) * (Area / 5) * WeightModifier(count)
                End If
                iCost(count) = iWeight(count) * CostModifier(count)
                TempWeight = TempWeight + iWeight(count)
                TempCost = TempCost + iCost(count)
            Next
            
    Case ArmorBasicFacing

        'convert the Selected into an integer
        Select Case mvarMaterial
            Case "wood"
                iSelectedMaterial(0) = 1
            Case "metal"
                iSelectedMaterial(0) = 2
            Case "ablative"
                iSelectedMaterial(0) = 3
            Case "fireproof ablative"
                iSelectedMaterial(0) = 4
            Case "nonrigid"
                iSelectedMaterial(0) = 5
            Case "composite"
                iSelectedMaterial(0) = 6
            Case "laminate"
                iSelectedMaterial(0) = 7
        End Select
        Select Case mvarQuality
            Case "cheap"
                iSelectedQuality(0) = 1
            Case "standard"
                iSelectedQuality(0) = 2
            Case "expensive"
                iSelectedQuality(0) = 3
            Case "advanced"
                iSelectedQuality(0) = 4
       End Select
       ' Get the Cost and Weight Modifiers
       TempTL = Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
       TempTL = Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
       For i = 1 To UBound(ArmorMatrix)
           If ArmorMatrix(i).TL = TempTL Then
               If ArmorMatrix(i).MaterialType = iSelectedMaterial(0) Then
                   If ArmorMatrix(i).Quality = iSelectedQuality(0) Then
                       CostModifier(0) = ArmorMatrix(i).Cost
                       WeightModifier(0) = ArmorMatrix(i).WeightMod
                    End If
                End If
           End If
       Next
       'call routine to calc averagedr
       CalcAverageDr
       ' Get the Final Cost and Final Weight
       TempWeight = AverageDR * Area * WeightModifier(0)
       TempCost = TempWeight * CostModifier(0)
End Select


'save these cost and weight results to the armor class
mvarCost = TempCost
mvarWeight = TempWeight

'get the final cost which includes the cost of the surface features
CalcSurfaceFeaturesCostandWeight Area
End Sub


Function CalcEffectiveDR(DR As Long, Slope As String)
Dim Modifier As Single

If Slope = "none" Then
    Modifier = 1
ElseIf Slope = "30 degrees" Then
    Modifier = 1.5
ElseIf Slope = "60 degrees" Then
    Modifier = 2
End If

CalcEffectiveDR = DR * Modifier
End Function


Function GetLowestDR() As Long
    '//this function only gets called during Aerial performance calculations.
    '//its job is to return the DR of the armor. In the case of seperate DR's
    '//for each face, then  it will return the lowest one.
    '//only DR from metal, composite or laminate armor counts.. all other types
    '//return 0.
    On Error Resume Next
    Dim lngRetval As Long
    
    Select Case mvarDatatype
       Case ArmorComplexFacing
            Select Case mvarMaterial1
                Case "metal", "composite", "laminate"
                    lngRetval = mvarDR1
                    Select Case mvarMaterial2
                        Case "metal", "composite", "laminate"
                            lngRetval = Minimum(mvarDR1, mvarDR2)
                            Select Case mvarMaterial3
                                Case "metal", "composite", "laminate"
                                    lngRetval = Minimum(lngRetval, mvarDR3)
                                    Select Case mvarMaterial4
                                        Case "metal", "composite", "laminate"
                                            lngRetval = Minimum(lngRetval, mvarDR4)
                                            Select Case mvarMaterial5
                                                Case "metal", "composite", "laminate"
                                                    lngRetval = Minimum(lngRetval, mvarDR5)
                                                    Select Case mvarMaterial6
                                                        Case "metal", "composite", "laminate"
                                                            lngRetval = Minimum(lngRetval, mvarDR6)
                                                        Case Else
                                                            lngRetval = 0
                                                    End Select
                                                Case Else
                                                    lngRetval = 0
                                            End Select
                                        Case Else
                                            lngRetval = 0
                                    End Select
                                Case Else
                                    lngRetval = 0
                            End Select
                        Case Else
                            lngRetval = 0
                    End Select
                Case Else
                    lngRetval = 0
            End Select
           
       Case ArmorBasicFacing
           Select Case mvarMaterial
                Case "metal", "composite", "laminate"
                    lngRetval = Minimum(mvarDR1, mvarDR2)
                    lngRetval = Minimum(lngRetval, mvarDR3)
                    lngRetval = Minimum(lngRetval, mvarDR4)
                    lngRetval = Minimum(lngRetval, mvarDR5)
                    lngRetval = Minimum(lngRetval, mvarDR6)
                
                Case Else
                    lngRetval = 0
            End Select
       
       Case ArmorLocation, ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
            Select Case mvarMaterial
                Case "metal", "composite", "laminate"
                    lngRetval = mvarDR
                Case Else
                    lngRetval = 0
            End Select
          
    End Select


    GetLowestDR = lngRetval
End Function

Function GetLowestCrushDepthDR() As Long

    On Error Resume Next
    Dim lngRetval As Long
    
    Select Case mvarDatatype
       Case ArmorComplexFacing
           
                lngRetval = mvarDR1
                lngRetval = Minimum(mvarDR1, mvarDR2)
                lngRetval = Minimum(lngRetval, mvarDR3)
                lngRetval = Minimum(lngRetval, mvarDR4)
                lngRetval = Minimum(lngRetval, mvarDR5)
                lngRetval = Minimum(lngRetval, mvarDR6)
    
       Case ArmorBasicFacing
          
                lngRetval = Minimum(mvarDR1, mvarDR2)
                lngRetval = Minimum(lngRetval, mvarDR3)
                lngRetval = Minimum(lngRetval, mvarDR4)
                lngRetval = Minimum(lngRetval, mvarDR5)
                lngRetval = Minimum(lngRetval, mvarDR6)
    
       
       Case ArmorLocation, ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
               lngRetval = mvarDR
              
    End Select


    GetLowestCrushDepthDR = lngRetval
End Function

Private Sub CalcAverageDr()
    Dim i As Integer
    Dim TempAverage As Long
    Dim Divisor As Integer
    Dim arrDR(0 To 5) As Long
    
    arrDR(0) = Val(mvarDR1)
    arrDR(1) = Val(mvarDR2)
    arrDR(2) = Val(mvarDR3)
    arrDR(3) = Val(mvarDR4)
    arrDR(4) = Val(mvarDR5)
    arrDR(5) = Val(mvarDR6)
    
    TempAverage = 0
    
    For i = 0 To 5
        TempAverage = TempAverage + arrDR(i)
    Next
    
    'find the Average DR (this is done only for Basic Armor by facing
    If mvarParent = "1_" Then
        Divisor = 6
    Else
        Divisor = 5
    End If
    
    mvarAverageDR = Round(TempAverage / Divisor, 2)

End Sub



Public Property Get Weight() As Double
    Weight = 20.2  'm_dblWeight
End Property
Public Property Get Cost() As Double
    Cost = 1234.56 ' m_dblCost
End Property
Public Property Get SurfaceArea() As Double
    SurfaceArea = 1.9 ' m_dblSurfaceArea
End Property
Public Property Get AverageDR() As Single
    AverageDR = m_sngAverageDR
End Property
Public Property Get AveragePD() As Single
    AveragePD = m_sngAveragePD 'todo: i guess this would mean, average across faces? or maybe its irrelevant?  i should delete this
End Property
