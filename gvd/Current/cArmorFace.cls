VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArmorFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements cIPersist
Implements cIDisplay
Implements cINode


Private m_lngPropCount As Long
Private m_lngCurrentPropItem As Long
Private m_oProperties() As cPropertyItem

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



Private m_lngTL As Long


' armor specific crap
Private m_bRAP As Boolean
Private m_bElectrified As Boolean
Private m_bThermalCoating As Boolean
Private m_bRadShielding As Boolean
Private m_ReflectiveCoating As String
Private m_lngPD As Long
Private m_lngDR As Long   '<--- todo: need more space? DR is cumlative in the "Face" since it adds all layer's DR
Private m_dblSurfaceArea As Double
Private m_dblWeight As Double
Private m_dblCost As Double

Private Sub Class_Terminate()
    Dim i As Long
    For i = 0 To m_lngChildCount - 1
        Set m_oChildren(i) = Nothing
    Next
    
    For i = 0 To m_lngPropCount - 1
        Set m_oProperties(i) = Nothing
    Next
End Sub

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


Private Function GetSideLetterFromNumber(i As Long) As String
    Select Case i
        Case 1
            GetSideLetterFromNumber = "R"
        Case 2
            GetSideLetterFromNumber = "L"
        Case 3
            GetSideLetterFromNumber = "F"
        Case 4
            GetSideLetterFromNumber = "B"
        Case 5
            GetSideLetterFromNumber = "T"
        Case 6
            GetSideLetterFromNumber = "U"
    End Select
End Function

Function CalcPD(DR As Long, Slope As String, Material As String) As Long
    Dim PD As Integer
    
    If DR = 0 Then
        PD = 0
    ElseIf DR = 1 Then
        PD = 1
    ElseIf DR <= 4 Then
        PD = 2
    ElseIf DR <= 15 Then
        PD = 3
    ElseIf DR >= 16 Then
        PD = 4
    End If
        
    'check for max values for Wood and Nonrigid armor
    If Material = "nonrigid" Then
        If PD > 2 Then PD = 2
    ElseIf Material = "wood" Then
        If PD > 3 Then PD = 3
    End If
    
    'add bonus for slope
    'Note: the bonus's are placed below the checks for  max values for wood and nonrigid.
    'if users request, it can be moved above it
    If Slope = "none" Then
    ElseIf Slope = "30 degrees" Then
        PD = PD + 1
    ElseIf Slope = "60 degrees" Then
        PD = PD + 2
    End If
        
    CalcPD = PD
End Function

Sub CalcSurfaceFeaturesCostandWeight(SurfaceArea As Single)
    Const RadWeight = 2
    Const RadCost = 20
    Const ReflectCost = 30
    Const RetroCost = 150
    Const ThermCost = 250
    Const ThermWeight = 0.25
    Const RAPCost = 20
    Const RAPWeight = 8
    Const ElectCost = 10
    Const ElectWeight = 0.2
    Dim TempWeight As Single
    Dim TempCost As Single
    
    If SurfaceArea = 0 Then Exit Sub
    
    If mvarRadiation Then
        TempWeight = RadWeight * SurfaceArea
        TempCost = RadCost * SurfaceArea
    End If
    
    If mvarCoating = "reflective" Then
        TempCost = TempCost + (ReflectCost * SurfaceArea)
    ElseIf mvarCoating = "retro-reflective" Then
        TempCost = TempCost + (RetroCost * SurfaceArea)
    End If
    
    If mvarThermal Then
        TempWeight = TempWeight + (ThermWeight * SurfaceArea)
        TempCost = TempCost + (ThermCost * SurfaceArea)
    End If
    
    If mvarRAP Then
        TempWeight = TempWeight + (RAPWeight * SurfaceArea)
        TempCost = TempCost + (RAPCost * SurfaceArea)
    End If
    
    If mvarElectrified Then
        TempWeight = TempWeight + (ElectWeight * SurfaceArea)
        TempCost = TempCost + (ElectCost * SurfaceArea)
    End If
    
    mvarCost = Round(mvarCost + TempCost, 2)
    mvarWeight = Round(mvarWeight + TempWeight, 2)

End Sub


Public Property Let RAP(ByVal b As Boolean)
    m_bRAP = b
End Property
Public Property Get RAP() As Boolean
    RAP = m_bRAP
End Property
Public Property Let Electrified(ByVal b As Boolean)
    m_bElectrified = b
End Property
Public Property Get Electrified() As Boolean
    Electrified = m_bElectrified
End Property
Public Property Let ThermalCoating(ByVal b As Boolean)
    m_bThermalCoating = b
End Property
Public Property Get ThermalCoating() As Boolean
    ThermalCoating = m_bThermalCoating
End Property
Public Property Let RadShielding(ByVal b As Boolean)
    m_bRadShielding = b
End Property
Public Property Get RadShielding() As Boolean
    RadShielding = m_bRadShielding
End Property
Public Property Let ReflectiveCoating(ByRef s As String)
    m_sReflectiveCoating = s
End Property
Public Property Get ReflectiveCoating() As String
    ReflectiveCoating = m_sReflectiveCoating
End Property
Public Property Let PD(ByVal i As Long)
    m_lngPD = i
End Property
Public Property Get PD() As Long
    PD = m_lngPD
End Property
Public Property Let DR(ByVal i As Long)
    m_lngDR = i
End Property
Public Property Get DR() As Long
    DR = m_lngDR
End Property
Public Property Get Weight(ByVal D As Double)
    Weight = m_dblWeight
End Property
Public Property Get Cost(ByVal D As Double)
    Cost = m_dblCost
End Property
Public Property Get SurfaceArea(ByVal D As Double)
    SurfaceArea = m_dblSurfaceArea
End Property
