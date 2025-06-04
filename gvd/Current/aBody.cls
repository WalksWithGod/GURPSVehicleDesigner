VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "aBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements cIPersist
Implements cINode
Implements cIDisplay
Implements cIContainer
Implements cIComponent
Implements cIBuild

' -------cIBuild related variables
Private Type GVD_DATA_TABLE
    ID As Long
    ptrTable As Long
End Type
Private Type GVD_OPTIONS
    index As Long
    selectionCount As Long
    ptrTable As Long
End Type
Private Type GVD_USER_INPUT
    sngValue As Single
    sngURange As Single
    sngLRange As Single
End Type
Private Type GVD_FORMULA
    lngStatID As Long
    lngFormulaID As Long
End Type
Private m_Tables() As GVD_DATA_TABLE
Private m_Options() As GVD_OPTIONS
Private m_UserInput() As GVD_USER_INPUT
Private m_Formulas() As GVD_FORMULA
Private m_lngTableCount As Long
Private m_lngOptionCount As Long
Private m_lngUserInputCount As Long
Private m_lngFormulaCount As Long

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

' ------- cIComponent interface variables
Private m_sngTable() As Single
Private m_sngTL As Single
Private m_dblCost As Double
Private m_dblWeight As Double
Private m_dblVolume As Double
Private m_dblSurfaceArea As Double
Private m_dblHitpoints As Double
Private m_lngDR As Long 'todo: move DR back to cIComponent???

Private m_oArmor As cArmor



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
    If Not m_oProperties(0) Is Nothing Then
        Set cIDisplay_getFirstPropertyItem = m_oProperties(0)
        m_lngCurrentPropItem = 0
    End If
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

'///////////////////////////////////////////////////
'//cINode Implemented Properties and Functions
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

'///////////////////////////////////////////////////
'//cIComponent Implemented Properties and Functions
Private Property Let cIComponent_LogicalParent(ByVal RHS As Long)
End Property
Private Property Get cIComponent_LogicalParent() As Long
End Property
Private Property Let cIComponent_TL(ByVal RHS As Single)
    m_sngTL = RHS
End Property
Private Property Get cIComponent_TL() As Single
    cIComponent_TL = m_sngTL
End Property
Private Property Let cIComponent_SurfaceArea(ByVal RHS As Double)
    m_dblSurfaceArea = RHS
End Property
Private Property Get cIComponent_SurfaceArea() As Double
    cIComponent_SurfaceArea = m_dblSurfaceArea
End Property
Private Property Let cIComponent_Cost(ByVal RHS As Double)
   
End Property
Private Property Get cIComponent_Cost() As Double
    cIComponent_Cost = m_dblCost
End Property
Private Property Let cIComponent_Volume(ByVal RHS As Double)
    'todo: delete let?
End Property
Private Property Get cIComponent_Volume() As Double
    cIComponent_Volume = m_dblVolume
End Property
Private Property Let cIComponent_Weight(ByVal RHS As Double)
    ' delete let?
End Property
Private Property Get cIComponent_Weight() As Double
    cIComponent_Weight = 300.544442191 'todo: temp val m_dblWeight
End Property
Private Property Let cIComponent_HitPoints(ByVal RHS As Double)
    'delete let
End Property
Private Property Get cIComponent_HitPoints() As Double
    cIComponent_HitPoints = m_dblHitpoints
End Property


'///////////////////////////////////////////////////
'//cIBuild Implemented Properties and Functions
Private Function cIBuild_getOption(ByVal lngIndex As Long) As Long
    On Error GoTo err
    cIBuild_getOption = m_Options(lngIndex).index
    Exit Function
err:
    Debug.Print TypeName(Me) & ":cIBuild_getOption -- ERROR #" & err.Number & " " & err.Description
End Function
Private Function cIBuild_setOption(ByVal lngIndex As Long, ByVal lngSelection As Long) As Boolean
On Error GoTo err
    ' before assigning the value, check that the selection is valid by determining if its in the range of 0 to (SelectCount -1)
    If (lngSelection <= m_Options(lngIndex).selectionCount - 1) Then
        m_Options(lngIndex).index = lngSelection
        cIBuild_setOption = True
    Else
        InfoPrint 1, TypeName(Me) & ":cIBuild_setOption() -- ERROR.  Selection invalid.  Are you a hacker?"
    End If
    Exit Function
err:
    Debug.Print TypeName(Me) & ":cIBuild_setOption -- ERROR #" & err.Number & " " & err.Description
    cIBuild_setOption = False
End Function
Private Function cIBuild_getUserInput(ByVal lngIndex As Long) As Single
    On Error GoTo err
    cIBuild_getUserInput = m_UserInput(lngIndex).sngValue
    Exit Function
err:
    Debug.Print TypeName(Me) & ":cIBuild_getUserInput -- ERROR #" & err.Number & " " & err.Description
End Function
Private Function cIBuild_setUserInput(ByVal lngIndex As Long, ByVal sngValue As Single) As Boolean
On Error GoTo err
    If (sngValue >= m_UserInput(lngIndex).sngLRange) And (sngValue <= m_UserInput(lngIndex).sngURange) Then
        m_UserInput(lngIndex).sngValue = sngValue
        cIBuild_setUserInput = True
    Else
        InfoPrint 1, "User input for this field limited to values between " & m_UserInput(lngIndex).sngLRange & " and " & m_UserInput(lngIndex).sngURange
        cIBuild_setUserInput = False
    End If
    Exit Function
err:
    Debug.Print TypeName(Me) & ":cIBuild_setUserInput -- ERROR #" & err.Number & " " & err.Description
    cIBuild_setUserInput = False
End Function
Private Function cIBuild_calcStats(ByRef oVisitor As cStats) As Boolean
    oVisitor.calcStats Me
    ' todo: actually before we can pass m_lngTL to the stat table, we need to know the bounds
    'm_dblCost = m_sngTable(0, 2)
    'm_dblWeight = m_sngTable(0, 0)
    'm_dblVolume = m_sngTable(0, 1) '+ AddedVolume
    'm_dblSurfaceArea = CalcSurfaceArea(m_dblVolume)
    'm_dblHitpoints = CalcComponentHitpoints(m_dblSurfaceArea)
End Function

'///////////////////////////////////////////////////
'//cIContainer Implemented Properties and Functions
Private Function cIContainer_AcceptObject(oComponent As cIComponent) As Long
End Function
Private Property Get cIContainer_ContainerAbbrev() As String
End Property
Private Property Get cIContainer_VolumeAccessSpace() As Double
End Property
Private Property Let cIContainer_VolumeAccessSpace(ByVal RHS As Double)
End Property
Private Property Get cIContainer_VolumeBattleSuit() As Double
End Property
Private Property Let cIContainer_VolumeBattleSuit(ByVal RHS As Double)
End Property
Private Property Get cIContainer_VolumeExtra() As Double
    cIContainer_VolumeExtra = 890
End Property
Private Property Let cIContainer_VolumeExtra(ByVal RHS As Double)
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

