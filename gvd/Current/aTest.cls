VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "aTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' NOTE: This is our reference test class.  From it we will increase functionality until we've covered all the
' basic interfaces and component types.  We will restrict most todo and ntoe comments to here for now
' as we hash out the design
Implements cIPersist
Implements cINode
Implements cIDisplay
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
' todo: since this wont implement cIContainer, it actually doesnt need child count or the child array
' after i get this thing to run and load into the tree, i will delete these since aTest is my base
' class for which to model all other cIComponents (non Container)
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
Private m_sngTL As Single  '<-- this is the only "stat" the user can modify directly
'      Note: that we cant actually "test" for an overflow without actually causing an overflow.
'      That is to say, our "test" might overflow too.  This puts us back to the max user input values
'      As well as max number of components in a vehicle.
'      Quantity must be taken into account too.  Still havent finalized how tohandlle that.
Private m_dblCost As Double
Private m_dblWeight As Double
Private m_dblVolume As Double
Private m_dblSurfaceArea As Double
Private m_dblHitpoints As Double

' ------- cIArmor interface variables todo: note i dont think this will be a seperate interface but rather
'         apart of cIcomponent.
Private m_lngDR As Long 'todo: this dr is probably obsolete since it will come directly from armor?  hrm, but
'whta about default DR for components that dont actually set component armor?
Private m_oArmor As cArmor



Private Sub Class_Terminate()
    Dim i As Long
    For i = 0 To m_lngChildCount - 1
        Set m_oChildren(i) = Nothing
    Next
    
    For i = 0 To m_lngPropCount - 1
        Set m_oProperties(i) = Nothing
    Next
    For i = 0 To m_lngTableCount - 1
        destroyTable m_Tables(i).ptrTable
    Next
    For i = 0 To m_lngOptionCount - 1
        destroyTable m_Options(i).ptrTable
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
    'todo: this is a leaf node, probably dont need any implementation just return false
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
    'todo: dont need implemenation, this is leaf right?
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
Private Property Let cIComponent_SurfaceArea(ByVal RHS As Double)
    m_dblSurfaceArea = RHS
End Property
Private Property Get cIComponent_SurfaceArea() As Double
    cIComponent_SurfaceArea = m_dblSurfaceArea
End Property
Private Property Let cIComponent_Cost(ByVal RHS As Double)
 'todo: delete all the lets since these are calculated internally?
    ' err... but not for custom components... perhaps a seperate interface for those right?
    ' these should all be referenced as "Added_XXXXX" e.g. Added_Cost
    ' seems like the best way to do it since we will be using function calls
    ' to calcStats ONLY when a stat specific variable is altered.
    
    ' Also, when modifying something like "Description" or "Notes" we dont want to call that
    ' function, but just update the print string.
    
End Property
Private Property Get cIComponent_Cost() As Double
End Property
Private Property Let cIComponent_Volume(ByVal RHS As Double)
End Property
Private Property Get cIComponent_Volume() As Double
End Property
Private Property Let cIComponent_Weight(ByVal RHS As Double)
End Property
Private Property Get cIComponent_Weight() As Double
End Property
Private Property Let cIComponent_HitPoints(ByVal RHS As Double)
    m_dblHitpoints = RHS
End Property
Private Property Get cIComponent_HitPoints() As Double
    cIComponent_HitPoints = m_dblHitpoints
End Property

' TL is mostly just a keystone build option modifier.
' So cIComponent is basically the key "build" interface. I think the only
' reason i dont include it as m_options() local to each class is that
' i need to set a default TL when a component is added to the tree.
' NOTE: I probably should have TL in cIBuild_TL AND cIComponent.  They'll both
' access the same internal variable, but this way we can access it from either interface
Private Property Let cIComponent_TL(ByVal RHS As Single)
    m_sngTL = RHS
End Property
Private Property Get cIComponent_TL() As Single
    cIComponent_TL = m_sngTL
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
'//cIPersist Implemented Properties and Functions
Private Property Get cIPersist_Classname() As String
    'todo: whats this property for?
    cIPersist_Classname = TypeName(Me)
End Property
Private Property Get cIPersist_GUID() As String
    
End Property

Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Long)
    Dim i As Long
    
    If iMode = cmp Then
        m_sName = op.Load(XML_NODE_NAME)
        m_sDescription = op.Load(XML_NODE_DESCRIPTION)
        'note: default values for options() and userinput() are always stored in the .cmp file and
        '      not the .def file
        'todo: load our user input saved values
        'todo: load our option saved values
        'todo: we need to gracefully resume if an option or userinput index
        '      is not represented... or do we require they all be in the .cmp even with 0 value?
        
    Else
        m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
        m_sImage = op.Load(XML_NODE_IMAGE)
                
        ' todo: testing of loading formulas, stats, and multipliers
        'm_lngFormulaCount = op.Load("formula_count")
        'ReDim m_Formulas(m_lngFormulaCount - 1)
        'For i = 0 to m_lngFormulaCount -1
        '   m_Formulas(i) = op.Load(XML_NODE_FORMULA)
        'Next
        m_lngTableCount = op.Load(XML_NODE_STATS_TABLECOUNT)
        ReDim m_Tables(m_lngTableCount - 1)
        For i = 0 To m_lngTableCount - 1
            m_Tables(i).ptrTable = op.Load(XML_NODE_TABLE & i)
        Next
        
        m_lngOptionCount = op.Load(XML_NODE_OPTION_MODIFER_TABLE_COUNT)
        If m_lngOptionCount > 0 Then
            ReDim m_Options(0 To m_lngOptionCount - 1)
            For i = 0 To m_lngOptionCount - 1
                m_Options(i).selectionCount = op.Load("option_selectioncount" & i)
                m_Options(i).ptrTable = op.Load("option_table" & i)
            Next
        End If

        m_lngUserInputCount = op.Load("userinput_count")
        If m_lngUserInputCount > 0 Then
            ReDim m_UserInput(0 To m_lngUserInputCount - 1)
            'load in the min/max allowed ranges
            For i = 0 To m_lngUserInputCount - 1
                m_UserInput(i).sngURange = op.Load("userinput_urange" & i)
                m_UserInput(i).sngLRange = op.Load("userinput_lrange" & i)
            Next
        End If
        
        ' load properties last, these will reference variables initialized above
        m_lngPropCount = op.Load(XML_NODE_PROPERTYCOUNT)
        If m_lngPropCount > 0 Then
            ReDim m_oProperties(m_lngPropCount - 1)
            For i = 0 To m_lngPropCount - 1
                Set m_oProperties(i) = op.Load(XML_NODE_PROPERTY & i)
            Next
        End If
    End If
    'todo: we need to abort if any of these loads fail right?  Like imagine if the
    '      stats tables dont load, thats disaster pretty much so we have to abort... means the def
    '      was corrupt.

    '      I could change this to a function and return TRUE if we make it through with no errors.
    
End Sub
Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
End Sub

