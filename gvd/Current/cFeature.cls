VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'todo: shoudl make these features appear in the component drop down list and that they can only be added to a parent
' node of cSurface

Implements cIPersist
Implements cINode
Implements cIDisplay
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

Private Sub Class_Initialize()
    'todo: should i hardcode limits on m_lngMaxChildren so that def file will never accidentally override?
End Sub

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
'TODO: will leaf options ever accept children?  Probably not... just remember
' to investigate and then delete this code if its not ncessary, same for the getchild stuff
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
    cIPersist_Classname = TypeName(Me)
End Property
Private Property Get cIPersist_GUID() As String
End Property
Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Long)
    Dim i As Long
    
    If iMode = cmp Then
        m_sName = op.Load(XML_NODE_NAME)
        m_sDescription = op.Load(XML_NODE_DESCRIPTION)
        
        
        ' now load the saved option indexes and userinput values from the .cmp file
        If m_lngOptionCount > 0 Then
            ReDim m_Options(0 To m_lngOptionCount - 1)
            For i = 0 To m_lngOptionCount - 1
                m_Options(i).index = op.Load("option_value" & i)
            Next
        End If

        If m_lngUserInputCount > 0 Then
            ReDim m_UserInput(0 To m_lngUserInputCount - 1)
            'load in the min/max allowed ranges
            For i = 0 To m_lngUserInputCount - 1
                m_UserInput(i).sngValue = op.Load("userinput_value" & i)
            Next
        End If
        
        ' todo: call to update stats go here?
        ' If so, then no need to ever save/load the statistics since we recalc them.
        ' We need to make a decision here to either set a bChanged flag so that we can update later
        ' or we update the stats NOW, even though we are still loading  potentially, other vehicle nodes.
        ' Since this is a feature, calc'ing it now could result in wrong surface area being used.  If we use
        ' flags, we can specify bChanged and bPriority to indicate when its ok to calc stats for this node.
        ' we can also flag if bIndependant if these stats have zero requirements of other stats. (e.g. changing
        ' of body surface area still wont effect the weight or volume of a radio component)
        
        ' come to think of it, we have to use flags.  Just think about when adding a Turret, we can't calc stats
        ' right after loading cuz there are children that havent been added yet which will ultimately effect its
        ' stats.
        
        ' Recall that some components like wheels, you cant load its stats anyway since they are in fact
        ' dependant on the body's stats.  So this is yet another reason why we should flag here and re-calc
        ' rather than just load the stats and not make any calls to update or flag for update.
        
        ' So assuming we are going to set a flag here, what is the process?  Perhaps call to setFlag() internally
        ' makes call to notifyParent() which recurses up to the root vehicle.  Once the root vehicle is reached, the
        ' recursion can begin starting with the leaf nodes.  But what about sibling nodes.. wont that have some funky
        ' tendancy to recurse all the way to the vehicle node more than once?  I could end the recursion prematurely
        ' if i see the flag of a parent node is already set, therefore assuming the parent is already notified all the
        ' way to the top.
        
        ' And note that performances need to be updated virtually every change even though no child node will ever
        ' be able to notify it (same with features)  Thus these need to be done seperately.  Wheel too though
        ' it is on the components branch, but its dependant on body's stat.
        
        ' Dont forget about armor
        ' Dont forget about airbags and stabilization gear who's stats depend on the component they compliment
        
    Else 'DEF
        m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
        m_lngAttributes = op.Load("attributes")
        
        #If DEBUG_MODE Then
            Debug.Assert m_lngMaxChildren <= 0
        #End If
        m_sImage = op.Load(XML_NODE_IMAGE)
        
        m_lngTableCount = op.Load(XML_NODE_STATS_TABLECOUNT)
        If m_lngTableCount > 0 Then
        ReDim m_Tables(m_lngTableCount - 1)
            For i = 0 To m_lngTableCount - 1
                m_Tables(i).ptrTable = op.Load(XML_NODE_TABLE & i)
            Next
        End If
        
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
End Sub
Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)

    op.Store "name", m_sName
    
End Sub

