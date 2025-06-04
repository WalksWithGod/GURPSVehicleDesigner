VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements cIComponent
Implements cIDisplay
Implements cINode
Implements cIPersist

' -------cINode interface variables
Private m_lngMaxChildren As Long
Private m_lngChildCount As Long
Private m_lngCurrentChild As Long
Private m_oChildren() As cINode
Private m_hParent As Long
Private m_hMe As Long
Private m_sDescription As String
Private m_sImage As String

' -------cIDisplay interface variables
Private m_lngPropCount As Long
Private m_lngCurrentPropItem As Long
Private m_oProperties() As cPropertyItem

' ------- component interface properties
Private m_Table() As Single
Private m_lngTL As Long
Private m_dblHitpoints As Double
Private m_dblSurfaceArea As Double
Private m_dblCost As Double
Private m_dblVolume As Double
Private m_dblWeight As Double

'Public Stats As clsStats   'todo: these must all implement cIPersist
'Public Crew As clsCrew
'Public Surface As clsSurface
'Public Options As clsOptions
'Public Description As clsDescription


'Public PowerProfiles As Collection
'Public FuelProfiles As Collection
'Public WeaponProfiles As clsProfile
'Public PerformanceProfiles As clsProfile
'Public BatteryProfiles As clsProfile
'Public Profiles As clsProfile

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
'/////////////////////////////////////////////
'//Implemented cINode Properties and Functions
Private Function cINode_AddChild(oNode As cINode) As Boolean
    If m_lngMaxChildren = m_lngChildCount Then
        cINode_AddChild = False
    Else
        m_lngChildCount = m_lngChildCount + 1
        ReDim Preserve m_oChildren(m_lngChildCount)
        Set m_oChildren(m_lngChildCount) = oNode
        cINode_AddChild = True
    End If
End Function
Private Function cINode_getChildFromHandle(ByVal h As Long) As cINode
End Function
Private Function cINode_getChildrenByClassName(ByRef Classname As String, hChilds() As Long) As Boolean
End Function
Private Function cINode_getFirstChild() As cINode
    m_lngCurrentChild = 0
    If m_lngCurrentChild <= m_lngChildCount Then
        If Not m_oChildren(m_lngCurrentChild) Is Nothing Then
            Set cINode_getFirstChild = m_oChildren(m_lngCurrentChild) 'todo:
        End If
    End If
End Function
Private Function cINode_getNextChild() As cINode
    m_lngCurrentChild = m_lngCurrentChild + 1
    
    If m_lngCurrentChild <= m_lngChildCount - 1 Then
        If Not m_oChildren(m_lngCurrentChild) Is Nothing Then
            Set cINode_getNextChild = m_oChildren(m_lngCurrentChild)
        End If
    End If
End Function
Private Function cINode_RemoveChild(oNode As cINode) As Boolean
End Function
Private Property Get cINode_ClassName() As String
    cINode_ClassName = TypeName(Me)
End Property
Private Property Get cINode_GraphDown() As Boolean
    cINode_GraphDown = True
End Property
Private Property Let cINode_Parent(ByVal RHS As Long)
End Property
Private Property Get cINode_Parent() As Long
End Property
Private Property Get cINode_ContainerAbbrev() As String
End Property
Private Property Get cINode_AllowUserDelete() As Boolean
End Property
Private Property Let cINode_Description(ByVal RHS As String)
    m_sDescription = RHS
End Property
Private Property Get cINode_Description() As String
    cINode_Description = m_sDescription
End Property
Private Property Let cINode_Image(ByVal RHS As String)
    m_sImage = RHS
End Property
Private Property Get cINode_Image() As String
    cINode_Image = m_sImage
End Property
Private Property Let cINode_Note(ByVal RHS As String)
End Property
Private Property Get cINode_Note() As String
End Property
Private Property Let cINode_Handle(ByVal RHS As Long)
    m_hMe = RHS
End Property
Private Property Get cINode_Handle() As Long
    cINode_Handle = m_hMe
End Property




'//Implemented cIComponent Properties and Functions
Private Property Let cIComponent_LogicalParent(ByVal RHS As Long)
End Property
Private Property Get cIComponent_LogicalParent() As Long
End Property
Private Property Let cIComponent_TL(ByVal RHS As Long)
    m_lngTL = RHS
End Property
Private Property Get cIComponent_TL() As Long
    cIComponent_TL = m_lngTL
End Property
Private Property Let cIComponent_HitPoints(ByVal RHS As Double)
End Property
Private Property Get cIComponent_HitPoints() As Double
End Property
Private Property Let cIComponent_SurfaceArea(ByVal RHS As Double)
End Property
Private Property Get cIComponent_SurfaceArea() As Double
End Property
End Property
Private Property Let cIComponent_Cost(ByVal RHS As Double)
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


'//cIPersist Interface
Private Property Get cIPersist_Classname() As String
End Property
Private Property Get cIPersist_GUID() As String
End Property

Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As GVD_XML_TYPE)
    Dim i As Long
    
    If iMode = cmp Then
        m_sDescription = op.Load(XML_NODE_DESCRIPTION)
        m_lngChildCount = op.Load(XML_NODE_CHILDCOUNT)
        
        If m_lngChildCount > 0 Then
            ReDim m_oChildren(m_lngChildCount - 1)
            For i = 0 To m_lngChildCount - 1
                Set m_oChildren(i) = op.Load(XML_NODE_CHILD & i)
            Next
        End If
    Else
        m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
        m_sImage = op.Load(XML_NODE_IMAGE)
        m_lngPropCount = op.Load(XML_NODE_PROPERTYCOUNT)
        
        If m_lngPropCount > 0 Then
            ReDim m_oProperties(m_lngPropCount - 1)
            For i = 0 To m_lngPropCount - 1
                Set m_oProperties(i) = op.Load(XML_NODE_PROPERTY & i)
            Next
        End If
    End If
    'todo: check that child/member objects are "not" Is  Nothing after "load" attempt? or some other way?
End Sub
Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
End Sub
'//
Private Sub Class_Initialize()

    'LoadMatrices 'load all of our component matrix data 'todo: obsolete this
    Set Veh = Me ' public reference to this class
    'Set Components = New Collection    'todo: obsolete this
    
    'todo: i shouldnt be Setting any of these here!  Instead, I should only check after loading from file that all
    ' reqt objects are set
    'Set Crew = New clsCrew
    'Set Surface = New clsSurface
    'Set Options = New clsOptions
    'Set Description = New clsDescription
    'Set Stats = New clsStats
    
    ' holds collection of link objects
    'Set WeaponProfiles = New clsProfile
    'Set Profiles = New clsProfile
    'Set BatteryProfiles = New clsProfile
   
    ' holds performance class objects
    'Set PerformanceProfiles = New clsProfile
    
End Sub

Private Sub Class_Terminate()
    
    'todo: loop Set m_oProperties(0) = Nothing
    
    
    'Set Body = Nothing
   
    'Set Crew = Nothing
    'Set Surface = Nothing
    'Set Options = Nothing
    'Set Description = Nothing
    'Set Stats = Nothing
    '--

    '--
    'Set WeaponProfiles = Nothing
    'Set Profiles = Nothing
    'Set BatteryProfiles = Nothing
    'Set PerformanceProfiles = Nothing
    Dim i As Long
    For i = 0 To m_lngChildCount - 1
        Set m_oChildren(i) = Nothing
    Next
    
    For i = 0 To m_lngPropCount - 1
        Set m_oProperties(i) = Nothing
    Next
End Sub
