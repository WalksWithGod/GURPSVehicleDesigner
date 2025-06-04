VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements cIDisplay
Implements cINode
Implements cIPersist

' -------cINode interface variables
Private m_lngMaxChildren As Long
Private m_lngChildCount As Long
Private m_oChildren() As cINode  ' stats,crew, surface,hull and components classes need to go here.  all implement cInode and are thus autographed.  Only "components" implements cIContainer and can thus contain other 'components'.
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

' ------- cIomponent interface properties
Private m_Table() As Single
Private m_lngTL As Long
Private m_dblHitpoints As Double
Private m_dblSurfaceArea As Double
Private m_dblCost As Double
Private m_dblVolume As Double
Private m_dblWeight As Double
Public m_oArmor As cArmorLayer  ' if it includes armor here,it should implement cIComponent since that includes armor as a public composite object
'todo: make sure this is defined in the cIComponent interface

' ------- cVehicle required variables

' note that these are not declared as cINode since they do not have a seperte node in the tree.  instead they are simply
' referenced through the cVehicle's own node and display interfaces
Public Author As cAuthor
Public Version As cVersion
Public Description As cDescription


' Since the following will display in the treex, they must all implement cINode and cIDisplay
' Since they all must be persistable, they will implement cIPersist
' However, none of them need to implement cICompoent or cIContainer.  Perhaps they should also
' all be stored in the children() array.  this provides auto graphing.  We'll only need to do a quick
' search of first level children in cVehicles to find any particular one.

'Public Stats As cINode 'cStats
'Public Crew As cINode 'cCrew
'Public Surface As cINode 'cSurface
'Public Hull As cINode ' cHull


'Public PowerProfiles As cINode
'Public FuelProfiles As cINode
'Public WeaponProfiles As cINode
'Public PerformanceProfiles As cINode
'Public BatteryProfiles As cINode
'Public Profiles As clsProfile

'//
Private Sub Class_Initialize()

    ' Public Veh as cVehicle is declared in modHelper as global.  Thats definetly bad.
    Set Veh = Me ' public reference to this class 'todo: This is bad
    ' We do need a way for objects within a vehicle to refer to the root
    ' BUT, if we have arrays of vehicles, they will each overwrite Veh with a reference to themselves.
    ' We need to manage all vehicle references thru cManager and use a collection.
    
    '-----------------all below obsolete?
   
    ' holds collection of link objects
    'Set WeaponProfiles = New clsProfile
    'Set Profiles = New clsProfile
    'Set BatteryProfiles = New clsProfile
   
    ' holds performance class objects
    'Set PerformanceProfiles = New clsProfile
    '------------------------------------------
    
End Sub

Private Sub Class_Terminate()
    '-----------------obsolete?
    ' i believe so.  All of these items implement cInode and
    '  are added via XML script just like every other node in the tree
    'Set WeaponProfiles = Nothing
    'Set Profiles = Nothing
    'Set BatteryProfiles = Nothing
    'Set PerformanceProfiles = Nothing
    
    '---------------------------------
    Dim I As Long
    
    Set Author = Nothing
    Set Description = Nothing
    Set Version = Nothing
    
    For I = 0 To m_lngChildCount - 1
        Set m_oChildren(I) = Nothing
    Next
    For I = 0 To m_lngPropCount - 1
        Set m_oProperties(I) = Nothing
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
'/////////////////////////////////////////////
'//Implemented cINode Properties and Functions
Private Function cINode_AddChild(oNode As cINode) As Boolean
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
    Dim I As Long
    Dim lRet As Long
    lRet = -1
    For I = 0 To m_lngChildCount - 1
        If m_oChildren(I).Handle = h Then lRet = I: Exit For
    Next
    cINode_getChildIndexByHandle = lRet
End Function
Private Function cINode_getChild(ByVal lngIndex As Long) As cINode
    If (lngIndex >= 0) And (m_lngChildCount > 0) And (lngIndex <= m_lngChildCount - 1) Then
        Set cINode_getChild = m_oChildren(lngIndex)
    End If
End Function
Private Function cINode_removeChild(ByVal lngIndex As Long) As Boolean
    Dim I As Long
    If (lngIndex <= m_lngChildCount - 1) And (lngIndex >= 0) Then
        Set m_oChildren(lngIndex) = Nothing
        For I = lngIndex + 1 To m_lngChildCount - 1
            Set m_oChildren(I - 1) = m_oChildren(I)
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
End Property
Private Property Get cIPersist_GUID() As String
End Property

Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Long)
    Dim I As Long
    
    If iMode = cmp Then
        m_sName = op.Load(XML_NODE_NAME)
        m_sDescription = op.Load(XML_NODE_DESCRIPTION)
        m_lngChildCount = op.Load(XML_NODE_CHILDCOUNT)
        
        If m_lngChildCount > 0 Then
            ReDim m_oChildren(m_lngChildCount - 1)
            For I = 0 To m_lngChildCount - 1
                Set m_oChildren(I) = op.Load(XML_NODE_CHILD & I)
                m_oChildren(I).Parent = m_hMe
                Debug.Assert m_oChildren(I).Parent <> 0
            Next
        End If
        
        Set Author = op.Load(XML_NODE_AUTHOR)
        Set Version = op.Load("version")
        Set Description = op.Load("description")
    Else
        m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
        m_sImage = op.Load(XML_NODE_IMAGE)
        m_lngPropCount = op.Load(XML_NODE_PROPERTYCOUNT)
        
        If m_lngPropCount > 0 Then
            ReDim m_oProperties(m_lngPropCount - 1)
            For I = 0 To m_lngPropCount - 1
                Set m_oProperties(I) = op.Load(XML_NODE_PROPERTY & I)
            Next
        End If
    End If
    'todo: check that child/member objects are "not" Is  Nothing after "load" attempt? or some other way?
End Sub
Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
End Sub
