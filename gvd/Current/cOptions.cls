VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOptions"
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



'todo: this class is probably obsolete since cFeatures with cSurface covers them
' and ram and bulldozers, etc are actual components now that are dragged onto the vehicle

' On/Off settings
Private mvarRam As Boolean
Private mvarBulldozer As Boolean
Private mvarPlow As Boolean
Private mvarHitch As Boolean
Private mvarPin As String
Private mvarConvertible As String

' Weight/Cost for components which really should be things that you can ADD independantly
' rather than just jumpbling them all up as "options"
Private mvarRamWeight As Single
Private mvarRamCost As Single
Private mvarBullDozerWeight As Single
Private mvarBullDozerCost As Single
Private mvarPlowWeight As Single
Private mvarPlowCost As Single
Private mvarHitchCost As Single
Private mvarHitchWeight As Single
Private mvarPinCost As Single
Private mvarPinWeight As Single
Private mvarConvertibleCost As Single  '<-- why is this a "surface option" and topdeck isnt?  Inconsistancy in Vehicles if you ask me
Private mvarConvertibleWeight As Single




Private Sub Class_Initialize()
    mvarRam = False
    mvarBulldozer = False
    mvarPlow = False
    mvarHitch = False
    mvarPin = "none"
    mvarConvertible = "none"

End Sub

Private Sub Class_Terminate()
Dim i As Long
    For i = 0 To m_lngChildCount - 1
        Set m_oChildren(i) = Nothing
    Next
    For i = 0 To m_lngPropCount - 1
        Set m_oProperties(i) = Nothing
    Next
End Sub


'========================

Public Property Let Pin(ByVal vdata As String)
    mvarPin = vdata
End Property
Public Property Get Pin() As String
    Pin = mvarPin
End Property
Public Property Let Hitch(ByVal vdata As Boolean)
    mvarHitch = vdata
End Property
Public Property Get Hitch() As Boolean
    Hitch = mvarHitch
End Property
Public Property Let Plow(ByVal vdata As Boolean)
    mvarPlow = vdata
End Property
Public Property Get Plow() As Boolean
    Plow = mvarPlow
End Property
Public Property Let Bulldozer(ByVal vdata As Boolean)
    mvarBulldozer = vdata
End Property
Public Property Get Bulldozer() As Boolean
    Bulldozer = mvarBulldozer
End Property
Public Property Let Ram(ByVal vdata As Boolean)
    mvarRam = vdata
End Property
Public Property Get Ram() As Boolean
    Ram = mvarRam
End Property
Public Property Let RamWeight(ByVal vdata As Double)
    mvarRamWeight = vdata
End Property
Public Property Get RamWeight() As Double
    RamWeight = mvarRamWeight
End Property
Public Property Let RamCost(ByVal vdata As Double)
    mvarRamCost = vdata
End Property
Public Property Get RamCost() As Double
    RamCost = mvarRamCost
End Property
Public Property Let BullDozerWeight(ByVal vdata As Double)
    mvarBullDozerWeight = vdata
End Property
Public Property Get BullDozerWeight() As Double
    BullDozerWeight = mvarBullDozerWeight
End Property
Public Property Let BullDozerCost(ByVal vdata As Double)
    mvarBullDozerCost = vdata
End Property
Public Property Get BullDozerCost() As Double
    BullDozerCost = mvarBullDozerCost
End Property
Public Property Let PlowWeight(ByVal vdata As Double)
    mvarPlowWeight = vdata
End Property
Public Property Get PlowWeight() As Double
    PlowWeight = mvarPlowWeight
End Property
Public Property Let PlowCost(ByVal vdata As Double)
    mvarPlowCost = vdata
End Property
Public Property Get PlowCost() As Double
    PlowCost = mvarPlowCost
End Property
Public Property Let HitchCost(ByVal vdata As Double)
    mvarHitchCost = vdata
End Property
Public Property Get HitchCost() As Double
    HitchCost = mvarHitchCost
End Property
Public Property Let HitchWeight(ByVal vdata As Double)
    mvarHitchWeight = vdata
End Property
Public Property Get HitchWeight() As Double
    HitchWeight = mvarHitchWeight
End Property
Public Property Let PinCost(ByVal vdata As Double)
    mvarPinCost = vdata
End Property
Public Property Get PinCost() As Double
    PinCost = mvarPinCost
End Property
Public Property Let PinWeight(ByVal vdata As Double)
    mvarPinWeight = vdata
End Property
Public Property Get PinWeight() As Double
    PinWeight = mvarPinWeight
End Property
Public Property Let ConvertibleCost(ByVal vdata As Double)
    mvarConvertibleCost = vdata
End Property
Public Property Get ConvertibleCost() As Double
    ConvertibleCost = mvarConvertibleCost
End Property
Public Property Let ConvertibleWeight(ByVal vdata As Double)
    mvarConvertibleWeight = vdata
End Property
Public Property Get ConvertibleWeight() As Double
    ConvertibleWeight = mvarConvertibleWeight
End Property

Public Property Get OptionsCost() As Double
    OptionsCost = mvarOptionsCost
End Property
Public Property Let OptionsCost(ByVal vdata As Double)
    mvarOptionsCost = vdata
End Property
Public Property Let OptionsWeight(ByVal vdata As Double)
    mvarOptionsWeight = vdata
End Property
Public Property Get OptionsWeight() As Double
    OptionsWeight = mvarOptionsWeight
End Property

Public Property Let Convertible(ByVal vdata As String)
    If vdata = "hardtop" Then
        InfoPrint 1, "Reminder: Only vehicles with Six or Fewer seats, excluding Exposed seats, may have a detachable hardtop.  The Vehicle may not have top mounted Superstructures or Turrets."
    ElseIf vdata = "ragtop" Then
        MsgBox "Reminder: Only Vehicles with Exposed seats may have a fabric top"
    End If
'TODO: Move this check to the Design Check
    mvarConvertible = vdata
End Property
Public Property Get Convertible() As String
    Convertible = mvarConvertible
End Property



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
'TODO: will options ever accept children?  Probably not... just remember
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
    If iMode = cmp Then
        m_sName = op.Load(XML_NODE_NAME)
        m_sDescription = op.Load(XML_NODE_DESCRIPTION)
    Else 'DEF
        m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
        m_sImage = op.Load(XML_NODE_IMAGE)
    End If
End Sub
Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
End Sub

