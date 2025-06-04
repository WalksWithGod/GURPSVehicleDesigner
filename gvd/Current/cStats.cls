VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStats"
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
End Property
Private Property Get cINode_Description() As String
End Property
Private Property Let cINode_Image(ByVal RHS As String)
    m_sImage = RHS
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

        Debug.Assert m_lngChildCount = 0
    Else 'DEF
        
        m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
        Debug.Assert m_lngMaxChildren = 0
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


' just remember this, using objects for this crap will let me use shell interfaces for each component
'yet handle all the bulk chores with one object with one set of code instead of typing the same
' code handling in every single class since i cant inherit that code with VB.

'

Public Function calcStats(ByRef oComponent As cIComponent) As Boolean

    #If DEBUG_MODE Then
        Dim oNode As cINode
        
        Set oNode = oComponent
        InfoPrint 1, "calculating stats for " & oNode.Description
        Set oNode = Nothing
    #End If
End Function
