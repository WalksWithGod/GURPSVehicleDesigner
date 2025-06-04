VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGroupCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements cIPersist
Implements cINode

' -------cINode interface variables
Private Const MAX_CHILDREN = 32
Private m_lngMaxChildren As Long
Private m_lngChildCount As Long
Private m_oChildren() As cINode
Private m_lngAttributes As Long
Private m_hParent As Long
Private m_hMe As Long
Private m_sName As String
Private m_sDescription As String
Private m_sImage As String
'---------

Private Sub Class_Initialize()
vbwProfiler.vbwProcIn 532
vbwProfiler.vbwExecuteLine 10953
    m_lngChildCount = 0
vbwProfiler.vbwExecuteLine 10954
    m_lngMaxChildren = MAX_CHILDREN
vbwProfiler.vbwProcOut 532
vbwProfiler.vbwExecuteLine 10955
End Sub
Private Sub Class_Terminate()
vbwProfiler.vbwProcIn 533
    Dim i As Long
vbwProfiler.vbwExecuteLine 10956
    For i = 0 To m_lngChildCount - 1
vbwProfiler.vbwExecuteLine 10957
        Set m_oChildren(i) = Nothing
vbwProfiler.vbwExecuteLine 10958
    Next
vbwProfiler.vbwProcOut 533
vbwProfiler.vbwExecuteLine 10959
End Sub
'///////////////////////////////////////////////////

'///////////////////////////////////////////////////
Private Property Get cIPersist_Classname() As String
vbwProfiler.vbwProcIn 534
vbwProfiler.vbwProcOut 534
vbwProfiler.vbwExecuteLine 10960
End Property
Private Property Get cIPersist_GUID() As String
vbwProfiler.vbwProcIn 535
vbwProfiler.vbwProcOut 535
vbwProfiler.vbwExecuteLine 10961
End Property
Private Sub cIPersist_LoadProperties(ByVal op As PersistenceManager.clsObjProperties, ByVal iMode As Long)
vbwProfiler.vbwProcIn 536
    Dim i As Long
vbwProfiler.vbwExecuteLine 10962
    m_lngChildCount = op.Load("childcount")
vbwProfiler.vbwExecuteLine 10963
    m_sDescription = op.Load("description")

vbwProfiler.vbwExecuteLine 10964
    Debug.Assert m_lngChildCount <= MAX_CHILDREN
vbwProfiler.vbwExecuteLine 10965
    If m_lngChildCount > 0 Then
vbwProfiler.vbwExecuteLine 10966
        ReDim m_oChildren(m_lngChildCount - 1)
vbwProfiler.vbwExecuteLine 10967
        For i = 0 To m_lngChildCount - 1
vbwProfiler.vbwExecuteLine 10968
            Set m_oChildren(i) = op.Load("child" & i)
vbwProfiler.vbwExecuteLine 10969
        Next
    End If
vbwProfiler.vbwExecuteLine 10970 'B
vbwProfiler.vbwProcOut 536
vbwProfiler.vbwExecuteLine 10971
End Sub

Private Sub cIPersist_StoreProperties(ByVal op As PersistenceManager.clsObjProperties)
vbwProfiler.vbwProcIn 537
    Dim i As Long

vbwProfiler.vbwExecuteLine 10972
    op.Store "classname", TypeName(Me)
vbwProfiler.vbwExecuteLine 10973
    op.Store "defpath", ""
vbwProfiler.vbwExecuteLine 10974
    op.Store "guid", ""

vbwProfiler.vbwExecuteLine 10975
    op.Store "description", m_sDescription
vbwProfiler.vbwExecuteLine 10976
    op.Store "childcount", m_lngChildCount

vbwProfiler.vbwExecuteLine 10977
    For i = 0 To m_lngChildCount - 1
vbwProfiler.vbwExecuteLine 10978
        op.Store "child" & i, m_oChildren(i)
vbwProfiler.vbwExecuteLine 10979
    Next
vbwProfiler.vbwProcOut 537
vbwProfiler.vbwExecuteLine 10980
End Sub
'///////////////////////////////////////////////////
'//cINode Implemented Properties and Functions
Private Function cINode_AddChild(oNode As cINode) As Boolean
vbwProfiler.vbwProcIn 538
vbwProfiler.vbwExecuteLine 10981
    If m_lngMaxChildren = m_lngChildCount Then
vbwProfiler.vbwExecuteLine 10982
        cINode_AddChild = False
    Else
vbwProfiler.vbwExecuteLine 10983 'B
vbwProfiler.vbwExecuteLine 10984
        m_lngChildCount = m_lngChildCount + 1
vbwProfiler.vbwExecuteLine 10985
        ReDim Preserve m_oChildren(0 To m_lngChildCount - 1)
vbwProfiler.vbwExecuteLine 10986
        Set m_oChildren(m_lngChildCount - 1) = oNode
vbwProfiler.vbwExecuteLine 10987
        cINode_AddChild = True
    End If
vbwProfiler.vbwExecuteLine 10988 'B
vbwProfiler.vbwProcOut 538
vbwProfiler.vbwExecuteLine 10989
End Function
Private Function cINode_getChildrenByClassName(ByRef Classname As String, hChilds() As Long) As Boolean
vbwProfiler.vbwProcIn 539
vbwProfiler.vbwProcOut 539
vbwProfiler.vbwExecuteLine 10990
End Function
Private Function cINode_getChildIndexByHandle(ByVal h As Long) As Long
vbwProfiler.vbwProcIn 540
    Dim i As Long
    Dim lRet As Long
vbwProfiler.vbwExecuteLine 10991
    lRet = -1
vbwProfiler.vbwExecuteLine 10992
    For i = 0 To m_lngChildCount - 1
vbwProfiler.vbwExecuteLine 10993
        If m_oChildren(i).Handle = h Then
vbwProfiler.vbwExecuteLine 10994
             lRet = i
vbwProfiler.vbwExecuteLine 10995
             Exit For
        End If
vbwProfiler.vbwExecuteLine 10996 'B
vbwProfiler.vbwExecuteLine 10997
    Next
vbwProfiler.vbwExecuteLine 10998
    cINode_getChildIndexByHandle = lRet
vbwProfiler.vbwProcOut 540
vbwProfiler.vbwExecuteLine 10999
End Function
Private Function cINode_getChild(ByVal lngIndex As Long) As cINode
vbwProfiler.vbwProcIn 541
vbwProfiler.vbwExecuteLine 11000
    If (lngIndex >= 0) And (m_lngChildCount > 0) And (lngIndex <= m_lngChildCount - 1) Then
vbwProfiler.vbwExecuteLine 11001
        Set cINode_getChild = m_oChildren(lngIndex)
    End If
vbwProfiler.vbwExecuteLine 11002 'B
vbwProfiler.vbwProcOut 541
vbwProfiler.vbwExecuteLine 11003
End Function
Private Property Get cINode_childCount() As Long
vbwProfiler.vbwProcIn 542
vbwProfiler.vbwExecuteLine 11004
    cINode_childCount = m_lngChildCount
vbwProfiler.vbwProcOut 542
vbwProfiler.vbwExecuteLine 11005
End Property
Private Function cINode_removeChild(ByVal lngIndex As Long) As Boolean
vbwProfiler.vbwProcIn 543
    Dim i As Long
vbwProfiler.vbwExecuteLine 11006
    If (lngIndex <= m_lngChildCount - 1) And (lngIndex >= 0) Then
vbwProfiler.vbwExecuteLine 11007
        Set m_oChildren(lngIndex) = Nothing
vbwProfiler.vbwExecuteLine 11008
        For i = lngIndex + 1 To m_lngChildCount - 1
vbwProfiler.vbwExecuteLine 11009
            Set m_oChildren(i - 1) = m_oChildren(i)
vbwProfiler.vbwExecuteLine 11010
        Next
    End If
vbwProfiler.vbwExecuteLine 11011 'B
vbwProfiler.vbwExecuteLine 11012
    m_lngChildCount = m_lngChildCount - 1

vbwProfiler.vbwExecuteLine 11013
    If m_lngChildCount > 0 Then
vbwProfiler.vbwExecuteLine 11014
        ReDim Preserve m_oChildren(0 To m_lngChildCount - 1)
    Else
vbwProfiler.vbwExecuteLine 11015 'B
vbwProfiler.vbwExecuteLine 11016
        ReDim m_oChildren(0)
    End If
vbwProfiler.vbwExecuteLine 11017 'B
vbwProfiler.vbwProcOut 543
vbwProfiler.vbwExecuteLine 11018
End Function
Private Property Get cINode_ClassName() As String
vbwProfiler.vbwProcIn 544
vbwProfiler.vbwExecuteLine 11019
    cINode_ClassName = TypeName(Me)
vbwProfiler.vbwProcOut 544
vbwProfiler.vbwExecuteLine 11020
End Property
Private Property Get cINode_Attributes() As Long
vbwProfiler.vbwProcIn 545
vbwProfiler.vbwExecuteLine 11021
    cINode_Attributes = m_lngAttributes
vbwProfiler.vbwProcOut 545
vbwProfiler.vbwExecuteLine 11022
End Property
Private Property Let cINode_Handle(ByVal l As Long)
vbwProfiler.vbwProcIn 546
vbwProfiler.vbwExecuteLine 11023
    m_hMe = l
vbwProfiler.vbwProcOut 546
vbwProfiler.vbwExecuteLine 11024
End Property
Private Property Get cINode_Handle() As Long
vbwProfiler.vbwProcIn 547
vbwProfiler.vbwExecuteLine 11025
    cINode_Handle = m_hMe
vbwProfiler.vbwProcOut 547
vbwProfiler.vbwExecuteLine 11026
End Property
Private Property Let cINode_Parent(ByVal l As Long)
vbwProfiler.vbwProcIn 548
vbwProfiler.vbwProcOut 548
vbwProfiler.vbwExecuteLine 11027
End Property
Private Property Get cINode_Parent() As Long
vbwProfiler.vbwProcIn 549
vbwProfiler.vbwProcOut 549
vbwProfiler.vbwExecuteLine 11028
End Property
Private Property Let cINode_Name(ByVal s As String)
vbwProfiler.vbwProcIn 550
vbwProfiler.vbwExecuteLine 11029
    m_sName = s
vbwProfiler.vbwProcOut 550
vbwProfiler.vbwExecuteLine 11030
End Property
Private Property Get cINode_Name() As String
vbwProfiler.vbwProcIn 551
vbwProfiler.vbwExecuteLine 11031
    cINode_Name = m_sName
vbwProfiler.vbwProcOut 551
vbwProfiler.vbwExecuteLine 11032
End Property
Private Property Let cINode_Description(ByVal s As String)
vbwProfiler.vbwProcIn 552
vbwProfiler.vbwExecuteLine 11033
    m_sDescription = s
vbwProfiler.vbwProcOut 552
vbwProfiler.vbwExecuteLine 11034
End Property
Private Property Get cINode_Description() As String
vbwProfiler.vbwProcIn 553
vbwProfiler.vbwExecuteLine 11035
    cINode_Description = m_sDescription
vbwProfiler.vbwProcOut 553
vbwProfiler.vbwExecuteLine 11036
End Property
Private Property Let cINode_Image(ByVal s As String)
vbwProfiler.vbwProcIn 554
vbwProfiler.vbwExecuteLine 11037
    m_sImage = s
vbwProfiler.vbwProcOut 554
vbwProfiler.vbwExecuteLine 11038
End Property
Private Property Get cINode_Image() As String
vbwProfiler.vbwProcIn 555
vbwProfiler.vbwExecuteLine 11039
    cINode_Image = m_sImage
vbwProfiler.vbwProcOut 555
vbwProfiler.vbwExecuteLine 11040
End Property



