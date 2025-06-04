VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRuleGroup"
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
'----------

Private m_lngStatType As Long ' these types will be loaded from file

Private Sub Class_Initialize()
vbwProfiler.vbwProcIn 603
vbwProfiler.vbwExecuteLine 11207
    m_lngChildCount = 0
vbwProfiler.vbwExecuteLine 11208
    m_lngMaxChildren = MAX_CHILDREN
vbwProfiler.vbwProcOut 603
vbwProfiler.vbwExecuteLine 11209
End Sub
Private Sub Class_Terminate()
vbwProfiler.vbwProcIn 604
    Dim i As Long
vbwProfiler.vbwExecuteLine 11210
    For i = 0 To m_lngChildCount - 1
vbwProfiler.vbwExecuteLine 11211
        Set m_oChildren(i) = Nothing
vbwProfiler.vbwExecuteLine 11212
    Next
vbwProfiler.vbwProcOut 604
vbwProfiler.vbwExecuteLine 11213
End Sub
Public Property Get statType() As Long
vbwProfiler.vbwProcIn 605
vbwProfiler.vbwExecuteLine 11214
    statType = m_lngStatType
vbwProfiler.vbwProcOut 605
vbwProfiler.vbwExecuteLine 11215
End Property
Public Property Let statType(ByVal i As Long)
vbwProfiler.vbwProcIn 606
vbwProfiler.vbwExecuteLine 11216
    m_lngStatType = i
vbwProfiler.vbwProcOut 606
vbwProfiler.vbwExecuteLine 11217
End Property

' returns true if the value was run and alters the value of v
' note: We should return a string expression which is the completely formatted value
' including thousand seperators and postfixes.
Function runRule(ByVal dblInput As Double, ByRef sResult As String) As Boolean
vbwProfiler.vbwProcIn 607
    Dim oRule As cRule
    Dim oNode As cINode
    Dim oConvert As cUnitConverter
    Dim i As Long
    Dim bRet As Boolean

vbwProfiler.vbwExecuteLine 11218
    For i = 0 To m_lngChildCount - 1
vbwProfiler.vbwExecuteLine 11219
        Set oRule = m_oChildren(i)

vbwProfiler.vbwExecuteLine 11220
        If oRule.testRule(dblInput) Then
vbwProfiler.vbwExecuteLine 11221
            bRet = oConvert.convertValue(oRule.convertFrom, oRule.convertTo, dblInput)
vbwProfiler.vbwExecuteLine 11222
            runRule = bRet
vbwProfiler.vbwExecuteLine 11223
            Exit For
        End If
vbwProfiler.vbwExecuteLine 11224 'B
vbwProfiler.vbwExecuteLine 11225
    Next

vbwProfiler.vbwExecuteLine 11226
    Set oRule = Nothing
vbwProfiler.vbwExecuteLine 11227
    Set oNode = Nothing
vbwProfiler.vbwExecuteLine 11228
    Set oConvert = Nothing
vbwProfiler.vbwProcOut 607
vbwProfiler.vbwExecuteLine 11229
End Function


'///////////////////////////////////////////////////
'//cIPersist Implemented Properties and Functions
Private Property Get cIPersist_Classname() As String
vbwProfiler.vbwProcIn 608
vbwProfiler.vbwProcOut 608
vbwProfiler.vbwExecuteLine 11230
End Property
Private Property Get cIPersist_GUID() As String
vbwProfiler.vbwProcIn 609
vbwProfiler.vbwProcOut 609
vbwProfiler.vbwExecuteLine 11231
End Property
Private Sub cIPersist_LoadProperties(ByVal op As PersistenceManager.clsObjProperties, ByVal iMode As Long)
vbwProfiler.vbwProcIn 610
    Dim i As Long

vbwProfiler.vbwExecuteLine 11232
    m_lngChildCount = op.Load("childcount")
vbwProfiler.vbwExecuteLine 11233
    m_sName = op.Load("name")
vbwProfiler.vbwExecuteLine 11234
    m_sDescription = op.Load("description")

vbwProfiler.vbwExecuteLine 11235
    Debug.Assert m_lngChildCount <= MAX_CHILDREN
vbwProfiler.vbwExecuteLine 11236
    If m_lngChildCount > 0 Then
vbwProfiler.vbwExecuteLine 11237
        ReDim m_oChildren(m_lngChildCount - 1)
vbwProfiler.vbwExecuteLine 11238
        For i = 0 To m_lngChildCount - 1
vbwProfiler.vbwExecuteLine 11239
            Set m_oChildren(i) = op.Load("child" & i)
vbwProfiler.vbwExecuteLine 11240
        Next
    End If
vbwProfiler.vbwExecuteLine 11241 'B
vbwProfiler.vbwProcOut 610
vbwProfiler.vbwExecuteLine 11242
End Sub
Private Sub cIPersist_StoreProperties(ByVal op As PersistenceManager.clsObjProperties)
vbwProfiler.vbwProcIn 611
    Dim i As Long

vbwProfiler.vbwExecuteLine 11243
    op.Store "classname", TypeName(Me)
vbwProfiler.vbwExecuteLine 11244
    op.Store "defpath", ""
vbwProfiler.vbwExecuteLine 11245
    op.Store "guid", ""

vbwProfiler.vbwExecuteLine 11246
    op.Store "name", m_sName
vbwProfiler.vbwExecuteLine 11247
    op.Store "description", m_sDescription
vbwProfiler.vbwExecuteLine 11248
    op.Store "childcount", m_lngChildCount

vbwProfiler.vbwExecuteLine 11249
    For i = 0 To m_lngChildCount - 1
vbwProfiler.vbwExecuteLine 11250
        op.Store "child" & i, m_oChildren(i)
vbwProfiler.vbwExecuteLine 11251
    Next

vbwProfiler.vbwProcOut 611
vbwProfiler.vbwExecuteLine 11252
End Sub
'///////////////////////////////////////////////////
'//cINode Implemented Properties and Functions
Private Function cINode_AddChild(oNode As cINode) As Boolean
vbwProfiler.vbwProcIn 612
vbwProfiler.vbwExecuteLine 11253
    If m_lngMaxChildren = m_lngChildCount Then
vbwProfiler.vbwExecuteLine 11254
        cINode_AddChild = False
    Else
vbwProfiler.vbwExecuteLine 11255 'B
vbwProfiler.vbwExecuteLine 11256
        m_lngChildCount = m_lngChildCount + 1
vbwProfiler.vbwExecuteLine 11257
        ReDim Preserve m_oChildren(0 To m_lngChildCount - 1)
vbwProfiler.vbwExecuteLine 11258
        Set m_oChildren(m_lngChildCount - 1) = oNode
vbwProfiler.vbwExecuteLine 11259
        cINode_AddChild = True
    End If
vbwProfiler.vbwExecuteLine 11260 'B
vbwProfiler.vbwProcOut 612
vbwProfiler.vbwExecuteLine 11261
End Function
Private Function cINode_getChildrenByClassName(ByRef Classname As String, hChilds() As Long) As Boolean
vbwProfiler.vbwProcIn 613
vbwProfiler.vbwProcOut 613
vbwProfiler.vbwExecuteLine 11262
End Function
Private Function cINode_getChildIndexByHandle(ByVal h As Long) As Long
vbwProfiler.vbwProcIn 614
    Dim i As Long
    Dim lRet As Long
vbwProfiler.vbwExecuteLine 11263
    lRet = -1
vbwProfiler.vbwExecuteLine 11264
    For i = 0 To m_lngChildCount - 1
vbwProfiler.vbwExecuteLine 11265
        If m_oChildren(i).Handle = h Then
vbwProfiler.vbwExecuteLine 11266
             lRet = i
vbwProfiler.vbwExecuteLine 11267
             Exit For
        End If
vbwProfiler.vbwExecuteLine 11268 'B
vbwProfiler.vbwExecuteLine 11269
    Next
vbwProfiler.vbwExecuteLine 11270
    cINode_getChildIndexByHandle = lRet
vbwProfiler.vbwProcOut 614
vbwProfiler.vbwExecuteLine 11271
End Function
Private Function cINode_getChild(ByVal lngIndex As Long) As cINode
vbwProfiler.vbwProcIn 615
vbwProfiler.vbwExecuteLine 11272
    If (lngIndex >= 0) And (m_lngChildCount > 0) And (lngIndex <= m_lngChildCount - 1) Then
vbwProfiler.vbwExecuteLine 11273
        Set cINode_getChild = m_oChildren(lngIndex)
    End If
vbwProfiler.vbwExecuteLine 11274 'B
vbwProfiler.vbwProcOut 615
vbwProfiler.vbwExecuteLine 11275
End Function
Private Function cINode_removeChild(ByVal lngIndex As Long) As Boolean
vbwProfiler.vbwProcIn 616
    Dim i As Long
vbwProfiler.vbwExecuteLine 11276
    If (lngIndex <= m_lngChildCount - 1) And (lngIndex >= 0) Then
vbwProfiler.vbwExecuteLine 11277
        Set m_oChildren(lngIndex) = Nothing
vbwProfiler.vbwExecuteLine 11278
        For i = lngIndex + 1 To m_lngChildCount - 1
vbwProfiler.vbwExecuteLine 11279
            Set m_oChildren(i - 1) = m_oChildren(i)
vbwProfiler.vbwExecuteLine 11280
        Next
    End If
vbwProfiler.vbwExecuteLine 11281 'B
vbwProfiler.vbwExecuteLine 11282
    m_lngChildCount = m_lngChildCount - 1

vbwProfiler.vbwExecuteLine 11283
    If m_lngChildCount > 0 Then
vbwProfiler.vbwExecuteLine 11284
        ReDim Preserve m_oChildren(0 To m_lngChildCount - 1)
    Else
vbwProfiler.vbwExecuteLine 11285 'B
vbwProfiler.vbwExecuteLine 11286
        ReDim m_oChildren(0)
    End If
vbwProfiler.vbwExecuteLine 11287 'B
vbwProfiler.vbwProcOut 616
vbwProfiler.vbwExecuteLine 11288
End Function
Private Property Get cINode_childCount() As Long
vbwProfiler.vbwProcIn 617
vbwProfiler.vbwExecuteLine 11289
    cINode_childCount = m_lngChildCount
vbwProfiler.vbwProcOut 617
vbwProfiler.vbwExecuteLine 11290
End Property
Private Property Get cINode_ClassName() As String
vbwProfiler.vbwProcIn 618
vbwProfiler.vbwExecuteLine 11291
    cINode_ClassName = TypeName(Me)
vbwProfiler.vbwProcOut 618
vbwProfiler.vbwExecuteLine 11292
End Property
Private Property Get cINode_Attributes() As Long
vbwProfiler.vbwProcIn 619
vbwProfiler.vbwExecuteLine 11293
    cINode_Attributes = m_lngAttributes
vbwProfiler.vbwProcOut 619
vbwProfiler.vbwExecuteLine 11294
End Property
Private Property Let cINode_Handle(ByVal l As Long)
vbwProfiler.vbwProcIn 620
vbwProfiler.vbwExecuteLine 11295
    m_hMe = l
vbwProfiler.vbwProcOut 620
vbwProfiler.vbwExecuteLine 11296
End Property
Private Property Get cINode_Handle() As Long
vbwProfiler.vbwProcIn 621
vbwProfiler.vbwExecuteLine 11297
    cINode_Handle = m_hMe
vbwProfiler.vbwProcOut 621
vbwProfiler.vbwExecuteLine 11298
End Property
Private Property Let cINode_Parent(ByVal l As Long)
vbwProfiler.vbwProcIn 622
vbwProfiler.vbwProcOut 622
vbwProfiler.vbwExecuteLine 11299
End Property
Private Property Get cINode_Parent() As Long
vbwProfiler.vbwProcIn 623
vbwProfiler.vbwProcOut 623
vbwProfiler.vbwExecuteLine 11300
End Property
Private Property Let cINode_Name(ByVal s As String)
vbwProfiler.vbwProcIn 624
vbwProfiler.vbwExecuteLine 11301
    m_sName = s
vbwProfiler.vbwProcOut 624
vbwProfiler.vbwExecuteLine 11302
End Property
Private Property Get cINode_Name() As String
vbwProfiler.vbwProcIn 625
vbwProfiler.vbwExecuteLine 11303
    cINode_Name = m_sName
vbwProfiler.vbwProcOut 625
vbwProfiler.vbwExecuteLine 11304
End Property
Private Property Let cINode_Description(ByVal s As String)
vbwProfiler.vbwProcIn 626
vbwProfiler.vbwExecuteLine 11305
    m_sDescription = s
vbwProfiler.vbwProcOut 626
vbwProfiler.vbwExecuteLine 11306
End Property
Private Property Get cINode_Description() As String
vbwProfiler.vbwProcIn 627
vbwProfiler.vbwExecuteLine 11307
    cINode_Description = m_sDescription
vbwProfiler.vbwProcOut 627
vbwProfiler.vbwExecuteLine 11308
End Property
Private Property Let cINode_Image(ByVal s As String)
vbwProfiler.vbwProcIn 628
vbwProfiler.vbwExecuteLine 11309
    m_sImage = s
vbwProfiler.vbwProcOut 628
vbwProfiler.vbwExecuteLine 11310
End Property
Private Property Get cINode_Image() As String
vbwProfiler.vbwProcIn 629
vbwProfiler.vbwExecuteLine 11311
    cINode_Image = m_sImage
vbwProfiler.vbwProcOut 629
vbwProfiler.vbwExecuteLine 11312
End Property
'///////////////////////////////////////////////////

