VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRule"
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
'Private m_oChildren() As cINode  ' this is a leaf node and doesnt make use of childnodes
Private m_lngAttributes As Long
Private m_hParent As Long
Private m_hMe As Long
Private m_sName As String
Private m_sDescription As String
Private m_sImage As String
'---------

Private m_lngConvertFrom As Long
Private m_lngConvertTo As Long
Private m_sCategory As String
Private m_lngRoundType As ROUND_OPTION
Private m_lngRoundDigits As Long
Private m_bUseThousandSeperators As Boolean
Private m_bAppendPostfix As Boolean  ' whether we includes things like "lbs" and "feet" in the returned string
Private m_lngExpressionCount As Long
Private m_lngCurrentExpression As Long
Private m_Expressions() As uExpression



Private Sub Class_Initialize()
vbwProfiler.vbwProcIn 556
vbwProfiler.vbwExecuteLine 11041
    m_lngExpressionCount = 0
vbwProfiler.vbwExecuteLine 11042
    m_lngRoundDigits = 2
vbwProfiler.vbwExecuteLine 11043
    m_lngRoundType = ROUND_NONE
vbwProfiler.vbwExecuteLine 11044
    m_bUseThousandSeperators = True
vbwProfiler.vbwExecuteLine 11045
    m_bAppendPostfix = True
vbwProfiler.vbwExecuteLine 11046
    m_lngConvertTo = -1
vbwProfiler.vbwExecuteLine 11047
    m_lngConvertFrom = -1
vbwProfiler.vbwProcOut 556
vbwProfiler.vbwExecuteLine 11048
End Sub
Private Sub Class_Terminate()
vbwProfiler.vbwProcIn 557
vbwProfiler.vbwExecuteLine 11049
    Erase m_Expressions
vbwProfiler.vbwProcOut 557
vbwProfiler.vbwExecuteLine 11050
End Sub

Public Sub addExpression(ByVal dValue As Double, ByVal eType As EXPRESSION_TYPE, ByVal eEval As EVALUATOR_TYPE)
vbwProfiler.vbwProcIn 558
vbwProfiler.vbwExecuteLine 11051
    m_lngExpressionCount = m_lngExpressionCount + 1
vbwProfiler.vbwExecuteLine 11052
    ReDim Preserve m_Expressions(m_lngExpressionCount - 1) ' -1 since our arrays starts at 0

    'todo: should i loop and make sure that an expression of same CONDITION
    ' as another cant have the same evaluator?  I mean, i know we dont want to allow it
    ' but should we check it ihere?  probably...
vbwProfiler.vbwExecuteLine 11053
    With m_Expressions(m_lngExpressionCount - 1)
vbwProfiler.vbwExecuteLine 11054
        .evaluator = eEval
vbwProfiler.vbwExecuteLine 11055
        .type = eType
vbwProfiler.vbwExecuteLine 11056
        .value = dValue
vbwProfiler.vbwExecuteLine 11057
    End With
vbwProfiler.vbwProcOut 558
vbwProfiler.vbwExecuteLine 11058
End Sub

Public Function expressionExists(ByVal eType As EXPRESSION_TYPE, ByVal eEval As EVALUATOR_TYPE) As Boolean
vbwProfiler.vbwProcIn 559
    Dim i As Long
vbwProfiler.vbwExecuteLine 11059
    For i = 0 To m_lngExpressionCount - 1
vbwProfiler.vbwExecuteLine 11060
        If m_Expressions(i).type = eType Then
vbwProfiler.vbwExecuteLine 11061
            If m_Expressions(i).evaluator = eEval Then
vbwProfiler.vbwExecuteLine 11062
                expressionExists = True
vbwProfiler.vbwProcOut 559
vbwProfiler.vbwExecuteLine 11063
                Exit Function
            End If
vbwProfiler.vbwExecuteLine 11064 'B
        End If
vbwProfiler.vbwExecuteLine 11065 'B
vbwProfiler.vbwExecuteLine 11066
    Next
vbwProfiler.vbwProcOut 559
vbwProfiler.vbwExecuteLine 11067
End Function
Public Sub removeExpression(ByVal eType As EXPRESSION_TYPE, ByVal eEval As EVALUATOR_TYPE)
vbwProfiler.vbwProcIn 560
    Dim i As Long
    Dim j As Long
vbwProfiler.vbwExecuteLine 11068
    For i = 0 To m_lngExpressionCount - 1
vbwProfiler.vbwExecuteLine 11069
        If m_Expressions(i).type = eType Then
vbwProfiler.vbwExecuteLine 11070
            If m_Expressions(i).evaluator = eEval Then
                ' this is the one we want to remove
vbwProfiler.vbwExecuteLine 11071
                For j = i To m_lngExpressionCount - 2
vbwProfiler.vbwExecuteLine 11072
                    m_Expressions(j) = m_Expressions(j + 1)
vbwProfiler.vbwExecuteLine 11073
                Next
vbwProfiler.vbwExecuteLine 11074
                m_lngExpressionCount = m_lngExpressionCount - 1
vbwProfiler.vbwExecuteLine 11075
                If m_lngExpressionCount > 0 Then

vbwProfiler.vbwExecuteLine 11076
                ReDim Preserve m_Expressions(0 To m_lngExpressionCount - 1)
                Else
vbwProfiler.vbwExecuteLine 11077 'B
vbwProfiler.vbwExecuteLine 11078
                    Erase m_Expressions()
                End If
vbwProfiler.vbwExecuteLine 11079 'B
vbwProfiler.vbwProcOut 560
vbwProfiler.vbwExecuteLine 11080
                Exit Sub
            End If
vbwProfiler.vbwExecuteLine 11081 'B
        End If
vbwProfiler.vbwExecuteLine 11082 'B
vbwProfiler.vbwExecuteLine 11083
    Next
vbwProfiler.vbwProcOut 560
vbwProfiler.vbwExecuteLine 11084
End Sub

Public Sub setExpressionValue(ByVal index As Long, ByVal dbl As Double)
vbwProfiler.vbwProcIn 561
vbwProfiler.vbwExecuteLine 11085
    m_Expressions(index).value = dbl
vbwProfiler.vbwProcOut 561
vbwProfiler.vbwExecuteLine 11086
End Sub

Public Function getExpression(ByVal index As Long) As Long
vbwProfiler.vbwProcIn 562
vbwProfiler.vbwExecuteLine 11087
    If index <= m_lngExpressionCount - 1 Then
vbwProfiler.vbwExecuteLine 11088
        On Error Resume Next
vbwProfiler.vbwExecuteLine 11089
        getExpression = VarPtr(m_Expressions(index))
vbwProfiler.vbwProcOut 562
vbwProfiler.vbwExecuteLine 11090
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 11091 'B
vbwProfiler.vbwExecuteLine 11092
    getExpression = 0
vbwProfiler.vbwProcOut 562
vbwProfiler.vbwExecuteLine 11093
End Function

Public Property Get expressionCount() As Long
vbwProfiler.vbwProcIn 563
vbwProfiler.vbwExecuteLine 11094
    expressionCount = m_lngExpressionCount
vbwProfiler.vbwProcOut 563
vbwProfiler.vbwExecuteLine 11095
End Property

Public Property Let Name(s As String)
    'm_sName = s 'todo: is it proper design to send this over to the cINode_Name property instead?
vbwProfiler.vbwProcIn 564
vbwProfiler.vbwExecuteLine 11096
    cINode_Name = s
vbwProfiler.vbwProcOut 564
vbwProfiler.vbwExecuteLine 11097
End Property
Public Property Get Name() As String
    'Name = m_sDescription
vbwProfiler.vbwProcIn 565
vbwProfiler.vbwExecuteLine 11098
    Name = cINode_Description
vbwProfiler.vbwProcOut 565
vbwProfiler.vbwExecuteLine 11099
End Property

Public Property Get Category() As String
vbwProfiler.vbwProcIn 566
vbwProfiler.vbwExecuteLine 11100
    Category = m_sCategory
vbwProfiler.vbwProcOut 566
vbwProfiler.vbwExecuteLine 11101
End Property
Public Property Let Category(s As String)
vbwProfiler.vbwProcIn 567
vbwProfiler.vbwExecuteLine 11102
    m_sCategory = s
vbwProfiler.vbwProcOut 567
vbwProfiler.vbwExecuteLine 11103
End Property
Public Property Get convertFrom() As Long
vbwProfiler.vbwProcIn 568
vbwProfiler.vbwExecuteLine 11104
    convertFrom = m_lngConvertFrom
vbwProfiler.vbwProcOut 568
vbwProfiler.vbwExecuteLine 11105
End Property
Public Property Let convertFrom(ByVal i As Long)
vbwProfiler.vbwProcIn 569
vbwProfiler.vbwExecuteLine 11106
    m_lngConvertFrom = i
    ' note: when this changes, we must reset convertTO as well
vbwProfiler.vbwExecuteLine 11107
    convertTo = -1
vbwProfiler.vbwProcOut 569
vbwProfiler.vbwExecuteLine 11108
End Property
Public Property Get convertTo() As Long
vbwProfiler.vbwProcIn 570
vbwProfiler.vbwExecuteLine 11109
    convertTo = m_lngConvertTo
vbwProfiler.vbwProcOut 570
vbwProfiler.vbwExecuteLine 11110
End Property
Public Property Let convertTo(ByVal i As Long)
vbwProfiler.vbwProcIn 571
vbwProfiler.vbwExecuteLine 11111
    m_lngConvertTo = i
vbwProfiler.vbwProcOut 571
vbwProfiler.vbwExecuteLine 11112
End Property
Public Property Get RoundType() As ROUND_OPTION
vbwProfiler.vbwProcIn 572
vbwProfiler.vbwExecuteLine 11113
    RoundType = m_lngRoundType
vbwProfiler.vbwProcOut 572
vbwProfiler.vbwExecuteLine 11114
End Property
Public Property Let RoundType(ByVal i As ROUND_OPTION)
vbwProfiler.vbwProcIn 573
vbwProfiler.vbwExecuteLine 11115
    m_lngRoundType = i
vbwProfiler.vbwProcOut 573
vbwProfiler.vbwExecuteLine 11116
End Property
Public Property Let roundDigits(ByVal i As Long)
vbwProfiler.vbwProcIn 574
vbwProfiler.vbwExecuteLine 11117
    m_lngRoundDigits = i
vbwProfiler.vbwProcOut 574
vbwProfiler.vbwExecuteLine 11118
End Property
Public Property Get roundDigits() As Long
vbwProfiler.vbwProcIn 575
vbwProfiler.vbwExecuteLine 11119
    roundDigits = m_lngRoundDigits
vbwProfiler.vbwProcOut 575
vbwProfiler.vbwExecuteLine 11120
End Property
Public Property Let useThousandSeperators(ByVal b As Boolean)
vbwProfiler.vbwProcIn 576
vbwProfiler.vbwExecuteLine 11121
    m_bUseThousandSeperators = b
vbwProfiler.vbwProcOut 576
vbwProfiler.vbwExecuteLine 11122
End Property
Public Property Get useThousandSeperators() As Boolean
vbwProfiler.vbwProcIn 577
vbwProfiler.vbwExecuteLine 11123
    useThousandSeperators = m_bUseThousandSeperators
vbwProfiler.vbwProcOut 577
vbwProfiler.vbwExecuteLine 11124
End Property
Public Property Let appendPostfix(ByVal b As Boolean)
vbwProfiler.vbwProcIn 578
vbwProfiler.vbwExecuteLine 11125
    m_bAppendPostfix = b
vbwProfiler.vbwProcOut 578
vbwProfiler.vbwExecuteLine 11126
End Property
Public Property Get appendPostfix() As Boolean
vbwProfiler.vbwProcIn 579
vbwProfiler.vbwExecuteLine 11127
    appendPostfix = m_bAppendPostfix
vbwProfiler.vbwProcOut 579
vbwProfiler.vbwExecuteLine 11128
End Property

Public Function testRule(ByVal dbl As Double) As Boolean
vbwProfiler.vbwProcIn 580
    Dim i As Long
    Dim bFlag As Boolean

    ' if any condition is not true, we do not apply the rule
vbwProfiler.vbwExecuteLine 11129
    For i = 0 To UBound(m_Expressions)
        'If Not evaluatesTrue(dbl, m_Expressions(i).value, m_Expressions(i).evaluator) Then testRule = False: Exit Function
vbwProfiler.vbwExecuteLine 11130
    Next

    ' passed all conditions and exceptions
vbwProfiler.vbwExecuteLine 11131
    testRule = True
vbwProfiler.vbwProcOut 580
vbwProfiler.vbwExecuteLine 11132
End Function


'///////////////////////////////////////////////////
'//cIPersist Implemented Properties and Functions
Private Property Get cIPersist_Classname() As String
vbwProfiler.vbwProcIn 581
vbwProfiler.vbwProcOut 581
vbwProfiler.vbwExecuteLine 11133
End Property
Private Property Get cIPersist_GUID() As String
vbwProfiler.vbwProcIn 582
vbwProfiler.vbwProcOut 582
vbwProfiler.vbwExecuteLine 11134
End Property
Private Sub cIPersist_LoadProperties(ByVal op As PersistenceManager.clsObjProperties, ByVal iMode As Long)
vbwProfiler.vbwProcIn 583
    Dim i As Long
vbwProfiler.vbwExecuteLine 11135
    m_sName = op.Load("name")
vbwProfiler.vbwExecuteLine 11136
    m_sDescription = op.Load("description")
vbwProfiler.vbwExecuteLine 11137
    m_lngRoundType = op.Load("roundtype")
vbwProfiler.vbwExecuteLine 11138
    m_lngRoundDigits = op.Load("rounddigits")
vbwProfiler.vbwExecuteLine 11139
    m_bUseThousandSeperators = op.Load("thousandseperators")
vbwProfiler.vbwExecuteLine 11140
    m_bAppendPostfix = op.Load("appendsuffix")
vbwProfiler.vbwExecuteLine 11141
    m_lngConvertTo = op.Load("convertto")
vbwProfiler.vbwExecuteLine 11142
    m_lngConvertFrom = op.Load("convertfrom")
vbwProfiler.vbwExecuteLine 11143
    m_sCategory = op.Load("category")
vbwProfiler.vbwExecuteLine 11144
    m_lngExpressionCount = op.Load("expressioncount")

vbwProfiler.vbwExecuteLine 11145
    If m_lngExpressionCount > 0 Then
vbwProfiler.vbwExecuteLine 11146
        ReDim m_Expressions(m_lngExpressionCount - 1)
vbwProfiler.vbwExecuteLine 11147
        For i = 0 To m_lngExpressionCount - 1 'todo make sure bounds are ok
vbwProfiler.vbwExecuteLine 11148
            With m_Expressions(i)
vbwProfiler.vbwExecuteLine 11149
                .type = op.Load("expression" & i & ".type")
vbwProfiler.vbwExecuteLine 11150
                .value = op.Load("expression" & i & ".value")
vbwProfiler.vbwExecuteLine 11151
                .evaluator = op.Load("expression" & i & ".evalcode")
vbwProfiler.vbwExecuteLine 11152
            End With
vbwProfiler.vbwExecuteLine 11153
        Next
    End If
vbwProfiler.vbwExecuteLine 11154 'B
vbwProfiler.vbwProcOut 583
vbwProfiler.vbwExecuteLine 11155
End Sub
Private Sub cIPersist_StoreProperties(ByVal op As PersistenceManager.clsObjProperties)
vbwProfiler.vbwProcIn 584
    Dim i As Long

vbwProfiler.vbwExecuteLine 11156
    op.Store "classname", TypeName(Me)
vbwProfiler.vbwExecuteLine 11157
    op.Store "defpath", ""
vbwProfiler.vbwExecuteLine 11158
    op.Store "guid", ""

vbwProfiler.vbwExecuteLine 11159
    op.Store "name", m_sName
vbwProfiler.vbwExecuteLine 11160
    op.Store "description", m_sDescription
vbwProfiler.vbwExecuteLine 11161
    op.Store "roundtype", m_lngRoundType
vbwProfiler.vbwExecuteLine 11162
    op.Store "rounddigits", m_lngRoundDigits
vbwProfiler.vbwExecuteLine 11163
    op.Store "thousandseperators", m_bUseThousandSeperators
vbwProfiler.vbwExecuteLine 11164
    op.Store "appendsuffix", m_bAppendPostfix
vbwProfiler.vbwExecuteLine 11165
    op.Store "convertto", m_lngConvertTo
vbwProfiler.vbwExecuteLine 11166
    op.Store "convertfrom", m_lngConvertFrom
vbwProfiler.vbwExecuteLine 11167
    op.Store "category", m_sCategory
vbwProfiler.vbwExecuteLine 11168
    op.Store "expressioncount", m_lngExpressionCount

vbwProfiler.vbwExecuteLine 11169
    For i = 0 To m_lngExpressionCount - 1
vbwProfiler.vbwExecuteLine 11170
        With m_Expressions(i)
vbwProfiler.vbwExecuteLine 11171
            op.Store "expression" & i & ".type", .type
vbwProfiler.vbwExecuteLine 11172
            op.Store "expression" & i & ".value", .value
vbwProfiler.vbwExecuteLine 11173
            op.Store "expression" & i & ".evalcode", .evaluator
vbwProfiler.vbwExecuteLine 11174
        End With
vbwProfiler.vbwExecuteLine 11175
    Next
vbwProfiler.vbwProcOut 584
vbwProfiler.vbwExecuteLine 11176
End Sub
'///////////////////////////////////////////////////
'//cINode Implemented Properties and Functions
Private Function cINode_AddChild(oBase As cINode) As Boolean
   ' this is a leaf and cant have any children
vbwProfiler.vbwProcIn 585
vbwProfiler.vbwExecuteLine 11177
   cINode_AddChild = False
vbwProfiler.vbwProcOut 585
vbwProfiler.vbwExecuteLine 11178
End Function
Private Function cINode_getChildrenByClassName(ByRef Classname As String, hChilds() As Long) As Boolean
vbwProfiler.vbwProcIn 586
vbwProfiler.vbwProcOut 586
vbwProfiler.vbwExecuteLine 11179
End Function
Private Function cINode_getChildIndexByHandle(ByVal h As Long) As Long
vbwProfiler.vbwProcIn 587
vbwProfiler.vbwProcOut 587
vbwProfiler.vbwExecuteLine 11180
End Function
Private Function cINode_getChild(ByVal lngIndex As Long) As cINode
    'this is a leaf and had no children ' is this how to handle these? or make them more generic?
    ' we can also check that it _does not_ implement cIContainer
vbwProfiler.vbwProcIn 588
vbwProfiler.vbwProcOut 588
vbwProfiler.vbwExecuteLine 11181
End Function
Private Property Get cINode_childCount() As Long
vbwProfiler.vbwProcIn 589
vbwProfiler.vbwProcOut 589
vbwProfiler.vbwExecuteLine 11182
End Property
Private Function cINode_removeChild(ByVal lngIndex As Long) As Boolean
vbwProfiler.vbwProcIn 590
vbwProfiler.vbwExecuteLine 11183
    cINode_removeChild = False
vbwProfiler.vbwProcOut 590
vbwProfiler.vbwExecuteLine 11184
End Function
Private Property Get cINode_ClassName() As String
vbwProfiler.vbwProcIn 591
vbwProfiler.vbwExecuteLine 11185
    cINode_ClassName = TypeName(Me)
vbwProfiler.vbwProcOut 591
vbwProfiler.vbwExecuteLine 11186
End Property
Private Property Get cINode_Attributes() As Long
vbwProfiler.vbwProcIn 592
vbwProfiler.vbwExecuteLine 11187
    cINode_Attributes = m_lngAttributes
vbwProfiler.vbwProcOut 592
vbwProfiler.vbwExecuteLine 11188
End Property
Private Property Let cINode_Handle(ByVal l As Long)
vbwProfiler.vbwProcIn 593
vbwProfiler.vbwExecuteLine 11189
    m_hMe = l
vbwProfiler.vbwProcOut 593
vbwProfiler.vbwExecuteLine 11190
End Property
Private Property Get cINode_Handle() As Long
vbwProfiler.vbwProcIn 594
vbwProfiler.vbwExecuteLine 11191
    cINode_Handle = m_hMe
vbwProfiler.vbwProcOut 594
vbwProfiler.vbwExecuteLine 11192
End Property
Private Property Let cINode_Parent(ByVal l As Long)
vbwProfiler.vbwProcIn 595
vbwProfiler.vbwProcOut 595
vbwProfiler.vbwExecuteLine 11193
End Property
Private Property Get cINode_Parent() As Long
vbwProfiler.vbwProcIn 596
vbwProfiler.vbwProcOut 596
vbwProfiler.vbwExecuteLine 11194
End Property
Private Property Let cINode_Name(ByVal s As String)
vbwProfiler.vbwProcIn 597
vbwProfiler.vbwExecuteLine 11195
    m_sName = s
vbwProfiler.vbwProcOut 597
vbwProfiler.vbwExecuteLine 11196
End Property
Private Property Get cINode_Name() As String
vbwProfiler.vbwProcIn 598
vbwProfiler.vbwExecuteLine 11197
    cINode_Name = m_sName
vbwProfiler.vbwProcOut 598
vbwProfiler.vbwExecuteLine 11198
End Property
Private Property Let cINode_Description(ByVal s As String)
vbwProfiler.vbwProcIn 599
vbwProfiler.vbwExecuteLine 11199
    m_sDescription = s
vbwProfiler.vbwProcOut 599
vbwProfiler.vbwExecuteLine 11200
End Property
Private Property Get cINode_Description() As String
vbwProfiler.vbwProcIn 600
vbwProfiler.vbwExecuteLine 11201
    cINode_Description = m_sDescription
vbwProfiler.vbwProcOut 600
vbwProfiler.vbwExecuteLine 11202
End Property
Private Property Let cINode_Image(ByVal s As String)
vbwProfiler.vbwProcIn 601
vbwProfiler.vbwExecuteLine 11203
    m_sImage = s
vbwProfiler.vbwProcOut 601
vbwProfiler.vbwExecuteLine 11204
End Property
Private Property Get cINode_Image() As String
vbwProfiler.vbwProcIn 602
vbwProfiler.vbwExecuteLine 11205
    cINode_Image = m_sImage
vbwProfiler.vbwProcOut 602
vbwProfiler.vbwExecuteLine 11206
End Property

'///////////////////////////////////////////////////




