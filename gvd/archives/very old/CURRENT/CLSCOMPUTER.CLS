VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 1


#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    Private mlClassDebugID As Long
#End If
Private mvarTL As Integer 'local copy
Private mvarWeight As Double 'local copy
Private mvarVolume As Double 'local copy
Private mvarCost As Double 'local copy
Private mvarPowerReqt As Double 'local copy
Private mvarComplexity As String 'local copy
Private mvarCompact As String 'local copy
Private mvarDedicated As String 'local copy
Private mvarDumb As String 'local copy
Private mvarGenius As String 'local copy
Private mvarHardened As String 'local copy
Private mvarHighCapacity As String 'local copy
Private mvarNeuralNet As String 'local copy
Private mvarRobotBrain As String 'local copy
Private mvarSentient As String 'local copy
Private mvarSubType As String 'local copy
Private mvarLocation As String 'local copy
Private mvarParent As String 'local copy
Private mvarKey As String 'local copy
Private mvarDR As Variant 'local copy
Private mvarRuggedized As Boolean 'local copy
Private mvarSurfaceArea As Double 'local copy
Private mvarHitPoints As Integer 'local copy
Private mvarMinTL As Integer 'local copy
Private mvarText As String 'local copy
Public Property Let Text(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Text = 5
    mvarText = vData
End Property


Public Property Get Text() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Text
    Text = mvarText
End Property



Public Property Let MinTL(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MinTL = 5
    mvarMinTL = vData
End Property


Public Property Get MinTL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinTL
    MinTL = mvarMinTL
End Property



Public Property Let HitPoints(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HitPoints = 5
    mvarHitPoints = vData
End Property


Public Property Get HitPoints() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HitPoints
    HitPoints = mvarHitPoints
End Property



Public Property Let SurfaceArea(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SurfaceArea = 5
    mvarSurfaceArea = vData
End Property


Public Property Get SurfaceArea() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SurfaceArea
    SurfaceArea = mvarSurfaceArea
End Property



Public Property Let Ruggedized(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ruggedized = 5
    mvarRuggedized = vData
End Property


Public Property Get Ruggedized() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ruggedized
    Ruggedized = mvarRuggedized
End Property



Public Property Let DR(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR = 5
    mvarDR = vData
End Property


Public Property Set DR(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DR = Form1
    Set mvarDR = vData
End Property


Public Property Get DR() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DR
    If IsObject(mvarDR) Then
        Set DR = mvarDR
    Else
        DR = mvarDR
    End If
End Property



Public Property Let Key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property



Public Property Let Parent(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Parent = 5
    mvarParent = vData
End Property


Public Property Get Parent() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parent
    Parent = mvarParent
End Property



Public Property Let Location(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Location = 5
    mvarLocation = vData
End Property


Public Property Get Location() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Location
    Location = mvarLocation
End Property



Public Property Let SubType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubType = 5
    mvarSubType = vData
End Property


Public Property Get SubType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubType
    SubType = mvarSubType
End Property



Public Property Let Sentient(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sentient = 5
    mvarSentient = vData
End Property


Public Property Get Sentient() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sentient
    Sentient = mvarSentient
End Property



Public Property Let RobotBrain(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RobotBrain = 5
    mvarRobotBrain = vData
End Property


Public Property Get RobotBrain() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RobotBrain
    RobotBrain = mvarRobotBrain
End Property



Public Property Let NeuralNet(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NeuralNet = 5
    mvarNeuralNet = vData
End Property


Public Property Get NeuralNet() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NeuralNet
    NeuralNet = mvarNeuralNet
End Property



Public Property Let HighCapacity(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HighCapacity = 5
    mvarHighCapacity = vData
End Property


Public Property Get HighCapacity() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HighCapacity
    HighCapacity = mvarHighCapacity
End Property



Public Property Let Hardened(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Hardened = 5
    mvarHardened = vData
End Property


Public Property Get Hardened() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Hardened
    Hardened = mvarHardened
End Property



Public Property Let Genius(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Genius = 5
    mvarGenius = vData
End Property


Public Property Get Genius() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Genius
    Genius = mvarGenius
End Property



Public Property Let Dumb(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Dumb = 5
    mvarDumb = vData
End Property


Public Property Get Dumb() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Dumb
    Dumb = mvarDumb
End Property



Public Property Let Dedicated(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Dedicated = 5
    mvarDedicated = vData
End Property


Public Property Get Dedicated() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Dedicated
    Dedicated = mvarDedicated
End Property



Public Property Let Compact(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Compact = 5
    mvarCompact = vData
End Property


Public Property Get Compact() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Compact
    Compact = mvarCompact
End Property



Public Property Let Complexity(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Complexity = 5
    mvarComplexity = vData
End Property


Public Property Get Complexity() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Complexity
    Complexity = mvarComplexity
End Property



Public Property Let PowerReqt(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Power = 5
    mvarPowerReqt = vData
End Property


Public Property Get PowerReqt() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Power
    PowerReqt = mvarPowerReqt
End Property



Public Property Let Cost(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cost = 5
    mvarCost = vData
End Property


Public Property Get Cost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cost
    Cost = mvarCost
End Property



Public Property Let Volume(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Volume = 5
    mvarVolume = vData
End Property


Public Property Get Volume() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Volume
    Volume = mvarVolume
End Property



Public Property Let Weight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Weight = 5
    mvarWeight = vData
End Property


Public Property Get Weight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Weight
    Weight = mvarWeight
End Property



Public Property Let TL(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TL = 5
    mvarTL = vData
End Property


Public Property Get TL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    TL = mvarTL
End Property





Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
End Sub


Private Sub Class_Terminate()
    'the class is being destroyed
    #If DebugMode Then
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub



