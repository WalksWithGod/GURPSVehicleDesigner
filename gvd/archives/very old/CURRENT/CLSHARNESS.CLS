VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHarness"
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
Private mvarClassDebugID As Variant 'local copy
Private mvarBeast As String 'local copy
Private mvarQuantity As Integer 'local copy
Private mvarTL As Integer 'local copy
Private mvarWeight As Double 'local copy
Private mvarCost As Integer 'local copy
Private mvarEfficiency As String 'local copy
Private mvarMotivePower As Integer 'local copy
Private mvarMotiveThrust As Integer 'local copy
Private mvarSubType As String 'local copy
Private mvarLocation As String 'local copy
Private mvarParent As String 'local copy
Private mvarKey As String 'local copy
Private mvarDR As Variant 'local copy
Private mvarMinTL As Integer 'local copy
Private mvarText As String 'local copy
'local variable(s) to hold property value(s)
Private mvarBeastHexes As Integer 'local copy
Private mvarBeastST As Integer 'local copy
Public Property Let BeastST(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BeastST = 5
    mvarBeastST = vData
End Property


Public Property Get BeastST() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BeastST
    BeastST = mvarBeastST
End Property



Public Property Let BeastHexes(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BeastHexes = 5
    mvarBeastHexes = vData
End Property


Public Property Get BeastHexes() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BeastHexes
    BeastHexes = mvarBeastHexes
End Property



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



Public Property Let MotiveThrust(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MotiveThrust = 5
    mvarMotiveThrust = vData
End Property


Public Property Get MotiveThrust() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MotiveThrust
    MotiveThrust = mvarMotiveThrust
End Property



Public Property Let MotivePower(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MotivePower = 5
    mvarMotivePower = vData
End Property


Public Property Get MotivePower() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MotivePower
    MotivePower = mvarMotivePower
End Property



Public Property Let Efficiency(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Efficiency = 5
    mvarEfficiency = vData
End Property


Public Property Get Efficiency() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Efficiency
    Efficiency = mvarEfficiency
End Property



Public Property Let Cost(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cost = 5
    mvarCost = vData
End Property


Public Property Get Cost() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cost
    Cost = mvarCost
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



Public Property Let Quantity(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quantity = 5
    mvarQuantity = vData
End Property


Public Property Get Quantity() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quantity
    Quantity = mvarQuantity
End Property



Public Property Let Beast(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Beast = 5
    mvarBeast = vData
End Property


Public Property Get Beast() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Beast
    Beast = mvarBeast
End Property



Public Property Get ClassDebugID() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClassDebugID
    If IsObject(mvarClassDebugID) Then
        Set ClassDebugID = mvarClassDebugID
    Else
        ClassDebugID = mvarClassDebugID
    End If
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



