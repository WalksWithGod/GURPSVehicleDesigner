VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCrew"
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
Private mvarCrewStations As String 'local copy
Private mvarRowers As String 'local copy
Private mvarLoaders As String 'local copy
Private mvarSails As String 'local copy
Private mvarLiftingGas As String 'local copy
Private mvarMechanics As String 'local copy
Private mvarPayingPassengers As String 'local copy
Private mvarAuxiliaryVehicles As String 'local copy
Private mvarLaboratories As String 'local copy
Private mvarMedicalEquiptment As String 'local copy
Private mvarServiceCrew As String 'local copy
Private mvarMilitaryCraft As String 'local copy
Private mvarCrewSize As String 'local copy
Private mvarLocation As String 'local copy
Private mvarParent As String 'local copy
Private mvarKey As String 'local copy
Private mvarDR As Variant 'local copy
Private mvarRuggedized As Boolean 'local copy
Private mvarSurfaceArea As Double 'local copy
Private mvarHitPoints As Integer 'local copy
Private mvarTL As Integer 'local copy
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



Public Property Let CrewSize(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CrewSize = 5
    mvarCrewSize = vData
End Property


Public Property Get CrewSize() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CrewSize
    CrewSize = mvarCrewSize
End Property



Public Property Let MilitaryCraft(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MilitaryCraft = 5
    mvarMilitaryCraft = vData
End Property


Public Property Get MilitaryCraft() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MilitaryCraft
    MilitaryCraft = mvarMilitaryCraft
End Property



Public Property Let ServiceCrew(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ServiceCrew = 5
    mvarServiceCrew = vData
End Property


Public Property Get ServiceCrew() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ServiceCrew
    ServiceCrew = mvarServiceCrew
End Property



Public Property Let MedicalEquiptment(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MedicalEquiptment = 5
    mvarMedicalEquiptment = vData
End Property


Public Property Get MedicalEquiptment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MedicalEquiptment
    MedicalEquiptment = mvarMedicalEquiptment
End Property



Public Property Let Laboratories(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Laboratories = 5
    mvarLaboratories = vData
End Property


Public Property Get Laboratories() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Laboratories
    Laboratories = mvarLaboratories
End Property



Public Property Let AuxiliaryVehicles(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AuxiliaryVehicles = 5
    mvarAuxiliaryVehicles = vData
End Property


Public Property Get AuxiliaryVehicles() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AuxiliaryVehicles
    AuxiliaryVehicles = mvarAuxiliaryVehicles
End Property



Public Property Let PayingPassengers(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PayingPassengers = 5
    mvarPayingPassengers = vData
End Property


Public Property Get PayingPassengers() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PayingPassengers
    PayingPassengers = mvarPayingPassengers
End Property



Public Property Let Mechanics(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Mechanics = 5
    mvarMechanics = vData
End Property


Public Property Get Mechanics() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Mechanics
    Mechanics = mvarMechanics
End Property



Public Property Let LiftingGas(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LiftingGas = 5
    mvarLiftingGas = vData
End Property


Public Property Get LiftingGas() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LiftingGas
    LiftingGas = mvarLiftingGas
End Property



Public Property Let Sails(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sails = 5
    mvarSails = vData
End Property


Public Property Get Sails() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sails
    Sails = mvarSails
End Property



Public Property Let Loaders(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Loaders = 5
    mvarLoaders = vData
End Property


Public Property Get Loaders() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Loaders
    Loaders = mvarLoaders
End Property



Public Property Let Rowers(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Rowers = 5
    mvarRowers = vData
End Property


Public Property Get Rowers() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Rowers
    Rowers = mvarRowers
End Property



Public Property Let CrewStations(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CrewStations = 5
    mvarCrewStations = vData
End Property


Public Property Get CrewStations() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CrewStations
    CrewStations = mvarCrewStations
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


