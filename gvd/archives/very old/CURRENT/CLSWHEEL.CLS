VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWheel"
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
'local variable(s) to hold property value(s)
Private mvarSubType As String 'local copy
Private mvarParent As String 'local copy
Private mvarKey As String 'local copy
Private mvarText As String 'local copy
Private mvarImage As Integer 'local copy
Private mvarSelectedImage As Integer 'local copy
Private mvarTL As Integer 'local copy
Private mvarNumber As String 'local copy
Private mvarImprovedSuspension As Boolean 'local copy
Private mvarWheelblades As String 'local copy
Private mvarSnowTires As Boolean 'local copy
Private mvarRacingTires As Boolean 'local copy
Private mvarPunctureResistant As Boolean 'local copy
Private mvarRetractLocation As String 'local copy
Private mvarImprovedBreaks As Boolean 'local copy
Private mvarAllwheelSteering As Boolean 'local copy
Private mvarAllWheelDrive As Boolean
Private mvarSmartwheels As Boolean 'local copy
Private mvarFrameStrength As String 'local copy
Private mvarMaterials As String 'local copy
Private mvarResponsive As Boolean 'local copy
Private mvarRobotic As Boolean 'local copy
Private mvarBiomechanical As Boolean 'local copy
Private mvarLivingMetal As Boolean 'local copy
Private mvarCost As Double 'local copy
Private mvarWeight As Double 'local copy
Private mvarSurfaceArea As Double 'local copy
Private mvarVolume As Double 'local copy
Private mvarHitPoints As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarLocation As String 'local copy
'local variable(s) to hold property value(s)
Private mvarDR As Variant 'local copy
Private mvarRuggedized As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarMinTL As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarAccessSpace As Double 'local copy
Private mvarEmptySpace As Double 'local copy
Public Property Let EmptySpace(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmptySpace = 5
    mvarEmptySpace = vData
End Property


Public Property Get EmptySpace() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmptySpace
    EmptySpace = mvarEmptySpace
End Property



Public Property Let AccessSpace(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccessSpace = 5
    mvarAccessSpace = vData
End Property


Public Property Get AccessSpace() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccessSpace
    AccessSpace = mvarAccessSpace
End Property

Public Property Let AllWheelDrive(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AllWheelDrive = 5
    mvarAllWheelDrive = vData

End Property


Public Property Get AllWheelDrive() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AllWheelDrive
    AllWheelDrive = mvarAllWheelDrive
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



Public Property Let SubType(ByVal vData As String)
Attribute SubType.VB_Description = "Name of the wheels"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Subtype = 5
    mvarSubType = vData
End Property


Public Property Get SubType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Subtype
    SubType = mvarSubType
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



Public Property Let LivingMetal(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LivingMetal = 5
    mvarLivingMetal = vData
End Property


Public Property Get LivingMetal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LivingMetal
    LivingMetal = mvarLivingMetal
End Property



Public Property Let Biomechanical(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Biomechanical = 5
    mvarBiomechanical = vData
End Property


Public Property Get Biomechanical() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Biomechanical
    Biomechanical = mvarBiomechanical
End Property



Public Property Let Robotic(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Robotic = 5
    mvarRobotic = vData
End Property


Public Property Get Robotic() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Robotic
    Robotic = mvarRobotic
End Property



Public Property Let Responsive(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Responsive = 5
    mvarResponsive = vData
End Property


Public Property Get Responsive() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Responsive
    Responsive = mvarResponsive
End Property



Public Property Let Materials(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Materials = 5
    mvarMaterials = vData
End Property


Public Property Get Materials() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Materials
    Materials = mvarMaterials
End Property



Public Property Let FrameStrength(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FrameStrength = 5
    mvarFrameStrength = vData
End Property


Public Property Get FrameStrength() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FrameStrength
    FrameStrength = mvarFrameStrength
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



Public Property Let SelectedImage(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SelectedImage = 5
    mvarSelectedImage = vData
End Property


Public Property Get SelectedImage() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedImage
    SelectedImage = mvarSelectedImage
End Property



Public Property Let Image(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Image = 5
    mvarImage = vData
End Property


Public Property Get Image() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Image
    Image = mvarImage
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



Public Property Let Smartwheels(ByVal vData As Boolean)
    On Error GoTo SmartwheelsLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Smartwheels = 5
    mvarSmartwheels = vData
    Exit Property


SmartwheelsLetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:Smartwheels Property Let")
End Property


Public Property Get Smartwheels() As Boolean
    On Error GoTo SmartwheelsGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Smartwheels
    Smartwheels = mvarSmartwheels
    Exit Property


SmartwheelsGetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:Smartwheels Property Get")
End Property



Public Property Let AllwheelSteering(ByVal vData As Boolean)
    On Error GoTo AllwheelSteeringLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AllwheelSteering = 5
    mvarAllwheelSteering = vData
    Exit Property


AllwheelSteeringLetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:AllwheelSteering Property Let")
End Property


Public Property Get AllwheelSteering() As Boolean
    On Error GoTo AllwheelSteeringGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AllwheelSteering
    AllwheelSteering = mvarAllwheelSteering
    Exit Property


AllwheelSteeringGetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:AllwheelSteering Property Get")
End Property



Public Property Let ImprovedBreaks(ByVal vData As Boolean)
    On Error GoTo ImprovedBreaksLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ImprovedBreaks = 5
    mvarImprovedBreaks = vData
    Exit Property


ImprovedBreaksLetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:ImprovedBreaks Property Let")
End Property


Public Property Get ImprovedBreaks() As Boolean
    On Error GoTo ImprovedBreaksGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ImprovedBreaks
    ImprovedBreaks = mvarImprovedBreaks
    Exit Property


ImprovedBreaksGetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:ImprovedBreaks Property Get")
End Property



Public Property Let RetractLocation(ByVal vData As String)
    On Error GoTo RetractLocationLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RetractLocation = 5
    mvarRetractLocation = vData
    Exit Property


RetractLocationLetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:RetractLocation Property Let")
End Property


Public Property Get RetractLocation() As String
    On Error GoTo RetractLocationGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RetractLocation
    RetractLocation = mvarRetractLocation
    Exit Property


RetractLocationGetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:RetractLocation Property Get")
End Property



Public Property Let PunctureResistant(ByVal vData As Boolean)
    On Error GoTo PunctureResistantLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PunctureResistant = 5
    mvarPunctureResistant = vData
    Exit Property


PunctureResistantLetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:PunctureResistant Property Let")
End Property


Public Property Get PunctureResistant() As Boolean
    On Error GoTo PunctureResistantGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PunctureResistant
    PunctureResistant = mvarPunctureResistant
    Exit Property


PunctureResistantGetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:PunctureResistant Property Get")
End Property



Public Property Let RacingTires(ByVal vData As Boolean)
    On Error GoTo RacingTiresLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RacingTires = 5
    mvarRacingTires = vData
    Exit Property


RacingTiresLetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:RacingTires Property Let")
End Property


Public Property Get RacingTires() As Boolean
    On Error GoTo RacingTiresGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RacingTires
    RacingTires = mvarRacingTires
    Exit Property


RacingTiresGetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:RacingTires Property Get")
End Property



Public Property Let SnowTires(ByVal vData As Boolean)
    On Error GoTo SnowTiresLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SnowTires = 5
    mvarSnowTires = vData
    Exit Property


SnowTiresLetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:SnowTires Property Let")
End Property


Public Property Get SnowTires() As Boolean
    On Error GoTo SnowTiresGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SnowTires
    SnowTires = mvarSnowTires
    Exit Property


SnowTiresGetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:SnowTires Property Get")
End Property



Public Property Let Wheelblades(ByVal vData As String)
    On Error GoTo WheelbladesLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Wheelblades = 5
    mvarWheelblades = vData
    Exit Property


WheelbladesLetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:Wheelblades Property Let")
End Property


Public Property Get Wheelblades() As String
    On Error GoTo WheelbladesGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Wheelblades
    Wheelblades = mvarWheelblades
    Exit Property


WheelbladesGetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:Wheelblades Property Get")
End Property



Public Property Let ImprovedSuspension(ByVal vData As Boolean)
Attribute ImprovedSuspension.VB_Description = "Does the wheel sub assembly have improved suspension or not?"
    On Error GoTo ImprovedSuspensionLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ImprovedSuspension = 5
    mvarImprovedSuspension = vData
    Exit Property


ImprovedSuspensionLetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:ImprovedSuspension Property Let")
End Property


Public Property Get ImprovedSuspension() As Boolean
    On Error GoTo ImprovedSuspensionGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ImprovedSuspension
    ImprovedSuspension = mvarImprovedSuspension
    Exit Property


ImprovedSuspensionGetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:ImprovedSuspension Property Get")
End Property



Public Property Let Number(ByVal vData As String)
Attribute Number.VB_Description = "Number of wheels in the subassembly"
    On Error GoTo NumberLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Number = 5
    mvarNumber = vData
    Exit Property


NumberLetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:Number Property Let")
End Property


Public Property Get Number() As String
    On Error GoTo NumberGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Number
    Number = mvarNumber
    Exit Property


NumberGetErr:


    Call RaiseError(MyUnhandledError, "WheelSubAssembly:Number Property Get")
End Property








Private Sub Class_Initialize()
' set the default properties
mvarTL = colVehicle("1_").TL
mvarSubType = "Standard"
mvarNumber = "4"
mvarImprovedSuspension = False
mvarWheelblades = "None"
mvarSnowTires = False
mvarRacingTires = False
mvarPunctureResistant = False
mvarRetractLocation = "None"
mvarImprovedBreaks = False
mvarAllwheelSteering = False
mvarSmartwheels = False
mvarFrameStrength = "Medium"
mvarMaterials = "Standard"
mvarResponsive = False
mvarRobotic = False
mvarBiomechanical = False
mvarLivingMetal = False
mvarCost = 0
mvarWeight = 0
mvarSurfaceArea = 0
mvarVolume = 0
mvarHitPoints = 0
mvarAllWheelDrive = False

    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
End Sub


Private Sub Class_Terminate()
    'the class is being destroyed
    
' DEcrement the number of wings on aircraft
'colVehicle.Item("1_").NumWings = colVehicle.Item("1_").NumWings - 1
'MsgBox colVehicle.Item("1_").NumWings
    
    
    
    #If DebugMode Then
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub


Public Function PropQuery(ByVal Row As Integer) As Integer
Const Combobox = 1
Const Textbox = 2
Const CommandButton = 3
' clear the Properties form's combobox
frmProperties.Combo1.Clear
    Select Case Row
    Case 0
        frmProperties.Combo1.AddItem 0
        frmProperties.Combo1.AddItem 1
        frmProperties.Combo1.AddItem 2
        frmProperties.Combo1.AddItem 3
        frmProperties.Combo1.AddItem 4
        frmProperties.Combo1.AddItem 5
        frmProperties.Combo1.AddItem 6
        frmProperties.Combo1.AddItem 7
        frmProperties.Combo1.AddItem 8
        frmProperties.Combo1.AddItem 9
        frmProperties.Combo1.AddItem 10
        frmProperties.Combo1.AddItem 11
        frmProperties.Combo1.AddItem 12
        frmProperties.Combo1.AddItem 13
        frmProperties.Combo1.AddItem 14
        frmProperties.Combo1.AddItem 15
        frmProperties.Combo1.AddItem 16
        PropQuery = Combobox

    Case 1 ' subtype
        frmProperties.Combo1.AddItem "Standard Wheels"
        frmProperties.Combo1.AddItem "Small Wheels"
        frmProperties.Combo1.AddItem "Heavy Wheels"
        frmProperties.Combo1.AddItem "Railway Wheels"
        frmProperties.Combo1.AddItem "Off-Road Wheels"
        frmProperties.Combo1.AddItem "Retractable Wheels"
        PropQuery = Combobox
    Case 2 ' number of wheels
        frmProperties.Combo1.AddItem "0"
        frmProperties.Combo1.AddItem "1"
        frmProperties.Combo1.AddItem "2"
        frmProperties.Combo1.AddItem "3"
        frmProperties.Combo1.AddItem "4"
        frmProperties.Combo1.AddItem "5"
        frmProperties.Combo1.AddItem "6"
        frmProperties.Combo1.AddItem "7"
        frmProperties.Combo1.AddItem "8"
        frmProperties.Combo1.AddItem "9"
        frmProperties.Combo1.AddItem "10"
        frmProperties.Combo1.AddItem "11"
        frmProperties.Combo1.AddItem "12"
        frmProperties.Combo1.AddItem "13"
        frmProperties.Combo1.AddItem "14"
        frmProperties.Combo1.AddItem "15"
        frmProperties.Combo1.AddItem "16"
        frmProperties.Combo1.AddItem "17"
        frmProperties.Combo1.AddItem "18"
        frmProperties.Combo1.AddItem "19"
        frmProperties.Combo1.AddItem "20"
    Case 4 ' wheelblades
        frmProperties.Combo1.AddItem "None"
        frmProperties.Combo1.AddItem "Fixed"
        frmProperties.Combo1.AddItem "Rectractable"
        PropQuery = Combobox
    Case 18 To 22 ' cost, weight, surface, volume, hitpoints
        PropQuery = Textbox
    Case 8 ' retract location
        frmProperties.Combo1.AddItem "None"
        frmProperties.Combo1.AddItem "Body"
        frmProperties.Combo1.AddItem "Body & Wings"
        PropQuery = Combobox
    Case 12
        frmProperties.Combo1.AddItem "Super-Light"
        frmProperties.Combo1.AddItem "Extra-Light"
        frmProperties.Combo1.AddItem "Light"
        frmProperties.Combo1.AddItem "Medium"
        frmProperties.Combo1.AddItem "Heavy"
        frmProperties.Combo1.AddItem "Extra-Heavy"
        PropQuery = Combobox
    Case 13
        frmProperties.Combo1.AddItem "Very Cheap"
        frmProperties.Combo1.AddItem "Cheap"
        frmProperties.Combo1.AddItem "Standard"
        frmProperties.Combo1.AddItem "Expensive"
        frmProperties.Combo1.AddItem "Very Expensive"
        frmProperties.Combo1.AddItem "Advanced"
        PropQuery = Combobox
    Case 3, 5 To 7, 9 To 11, 14 To 17
        frmProperties.Combo1.AddItem "True"
        frmProperties.Combo1.AddItem "False"
        PropQuery = Combobox
    End Select
End Function

Public Sub PropUpdate(iRow As Integer, vNewValue As Variant)
Select Case iRow
Case 0 ' tech level
    TL = vNewValue
Case 1 ' subtype
    SubType = vNewValue
Case 2 'number of wheels
    Number = vNewValue
Case 3
    ImprovedSuspension = vNewValue
Case 4
    Wheelblades = vNewValue
Case 5
    SnowTires = vNewValue
Case 6
    RacingTires = vNewValue
Case 7
    PunctureResistant = vNewValue
Case 8
    RetractLocation = vNewValue
Case 9
    ImprovedBreaks = vNewValue
Case 10
    AllwheelSteering = vNewValue
Case 11
    Smartwheels = vNewValue
Case 12
    FrameStrength = vNewValue
Case 13
    Materials = vNewValue
Case 14
    Responsive = vNewValue
Case 15
    Robotic = vNewValue
Case 16
    Biomechanical = vNewValue
Case 17
    LivingMetal = vNewValue
Case 18
    ' textbox for Cost
Case 19
    ' textbox for Weight
Case 20
    ' textbox for SurfaceArea
Case 21
    'textbox for Volume
Case 22
    'textbox for HitPoints
End Select
StatsUpdate
End Sub


Public Sub StatsUpdate()
    ' Calculate the component volume
    Select Case SubType
        Case "Small", "Retractable"
            mvarVolume = colVehicle.Item("1_").Volume * 0.05
        Case "Standard"
            mvarVolume = colVehicle.Item("1_").Volume * 0.1
        Case "Heavy", "Off-Road", "Railway"
            mvarVolume = colVehicle.Item("1_").Volume * 0.2
    End Select
    mvarVolume = CalcCombinedVolume(mvarKey)
    ' calculate the surface
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    ' Calculate the cost
    mvarCost = BasicDesignCost(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials, mvarResponsive, mvarRobotic, mvarBiomechanical, mvarLivingMetal)
    ' Calculate the weight
    mvarWeight = BasicDesignWeight(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials)
    ' Calculate the Hit Points
    mvarHitPoints = CalcHitPoints(TypeName(Me), mvarFrameStrength, mvarSurfaceArea, Val(mvarNumber))
End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    colVehicle.Item(Parent).StatsUpdate
End If

End Sub

Public Sub QueryChild()
' if the object has children, query it and check to see if
' more stats/property updates are needed for other objects in the collection
' (see the StatusUpdate property for help on checking for childeren.  Can i use that one in place of this?)
End Sub




