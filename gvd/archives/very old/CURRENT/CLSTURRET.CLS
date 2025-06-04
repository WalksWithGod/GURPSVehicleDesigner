VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTurret"
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
Private mvarClassDebugID As String 'local copy
Private mvarText As String 'local copy
Private mvarImage As Integer 'local copy
Private mvarSelectedImage As Integer 'local copy
Private mvarParent As String 'local copy
Private mvarKey As String 'local copy
Private mvarSubType As String 'local copy
Private mvarTL As Integer 'local copy
Private mvarWeight As Double 'local copy
Private mvarCost As Double 'local copy
Private mvarFrameStrength As String 'local copy
Private mvarMaterials As String 'local copy
Private mvarResponsive As Boolean 'local copy
Private mvarRobotic As Boolean 'local copy
Private mvarBiomechanical As Boolean 'local copy
Private mvarLivingMetal As Boolean 'local copy
Private mvarSurfaceArea As Double 'local copy
Private mvarVolume As Double 'local copy
Private mvarSlopeL As String 'local copy
Private mvarSlopeR As String 'local copy
Private mvarSlopeF As String 'local copy
Private mvarSlopeB As String 'local copy
Private mvarRotation As String 'local copy
Private mvarCompartmentalization As String 'local copy
Private mvarRotationSpace As Long 'local copy
Private mvarEmptySpace As Double 'local copy
Private mvarHitPoints As Integer 'local copy
Private mvarLocation As String 'local copy
Private mvarDR As Variant 'local copy
Private mvarRuggedized As Boolean 'local copy
Private mvarMinTL As Integer 'local copy
Private mvarAccessSpace As Double 'local copy
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



Public Property Let RotationSpace(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RotationSpace = 5
    mvarRotationSpace = vData
End Property


Public Property Get RotationSpace() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RotationSpace
    RotationSpace = mvarRotationSpace
End Property



Public Property Let Compartmentalization(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Compartmentalization = 5
    mvarCompartmentalization = vData
End Property


Public Property Get Compartmentalization() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Compartmentalization
    Compartmentalization = mvarCompartmentalization
End Property



Public Property Let Rotation(ByVal vData As String)
Attribute Rotation.VB_Description = "Limited (180) or Full (360) "
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Rotation = 5
    mvarRotation = vData
End Property


Public Property Get Rotation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Rotation
    Rotation = mvarRotation
End Property



Public Property Let SlopeB(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlopeB = 5
    mvarSlopeB = vData
End Property


Public Property Get SlopeB() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlopeB
    SlopeB = mvarSlopeB
End Property



Public Property Let SlopeF(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlopeF = 5
    mvarSlopeF = vData
End Property


Public Property Get SlopeF() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlopeF
    SlopeF = mvarSlopeF
End Property



Public Property Let SlopeR(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlopeR = 5
    mvarSlopeR = vData
End Property


Public Property Get SlopeR() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlopeR
    SlopeR = mvarSlopeR
End Property



Public Property Let SlopeL(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlopeL = 5
    mvarSlopeL = vData
End Property


Public Property Get SlopeL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlopeL
    SlopeL = mvarSlopeL
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


Public Property Get ClassDebugID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClassDebugID
    ClassDebugID = mvarClassDebugID
End Property



Private Sub Class_Initialize()
' set the default properties
mvarTL = colVehicle("1_").TL
mvarRotation = "None"
mvarSlopeR = "None"
mvarSlopeL = "None"
mvarSlopeF = "None"
mvarSlopeB = "None"
mvarCompartmentalization = "None"
mvarFrameStrength = "Medium"
mvarMaterials = "Standard"
mvarResponsive = False
mvarRobotic = False
mvarBiomechanical = False
mvarLivingMetal = False
mvarRotationSpace = 0
mvarCost = 0
mvarWeight = 0
mvarVolume = 0
mvarSurfaceArea = 0
mvarHitPoints = 0

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

    Case 1
        frmProperties.Combo1.AddItem "None"
        frmProperties.Combo1.AddItem "Full"
        frmProperties.Combo1.AddItem "Limited"
        PropQuery = Combobox
    Case 2 To 5
        frmProperties.Combo1.AddItem "None"
        frmProperties.Combo1.AddItem "30 degrees"
        frmProperties.Combo1.AddItem "60 degrees"
        PropQuery = Combobox
    Case 6
        frmProperties.Combo1.AddItem "None"
        frmProperties.Combo1.AddItem "Heavy"
        frmProperties.Combo1.AddItem "Total"
        PropQuery = Combobox
    Case 7
        frmProperties.Combo1.AddItem "Super-Light"
        frmProperties.Combo1.AddItem "Extra-Light"
        frmProperties.Combo1.AddItem "Light"
        frmProperties.Combo1.AddItem "Medium"
        frmProperties.Combo1.AddItem "Heavy"
        frmProperties.Combo1.AddItem "Extra-Heavy"
        PropQuery = Combobox
    Case 8
        frmProperties.Combo1.AddItem "Very Cheap"
        frmProperties.Combo1.AddItem "Cheap"
        frmProperties.Combo1.AddItem "Standard"
        frmProperties.Combo1.AddItem "Expensive"
        frmProperties.Combo1.AddItem "Very Expensive"
        frmProperties.Combo1.AddItem "Advanced"
        PropQuery = Combobox
    Case 9 To 12
        frmProperties.Combo1.AddItem "True"
        frmProperties.Combo1.AddItem "False"
        PropQuery = Combobox
    Case 13 To 18
        PropQuery = Textbox
    
    End Select
End Function

Public Sub PropUpdate(iRow As Integer, vNewValue As Variant)
Select Case iRow
    Case 0
        TL = vNewValue
    Case 1
        Rotation = vNewValue
    Case 2
        SlopeR = vNewValue
    Case 3
        SlopeL = vNewValue
    Case 4
        SlopeF = vNewValue
    Case 5
        SlopeB = vNewValue
    Case 6
        Compartmentalization = vNewValue
    Case 7
        FrameStrength = vNewValue
    Case 8
        Materials = vNewValue
    Case 9
        Responsive = vNewValue
    Case 10
        Robotic = vNewValue
    Case 11
        Biomechanical = vNewValue
    Case 12
        LivingMetal = vNewValue
    Case 13
        ' just a textbox for rotation space
    Case 14
        ' just a textbox for cost
    Case 15
        ' just a textbox for weight
    Case 16
        ' just a textbox for volume
    Case 17
        ' just a textbox for surfacearea
    Case 18
        ' textbox for hitpoints
End Select
StatsUpdate
End Sub


Public Sub StatsUpdate()

    ' Calculate the component volume
    mvarVolume = CalcCombinedVolume(mvarKey)
    ' Calculate the new volume based on slope modifier
    mvarVolume = mvarVolume * CalcSlopeMultiplier(mvarKey)
    ' calculate the Rotation Space volume based on turret rotation setting
    If Rotation = "Full" Then
        mvarRotationSpace = mvarVolume * 1.2
    Else
        If Rotation = "Limited" Then
            mvarRotationSpace = mvarVolume * 1.1
        Else
            mvarRotationSpace = mvarVolume
        End If
    End If
    ' calculate the surface
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    ' Calculate the cost
    mvarCost = BasicDesignCost(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials, mvarResponsive, mvarRobotic, mvarBiomechanical, mvarLivingMetal)
    ' Calculate the weight
    mvarWeight = BasicDesignWeight(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials)
    ' Calculate the Hit Points
    mvarHitPoints = CalcHitPoints(TypeName(Me), mvarFrameStrength, mvarSurfaceArea)
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


