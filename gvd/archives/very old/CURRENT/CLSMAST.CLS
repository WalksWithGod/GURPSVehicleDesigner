VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMast"
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
Private mvarNumber As Integer 'local copy
Private mvarHeight As Integer 'local copy
Private mvarDependantStructure As Boolean 'local copy
Private mvarHitPoints As Integer 'local copy

'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent Change[(arg1, arg2, ... , argn)]
Public Event Change()
Private mvarLocation As String 'local copy
Private mvarDR As Variant 'local copy
Private mvarRuggedized As Boolean 'local copy
Private mvarMinTL As Integer 'local copy
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


Sub Change()
    RaiseEvent Change
    
    
' Const FullRotation = 1.2
' Const LimitedRotation = 1.1

' Volume = 0

' For each component with my key as their parent key
'   volume = volume + the components own volume
' Next

' If rotation = "Full" then
'   volume = volume * FullRotation
' Else
'    If Rotation = "Limited" Then
'        volume = Volume * LimitedRotation
'     Else ' its a fixed mount
'        Volume = Volume
'     End If
' End If
End Sub


Public Property Let DependantStructure(ByVal vData As Boolean)
Attribute DependantStructure.VB_Description = "The only structural subassembly that can be attached atop it is a single open mount and that open mount may not house components whose combined volume in cf is greater than half the mast's height."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DependantStructure = 5
    mvarDependantStructure = vData
End Property


Public Property Get DependantStructure() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DependantStructure
    DependantStructure = mvarDependantStructure
End Property



Public Property Let Height(ByVal vData As Integer)
Attribute Height.VB_Description = "This will need to be like Armor in that it brings up a mini dialog so that users can set seperate height values for each mass.  The ""number"" property will be read only and will be calculated and displayed to user when user defines each mast height."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Height = 5

    Dim MaxHeight As Integer
    mvarHeight = vData
    
    ' check to see if the user has entered an invalid mast height
    If TL <= 4 Then
        MaxHeight = 150
    Else
        If Materials = "Wood" Then
            MaxHeight = 150
        Else
            MaxHeight = 200
        End If
   End If
    ' If the user has entered invalid height, post error message then set Mast to
    ' maximum allowable height
    If mvarHeight > MaxHeight Then
        MsgBox "Maximum Mast height at Tech Level " & TL & " using " & Materials & " materials is " & MaxHeight & " feet."
        Height = MaxHeight
    End If
        
End Property


Public Property Get Height() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Height

   Height = mvarHeight
End Property



Public Property Let Number(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Number = 5
    mvarNumber = vData
End Property


Public Property Get Number() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Number
    Number = mvarNumber
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
    
    If TL = "0" Then
        MsgBox "Minimum Tech Level is 1 for Masts"
        TL = "1"
    End If
End Property


Public Property Get TL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    
    
    TL = mvarTL
    
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



Public Property Get ClassDebugID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClassDebugID
    ClassDebugID = mvarClassDebugID
End Property

Private Sub Class_Initialize()
' set the default properties
mvarSubType = "Temp"
mvarTL = colVehicle(1).TL
mvarMaterials = "Wood"
mvarHeight = 20



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
        PropQuery = Textbox
    Case 8
        PropQuery = Textbox
    Case 9
        frmProperties.Combo1.AddItem "Wood"
        frmProperties.Combo1.AddItem "Metal"
        PropQuery = Combobox
    End Select
End Function

Public Sub PropUpdate(iRow As Integer, vNewValue As Variant)
Select Case iRow
    Case 0
        TL = vNewValue
    Case 1
        Text = vNewValue
    Case 2
    Case 3
    Case 4
    Case 5
    Case 6
    Case 7
        HitPoints = vNewValue
    Case 8
        Height = vNewValue
    Case 9
        Materials = vNewValue
            
End Select
StatsUpdate
End Sub


Public Sub StatsUpdate()
Dim sRotation As String

mvarVolume = 0 ' reset it back to zero
    ' Calculate the volume for the Mast using the Mast specific formula
    mvarVolume = (mvarHeight ^ 3) / 10000
    
    ' calculate the surface
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    ' Calculate the cost using special Mast and Open Mount rules
    mvarCost = mvarSurfaceArea * 10
    ' Calculate the weight using special Mast and Open Mount rules
    mvarWeight = mvarSurfaceArea * TechLevelModifier(mvarTL)
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

