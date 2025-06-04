VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSuperStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit
Private mvarOrientation As String
Private mvarImage As Integer
Private mvarKey As String
Private mvarParent As String
Private mvarSelectedImage As Integer
Private mvarTL As Integer
Private mvarCompartmentalizationWeight As Single
Private mvarCompartmentalizationCost As Single

Private mvarCost As Double
Private mvarWeight As Double
Private mvarFrameStrength As String
Private mvarMaterials As String
Private mvarResponsive As Boolean
Private mvarRobotic As Boolean
Private mvarBiomechanical As Boolean
Private mvarLivingMetal As Boolean
Private mvarSurfaceArea As Double
Private mvarVolume As Double
Private mvarSlopeL As String
Private mvarSlopeR As String
Private mvarSlopeF As String
Private mvarSlopeB As String
Private mvarCompartmentalization As String
Private mvarEmptySpace As Single
Private mvarHitPoints As Double
Private mvarLocation As String
Private mvarDR As Long
Private mvarAccessSpace As Single
Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean

'top deck features
Private mvarTopDeck As Boolean 'todo: topdeck is to be moved into a seperate object!  This is so obvious to me now!
Private mvarPercentCovered As Integer
Private mvarPercentFlightDeck As Integer
Private mvarFlightDeckOption As String
Private mvarTotalDeckArea As Single
Private mvarCoveredDeckArea As Single
Private mvarFlightDeckArea As Single
Private mvarDeckCost As Single
Private mvarDeckWeight As Single
Private mvarFlightDeckLength As Single

Private mvarComment As String
Private mvarCName As String
Private mvarPrintOutput As String
Private mvarZZInit As Byte
Private mvarAbbrev As String
Private mvarIndex As Long
Private mvarLogicalParent As String

Public Property Let LogicalParent(ByVal vdata As String)
    mvarLogicalParent = vdata
End Property

Public Property Get LogicalParent() As String
    LogicalParent = mvarLogicalParent
End Property

Public Property Let index(ByVal vdata As Long)
    mvarIndex = vdata
End Property

Public Property Get index() As Long
    index = mvarIndex
End Property

Public Property Let Abbrev(ByVal vdata As String)
    mvarAbbrev = vdata
End Property

Public Property Get Abbrev() As String
    Abbrev = mvarAbbrev
End Property

Public Property Let PrintOutput(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrintOutput = 5
    mvarPrintOutput = vdata
End Property


Public Property Get PrintOutput() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrintOutput
    PrintOutput = mvarPrintOutput
End Property



Public Property Let Orientation(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Orientation = 5
  
  mvarOrientation = vdata

End Property


Public Property Get Orientation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Orientation
    Orientation = mvarOrientation
End Property

Public Property Let CName(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CName = 5
    mvarCName = vdata
End Property


Public Property Get CName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CName
    CName = mvarCName
End Property



Public Property Let Comment(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Comment = 5
    mvarComment = vdata
End Property


Public Property Get Comment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Comment
    Comment = mvarComment
End Property


Public Property Let compartmentalizationcost(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CompartmentalizationCost = 5
    mvarCompartmentalizationCost = vdata
End Property


Public Property Get compartmentalizationcost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CompartmentalizationCost
    compartmentalizationcost = mvarCompartmentalizationCost
End Property

Public Property Let compartmentalizationWeight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CompartmentalizationWeight = 5
    mvarCompartmentalizationWeight = vdata
End Property


Public Property Get compartmentalizationWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CompartmentalizationWeight
    compartmentalizationWeight = mvarCompartmentalizationWeight
End Property



Public Function LocationCheck() As Boolean
Dim TempCheck As Boolean
Dim InstallPoint As Integer
InstallPoint = Veh.Components(mvarParent).Datatype

        If (InstallPoint = Body) Or (InstallPoint = Turret) Then
            TempCheck = True
        Else
            InfoPrint 1, "Superstructures must be placed on Body or Turret."
            TempCheck = False
        End If

If TempCheck Then SetLogicalParent
LocationCheck = TempCheck
End Function


Private Function GetLocation() As String
    On Error Resume Next
    If mvarLogicalParent = "" Then SetLogicalParent
    GetLocation = Veh.Components(mvarLogicalParent).Abbrev
    
End Function

Public Sub SetLogicalParent()
    mvarLogicalParent = GetLogicalParent(mvarParent)
End Sub


Public Property Let TopDeck(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TopDeck = 5
    mvarTopDeck = vdata
End Property


Public Property Get TopDeck() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TopDeck
    TopDeck = mvarTopDeck
End Property

Public Property Let PercentFlightDeck(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PercentFlightDeck = 5
    If vdata > 100 Then vdata = 100
    
    mvarPercentFlightDeck = vdata
End Property


Public Property Get PercentFlightDeck() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PercentFlightDeck
    PercentFlightDeck = mvarPercentFlightDeck
End Property

Public Property Let PercentCovered(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PercentCovered = 5
    If vdata > 100 Then vdata = 100
    mvarPercentCovered = vdata
End Property


Public Property Get PercentCovered() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PercentCovered
    PercentCovered = mvarPercentCovered
End Property

Public Property Let FlightDeckOption(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FlightDeckOption = 5
    mvarFlightDeckOption = vdata
End Property


Public Property Get FlightDeckOption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FlightDeckOption
    FlightDeckOption = mvarFlightDeckOption
End Property

Public Property Let TotalDeckArea(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalDeckArea = 5
    mvarTotalDeckArea = vdata
End Property


Public Property Get TotalDeckArea() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalDeckArea
    TotalDeckArea = mvarTotalDeckArea
End Property

Public Property Let CoveredDeckArea(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CoveredDeckArea = 5
    mvarCoveredDeckArea = vdata
End Property


Public Property Get CoveredDeckArea() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CoveredDeckArea
    CoveredDeckArea = mvarCoveredDeckArea
End Property

Public Property Let FlightDeckArea(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FlightDeckArea = 5
    mvarFlightDeckArea = vdata
End Property


Public Property Get FlightDeckArea() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FlightDeckArea
    FlightDeckArea = mvarFlightDeckArea
End Property

Public Property Let DeckCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DeckCost = 5
    mvarDeckCost = vdata
End Property


Public Property Get DeckCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeckCost
    DeckCost = mvarDeckCost
End Property

Public Property Let DeckWeight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DeckWeight = 5
    mvarDeckWeight = vdata
End Property


Public Property Get DeckWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeckWeight
    DeckWeight = mvarDeckWeight
End Property


Public Property Let FlightDeckLength(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mvarFlightDeckLength = 5
    mvarFlightDeckLength = vdata
End Property


Public Property Get FlightDeckLength() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mvarFlightDeckLength
    FlightDeckLength = mvarFlightDeckLength
End Property
Public Property Let Custom(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Custom = 5
    mvarCustom = vdata
End Property


Public Property Get Custom() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Custom
    Custom = mvarCustom
End Property



Public Property Let CustomDescription(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CustomDescription = 5
    mvarCustomDescription = vdata
End Property


Public Property Get CustomDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CustomDescription
    CustomDescription = mvarCustomDescription
End Property



Public Property Let Description(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vdata
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property


Public Property Let ParentDatatype(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ParentDatatype = 5
    mvarParentDatatype = vdata
End Property


Public Property Get ParentDatatype() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ParentDatatype
    ParentDatatype = mvarParentDatatype
End Property



Public Property Let Datatype(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Datatype = 5
    mvarDatatype = vdata
End Property


Public Property Get Datatype() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Datatype
    Datatype = mvarDatatype
End Property



Public Property Let AccessSpace(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccessSpace = 5
    mvarAccessSpace = vdata
End Property


Public Property Get AccessSpace() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccessSpace
    AccessSpace = mvarAccessSpace
End Property



Public Property Let DR(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR = 5
    mvarDR = vdata
End Property



Public Property Get DR() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DR
     DR = mvarDR
End Property



Public Property Let Location(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Location = 5
    mvarLocation = vdata
End Property


Public Property Get Location() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Location
    Location = mvarLocation
End Property



Public Property Let HitPoints(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HitPoints = 5
    mvarHitPoints = vdata
End Property


Public Property Get HitPoints() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HitPoints
    HitPoints = mvarHitPoints
End Property



Public Property Let EmptySpace(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmptySpace = 5
    mvarEmptySpace = vdata
End Property


Public Property Get EmptySpace() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmptySpace
    EmptySpace = mvarEmptySpace
End Property



Public Property Let Compartmentalization(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Compartmentalization = 5
    mvarCompartmentalization = vdata
End Property


Public Property Get Compartmentalization() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Compartmentalization
    Compartmentalization = mvarCompartmentalization
End Property



Public Property Let SlopeB(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlopeB = 5
    mvarSlopeB = vdata
End Property


Public Property Get SlopeB() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlopeB
    SlopeB = mvarSlopeB
End Property



Public Property Let SlopeF(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlopeF = 5
    mvarSlopeF = vdata
End Property


Public Property Get SlopeF() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlopeF
    SlopeF = mvarSlopeF
End Property



Public Property Let SlopeR(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlopeR = 5
    mvarSlopeR = vdata
End Property


Public Property Get SlopeR() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlopeR
    SlopeR = mvarSlopeR
End Property



Public Property Let SlopeL(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlopeL = 5
    mvarSlopeL = vdata
End Property


Public Property Get SlopeL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlopeL
    SlopeL = mvarSlopeL
End Property



Public Property Let Volume(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Volume = 5
    mvarVolume = vdata
End Property


Public Property Get Volume() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Volume
    Volume = mvarVolume
End Property



Public Property Let SurfaceArea(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SurfaceArea = 5
    mvarSurfaceArea = vdata
End Property


Public Property Get SurfaceArea() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SurfaceArea
    SurfaceArea = mvarSurfaceArea
End Property



Public Property Let LivingMetal(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LivingMetal = 5
    mvarLivingMetal = vdata
End Property


Public Property Get LivingMetal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LivingMetal
    LivingMetal = mvarLivingMetal
End Property



Public Property Let Biomechanical(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Biomechanical = 5
    mvarBiomechanical = vdata
End Property


Public Property Get Biomechanical() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Biomechanical
    Biomechanical = mvarBiomechanical
End Property



Public Property Let Robotic(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Robotic = 5
    mvarRobotic = vdata
End Property


Public Property Get Robotic() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Robotic
    Robotic = mvarRobotic
End Property



Public Property Let Responsive(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Responsive = 5
    mvarResponsive = vdata
End Property


Public Property Get Responsive() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Responsive
    Responsive = mvarResponsive
End Property



Public Property Let Materials(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Materials = 5
    mvarMaterials = vdata
End Property


Public Property Get Materials() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Materials
    Materials = mvarMaterials
End Property



Public Property Let FrameStrength(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FrameStrength = 5
    mvarFrameStrength = vdata
End Property


Public Property Get FrameStrength() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FrameStrength
    FrameStrength = mvarFrameStrength
End Property



Public Property Let Weight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Weight = 5
    mvarWeight = vdata
End Property


Public Property Get Weight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Weight
    Weight = mvarWeight
End Property



Public Property Let Cost(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cost = 5
    mvarCost = vdata
End Property


Public Property Get Cost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cost
    Cost = mvarCost
End Property



Public Property Let TL(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TL = 5
    mvarTL = vdata
End Property


Public Property Get TL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    TL = mvarTL
End Property


Public Property Let SelectedImage(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SelectedImage = 5
    mvarSelectedImage = vdata
End Property


Public Property Get SelectedImage() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedImage
    SelectedImage = mvarSelectedImage
End Property



Public Property Let Parent(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Parent = 5
    mvarParent = vdata
End Property


Public Property Get Parent() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parent
    Parent = mvarParent
End Property



Public Property Let Key(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vdata
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property



Public Property Let Image(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Image = 5
    mvarImage = vdata
End Property


Public Property Get Image() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Image
    Image = mvarImage
End Property


Private Sub Class_Initialize()
' set the default properties
mvarTL = gVehicleTL
mvarSlopeR = "none"
mvarSlopeL = "none"
mvarSlopeF = "none"
mvarSlopeB = "none"
mvarCompartmentalization = Veh.Components(BODY_KEY).Compartmentalization
mvarFrameStrength = Veh.Components(BODY_KEY).FrameStrength
mvarMaterials = Veh.Components(BODY_KEY).Materials
mvarResponsive = Veh.Components(BODY_KEY).Responsive
mvarRobotic = Veh.Components(BODY_KEY).Robotic
mvarBiomechanical = Veh.Components(BODY_KEY).Biomechanical
mvarLivingMetal = Veh.Components(BODY_KEY).LivingMetal
mvarCost = 0
mvarWeight = 0
mvarVolume = 0
mvarSurfaceArea = 0
mvarHitPoints = 0
mvarFlightDeckOption = "none"
mvarOrientation = "top"
End Sub


Private Sub Class_Terminate()
    'the class is being destroyed
    
End Sub

Public Sub StatsUpdate()
    mvarZZInit = 1
'NOTE: No quantities for these either
    
    mvarLocation = GetLocation
    
    mvarAbbrev = "Su"
    If mvarIndex > 0 Then mvarAbbrev = mvarAbbrev & mvarIndex
    
    'get the accessspace
    mvarAccessSpace = CalcAccessSpace(mvarKey)
    'Calc volume for any Turrets
    mvarVolume = CalcRotationSpace(mvarKey)
    
    ' Calculate the component volume
    mvarVolume = mvarVolume + CalcCombinedVolume(mvarKey) + mvarEmptySpace + mvarAccessSpace
    ' Calculate the new volume based on slope modifier
    mvarVolume = mvarVolume * CalcSlopeMultiplier(mvarKey)
    ' calculate the surface
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    ' calculate the TopDeck surfacearea, cost and weight
    If mvarTopDeck = False Then
        mvarTotalDeckArea = 0
        mvarFlightDeckArea = 0
        mvarFlightDeckLength = 0
        mvarCoveredDeckArea = 0
        mvarDeckCost = 0
        mvarDeckWeight = 0
    Else
        mvarTotalDeckArea = Round(CalcTotalDeckArea(mvarSurfaceArea, mvarKey), 2)
        mvarFlightDeckArea = Round(mvarTotalDeckArea * (mvarPercentFlightDeck / 100), 2)
        mvarFlightDeckLength = Round(3 * Sqr(mvarFlightDeckArea), 2)
        mvarCoveredDeckArea = Round(mvarTotalDeckArea * (mvarPercentCovered / 100), 2)
        mvarDeckCost = CalcDeckCost(mvarFlightDeckArea, mvarCoveredDeckArea, mvarFlightDeckOption)
        mvarDeckWeight = CalcDeckWeight(mvarFlightDeckArea, mvarCoveredDeckArea, mvarFlightDeckOption)
    End If
    ' Calculate the cost
    mvarCost = mvarDeckCost + BasicDesignCost(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials, mvarResponsive, mvarRobotic, mvarBiomechanical, mvarLivingMetal)
    ' Calculate the weight
    mvarWeight = mvarDeckWeight + BasicDesignWeight(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials)
    ' Calculate the Hit Points
    mvarHitPoints = CalcHitPoints(TypeName(Me), mvarFrameStrength, mvarSurfaceArea)
    
    'generate print output
    mvarPrintOutput = mvarCustomDescription + " (" + mvarOrientation + " of " + Veh.Components(mvarParent).CustomDescription + ")."
End Sub

Public Sub CalcCompartmentalizationStats()
'this routine must be called AFTER the total vehicles
'structural surface area is known.  It then can
'compute the cost and weight associated with compartmentalizing
'this subassembly
Dim nDivisor As Single
nDivisor = Veh.Stats.StructuralSurfaceArea

If (mvarCompartmentalization <> "none") And (nDivisor > 0) Then
    mvarCompartmentalizationWeight = (0.1 * Veh.Stats.StructuralWeight / nDivisor) * mvarSurfaceArea
    If mvarCompartmentalization = "total" Then mvarCompartmentalizationWeight = mvarCompartmentalizationWeight * 2
    If mvarTL <= 6 Then
        mvarCompartmentalizationCost = mvarCompartmentalizationWeight
    Else
        mvarCompartmentalizationCost = 5 * mvarCompartmentalizationWeight
    End If
Else
    mvarCompartmentalizationWeight = 0
    mvarCompartmentalizationCost = 0
End If

End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If

End Sub

Public Sub QueryChild()
' if the object has children, query it and check to see if
' more stats/property updates are needed for other objects in the collection
' (see the StatusUpdate property for help on checking for childeren.  Can i use that one in place of this?)

End Sub

