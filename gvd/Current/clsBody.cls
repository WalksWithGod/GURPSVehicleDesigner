VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'////////////////////////////////////////////////
'Standard SubAssembly Properties
'///////////////////////////////////////////////

Private mvarSlopeR As String
Private mvarSlopeL As String
Private mvarSlopeF As String
Private mvarSlopeB As String
Private mvarRobotic As Boolean
Private mvarResponsive As Boolean
Private mvarMaterials As String
Private mvarLivingMetal As Boolean
Private mvarFrameStrength As String
Private mvarBiomechanical As Boolean
Private mvarFlexibodyOption As Boolean
Private mvarLiftingBody As Boolean
Private mvarImprovedSuspension As Boolean
Private mvarImprovedSuspensionCost As Single
Private mvarCompartmentalization As String
Private mvarCompartmentalizationWeight As Single
Private mvarCompartmentalizationCost As Single

'//////////////////////////////////////////////
' Body Stats
'/////////////////////////////////////////////
Private mvarMinimumVolume As Single
Private mvarVolume As Double
Private mvarSurfaceArea As Double
Private mvarCost As Double
Private mvarWeight As Double
Private mvarHitPoints As Single
Private mvarDR As Long
Private mvarEmptySpace As Single
Private mvarAccessSpace As Single
Private mvarBattlesuitVolumeAdded As Boolean

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


'Generic properties for the class
Private mvarImage As Integer
Private mvarSelectedImage As Integer
Private mvarParent As String
Private mvarKey As String
Private mvarComponent As String
Private mvarTL As Integer
Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarQuantity As Long



Private mvarLocation As String
Private mvarComment As String
Private mvarCName As String
Private mvarPrintOutput As String
Private mvarZZInit As Byte
Private mvarAbbrev As String

Private mvarLogicalParent As String

Public Property Let LogicalParent(ByVal vdata As String)
    mvarLogicalParent = vdata
End Property

Public Property Get LogicalParent() As String
    LogicalParent = mvarLogicalParent
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

Public Property Let MinimumVolume(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MinimumVolume = 5
    mvarMinimumVolume = vdata
End Property


Public Property Get MinimumVolume() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinimumVolume
    MinimumVolume = mvarMinimumVolume
End Property


Public Property Let BattleSuitVolumeAdded(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BattleSuitVolumeAdded = 5
    mvarBattlesuitVolumeAdded = vdata
End Property


Public Property Get BattleSuitVolumeAdded() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BattleSuitVolumeAdded
    BattleSuitVolumeAdded = mvarBattlesuitVolumeAdded
End Property



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





Public Property Let Quantity(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quantity = 5
    mvarQuantity = vdata
End Property


Public Property Get Quantity() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quantity
    Quantity = mvarQuantity
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
Public Property Let AccessSpace(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccessSpace = 5
    mvarAccessSpace = vdata
End Property


Public Property Get AccessSpace() As Double
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




Public Property Let component(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Component = 5
    mvarComponent = vdata
End Property


Public Property Get component() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Component
    component = mvarComponent
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

Public Property Let ImprovedSuspensionCost(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ImprovedSuspensionCost = 5
    mvarImprovedSuspensionCost = vdata
End Property


Public Property Get ImprovedSuspensionCost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ImprovedSuspensionCost
    ImprovedSuspensionCost = mvarImprovedSuspensionCost
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



Public Property Let FlexibodyOption(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FlexibodyOption = 5
    mvarFlexibodyOption = vdata
End Property


Public Property Get FlexibodyOption() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FlexibodyOption
    FlexibodyOption = mvarFlexibodyOption
End Property

Public Property Let ImprovedSuspension(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ImprovedSuspension = 5
    mvarImprovedSuspension = vdata
End Property


Public Property Get ImprovedSuspension() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ImprovedSuspension
    ImprovedSuspension = mvarImprovedSuspension
End Property

Public Property Let LiftingBody(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LiftingBody = 5
    mvarLiftingBody = vdata
End Property


Public Property Get LiftingBody() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LiftingBody
    LiftingBody = mvarLiftingBody
End Property



Public Property Let TL(ByVal vdata As Integer)
Attribute TL.VB_Description = "Is the age that the vehicle was built.  See page 6 Vehicles bottom first column"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TL = 5
    mvarTL = vdata
    gVehicleTL = mvarTL
End Property


Public Property Get TL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    TL = mvarTL
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
'set the default dimension of the keychain to 1 element
 
ReDim mvarPowerConsumptionKeyChain(1)
ReDim mvarPerformanceProfileKeychain(1)
ReDim mvarLegsKeychain(1)
ReDim mvarLegDrivetrainKeychain(1)
ReDim mvarRotorDrivetrainKeychain(1)
ReDim mvarOrnithopterDrivetrainKeychain(1)
ReDim mvarOtherGroundDrivetrainKeychain(1)
ReDim mvarRotorsKeychain(1)
ReDim mvarWeaponLinkKeychain(1)
ReDim mvarSubAssembliesKeychain(1)

' set the default property values

TL = 7 'note: this should change the gvehicletl global variable


mvarCompartmentalization = "none"

mvarLiftingBody = False
'mvarImprovedSuspension = False

'mvarTopDeckArea = 0
'mvarFlightDeckArea = 0
'mvarCoveredDeckArea = 0
'mvarLandingPad = False
'mvarAngledFlightDeck = False
mvarFlightDeckOption = "none"
mvarFlexibodyOption = False
mvarSlopeR = "none"
mvarSlopeL = "none"
mvarSlopeF = "none"
mvarSlopeB = "none"
mvarSurfaceArea = 0
mvarMinimumVolume = 0
mvarVolume = 0
mvarCost = 0
mvarWeight = 0
mvarHitPoints = 0
mvarRobotic = False
mvarResponsive = False
mvarMaterials = "standard"
mvarLivingMetal = False
mvarFrameStrength = "medium"
mvarBiomechanical = False


mvarAccessSpace = 0
mvarEmptySpace = 0


End Sub

Public Sub StatsUpdate()


mvarZZInit = 1

mvarAbbrev = "Bo"

'get the accessspace
mvarAccessSpace = CalcAccessSpace(mvarKey)

' Calculate the combined component volume
mvarVolume = CalcCombinedVolume(mvarKey) + mvarEmptySpace

' add volume for Battlesuit system if applicable
If mvarBattlesuitVolumeAdded Then
    mvarVolume = mvarVolume + Veh.Components(BATTLESUIT_KEY_SYSTEM).Volume1
End If

' Add any turret rotationspace if it exists
mvarVolume = mvarAccessSpace + mvarVolume + CalcRotationSpace(mvarKey)

' calculate the real body volume
mvarVolume = Round(CalcBodyVolume, 2)


' calculate the surface area
mvarSurfaceArea = CalcSurfaceArea(mvarVolume)

If TopDeck = False Then
    mvarTotalDeckArea = 0
    mvarFlightDeckArea = 0
    mvarFlightDeckLength = 0
    mvarCoveredDeckArea = 0
    mvarDeckCost = 0
    mvarDeckWeight = 0
Else
    ' calculate the TopDeck surfacearea, cost and weight
    mvarTotalDeckArea = Round(CalcTotalDeckArea(mvarSurfaceArea, mvarKey), 2)
    mvarFlightDeckArea = Round(mvarTotalDeckArea * (mvarPercentFlightDeck / 100), 2)
    mvarFlightDeckLength = Round(3 * Sqr(mvarFlightDeckArea), 2)
    mvarCoveredDeckArea = Round(mvarTotalDeckArea * (mvarPercentCovered / 100), 2)
    mvarDeckCost = CalcDeckCost(mvarFlightDeckArea, mvarCoveredDeckArea, mvarFlightDeckOption)
    mvarDeckWeight = CalcDeckWeight(mvarFlightDeckArea, mvarCoveredDeckArea, mvarFlightDeckOption)
    
End If

'now that we have surface area, get the cost of Improved suspension
If mvarImprovedSuspension Then mvarImprovedSuspensionCost = mvarSurfaceArea * 50 Else mvarImprovedSuspensionCost = 0

' Calculate the structural cost
mvarCost = mvarDeckCost + BasicDesignCost(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials, mvarResponsive, mvarRobotic, mvarBiomechanical, mvarLivingMetal)
' Calculate the structural weight
mvarWeight = mvarDeckWeight + BasicDesignWeight(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials)
' Calculate the Hit Points
mvarHitPoints = CalcHitPoints(TypeName(Me), mvarFrameStrength, mvarSurfaceArea)

'generate print
mvarPrintOutput = mvarCustomDescription + "."

End Sub



Private Function CalcBodyVolume() As Single

Dim SlopeMultiplier As Single ' vehicles total slope Volume multiplier
Dim RetractsCost As Single ' Rectractable Wheels or Skids Volume Multiplier
Dim StreamLiningCost As Single ' Streamlined Hull Volume Multiplier
Dim HydrodynamicHullCost As Single ' total of all Special Structure Cost Modifiers
Dim TotalOtherCost As Single ' total of all Other Volume Modifiers
Dim sRetracts As String ' vehicle body's rectract property

Const CatorTrimaranCost = 1.3 ' Catamaran or Trimaran Hull Volume Multiplier
Const SubmersibleCost = 1.25 ' Submersible Hull Volume Multiplier
   
'get the retract location
sRetracts = GetRetractLocation

' Calculate the Real Body Volume
Select Case Veh.surface.StreamLining
    Case "none"
        StreamLiningCost = 1
    Case "fair"
        StreamLiningCost = 1.1
    Case "good"
        StreamLiningCost = 1.2
    Case "very good"
        StreamLiningCost = 1.25
    Case "superior"
        StreamLiningCost = 1.3
    Case "excellent"
        StreamLiningCost = 1.35
    Case "radical"
        StreamLiningCost = 1.4
End Select

Select Case Veh.surface.HydrodynamicLines
    Case "none"
        HydrodynamicHullCost = 1
    Case "mediocre"
        HydrodynamicHullCost = 1.1
    Case "average"
        HydrodynamicHullCost = 1.2
    Case "submarine"
        HydrodynamicHullCost = 1.2
    Case "fine"
        HydrodynamicHullCost = 1.3
    Case "very fine"
        HydrodynamicHullCost = 1.3
End Select

Select Case sRetracts
    Case "none"
        RetractsCost = 1
    Case "body"
        RetractsCost = 1.075
    Case "body & wings"
        RetractsCost = 1.025
End Select

' calculate the Multiplier needed for the slopes applied to the Assembly
SlopeMultiplier = CalcSlopeMultiplier(mvarKey)
' Calculate Volume modifier value for Other Modifiers
TotalOtherCost = 1 ' initialize variable

If Veh.surface.Submersible Then TotalOtherCost = SubmersibleCost
If Veh.surface.CataTrimaran <> NONE Then TotalOtherCost = TotalOtherCost * CatorTrimaranCost
' Calculate Final Volume for the Body
CalcBodyVolume = mvarVolume * TotalOtherCost * StreamLiningCost * HydrodynamicHullCost * SlopeMultiplier * RetractsCost
End Function

Public Sub CalcCompartmentalizationStats()
'this routine must be called AFTER the total vehicles
'structural surface area is known.  It then can
'compute the cost and weight associated with compartmentalizing
'this subassembly
Dim Divisor As Double
Divisor = Veh.Stats.StructuralSurfaceArea

If (mvarCompartmentalization <> "none") And (Divisor > 0) Then
    mvarCompartmentalizationWeight = (0.1 * Veh.Stats.StructuralWeight / Divisor) * mvarSurfaceArea
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

Public Sub QueryChild()
' if the object has children, query it and check to see if
' more stats/property updates are needed for other objects in the collection
' (see the StatusUpdate property for help on checking for childeren.  Can i use that one in place of this?)

End Sub
