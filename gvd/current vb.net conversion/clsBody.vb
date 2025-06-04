Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsBody_NET.clsBody")> Public Class clsBody
	
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
	Private mvarDR As Integer
	Private mvarEmptySpace As Single
	Private mvarAccessSpace As Single
	Private mvarBattlesuitVolumeAdded As Boolean
	
	'top deck features
	Private mvarTopDeck As Boolean 'todo: topdeck is to be moved into a seperate object!  This is so obvious to me now!
	Private mvarPercentCovered As Short
	Private mvarPercentFlightDeck As Short
	Private mvarFlightDeckOption As String
	Private mvarTotalDeckArea As Single
	Private mvarCoveredDeckArea As Single
	Private mvarFlightDeckArea As Single
	Private mvarDeckCost As Single
	Private mvarDeckWeight As Single
	Private mvarFlightDeckLength As Single
	
	
	'Generic properties for the class
	Private mvarImage As Short
	Private mvarSelectedImage As Short
	Private mvarParent As String
	Private mvarKey As String
	Private mvarComponent As String
	Private mvarTL As Short
	Private mvarDatatype As Short
	Private mvarParentDatatype As Short
	Private mvarDescription As String
	Private mvarCustomDescription As String
	Private mvarCustom As Boolean
	Private mvarQuantity As Integer
	
	
	
	Private mvarLocation As String
	Private mvarComment As String
	Private mvarCName As String
	Private mvarPrintOutput As String
	Private mvarZZInit As Byte
	Private mvarAbbrev As String
	
	Private mvarLogicalParent As String
	
	
	Public Property LogicalParent() As String
		Get
			LogicalParent = mvarLogicalParent
		End Get
		Set(ByVal Value As String)
			mvarLogicalParent = Value
		End Set
	End Property
	
	
	Public Property Abbrev() As String
		Get
			Abbrev = mvarAbbrev
		End Get
		Set(ByVal Value As String)
			mvarAbbrev = Value
		End Set
	End Property
	
	
	
	
	
	Public Property PrintOutput() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PrintOutput
			PrintOutput = mvarPrintOutput
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PrintOutput = 5
			mvarPrintOutput = Value
		End Set
	End Property
	
	
	
	
	Public Property CName() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.CName
			CName = mvarCName
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.CName = 5
			mvarCName = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Comment() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Comment
			Comment = mvarComment
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Comment = 5
			mvarComment = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Location() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Location
			Location = mvarLocation
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Location = 5
			mvarLocation = Value
		End Set
	End Property
	
	
	
	Public Property MinimumVolume() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.MinimumVolume
			MinimumVolume = mvarMinimumVolume
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.MinimumVolume = 5
			mvarMinimumVolume = Value
		End Set
	End Property
	
	
	
	
	Public Property BattleSuitVolumeAdded() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.BattleSuitVolumeAdded
			BattleSuitVolumeAdded = mvarBattlesuitVolumeAdded
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.BattleSuitVolumeAdded = 5
			mvarBattlesuitVolumeAdded = Value
		End Set
	End Property
	
	
	
	
	
	Public Property TopDeck() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.TopDeck
			TopDeck = mvarTopDeck
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.TopDeck = 5
			mvarTopDeck = Value
		End Set
	End Property
	
	
	
	Public Property PercentFlightDeck() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PercentFlightDeck
			PercentFlightDeck = mvarPercentFlightDeck
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PercentFlightDeck = 5
			mvarPercentFlightDeck = Value
		End Set
	End Property
	
	
	
	Public Property PercentCovered() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PercentCovered
			PercentCovered = mvarPercentCovered
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PercentCovered = 5
			If Value > 100 Then Value = 100
			mvarPercentCovered = Value
		End Set
	End Property
	
	
	
	Public Property FlightDeckOption() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.FlightDeckOption
			FlightDeckOption = mvarFlightDeckOption
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.FlightDeckOption = 5
			
			mvarFlightDeckOption = Value
		End Set
	End Property
	
	
	
	Public Property TotalDeckArea() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.TotalDeckArea
			TotalDeckArea = mvarTotalDeckArea
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.TotalDeckArea = 5
			mvarTotalDeckArea = Value
		End Set
	End Property
	
	
	
	Public Property CoveredDeckArea() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.CoveredDeckArea
			CoveredDeckArea = mvarCoveredDeckArea
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.CoveredDeckArea = 5
			mvarCoveredDeckArea = Value
		End Set
	End Property
	
	
	
	Public Property FlightDeckArea() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.FlightDeckArea
			FlightDeckArea = mvarFlightDeckArea
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.FlightDeckArea = 5
			mvarFlightDeckArea = Value
		End Set
	End Property
	
	
	
	Public Property DeckCost() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.DeckCost
			DeckCost = mvarDeckCost
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.DeckCost = 5
			mvarDeckCost = Value
		End Set
	End Property
	
	
	
	Public Property DeckWeight() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.DeckWeight
			DeckWeight = mvarDeckWeight
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.DeckWeight = 5
			mvarDeckWeight = Value
		End Set
	End Property
	
	
	
	
	Public Property FlightDeckLength() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.mvarFlightDeckLength
			FlightDeckLength = mvarFlightDeckLength
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.mvarFlightDeckLength = 5
			mvarFlightDeckLength = Value
		End Set
	End Property
	
	
	
	
	
	
	
	Public Property Quantity() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Quantity
			Quantity = mvarQuantity
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Quantity = 5
			mvarQuantity = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Custom() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Custom
			Custom = mvarCustom
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Custom = 5
			mvarCustom = Value
		End Set
	End Property
	
	
	
	
	
	Public Property CustomDescription() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.CustomDescription
			CustomDescription = mvarCustomDescription
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.CustomDescription = 5
			mvarCustomDescription = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Description() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Description
			Description = mvarDescription
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Description = 5
			mvarDescription = Value
		End Set
	End Property
	
	
	
	
	Public Property ParentDatatype() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.ParentDatatype
			ParentDatatype = mvarParentDatatype
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.ParentDatatype = 5
			mvarParentDatatype = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Datatype() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Datatype
			Datatype = mvarDatatype
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Datatype = 5
			mvarDatatype = Value
		End Set
	End Property
	
	
	
	
	
	Public Property EmptySpace() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.EmptySpace
			EmptySpace = mvarEmptySpace
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.EmptySpace = 5
			mvarEmptySpace = Value
		End Set
	End Property
	
	
	Public Property AccessSpace() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.AccessSpace
			AccessSpace = mvarAccessSpace
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.AccessSpace = 5
			mvarAccessSpace = Value
		End Set
	End Property
	
	
	
	
	Public Property DR() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.DR
			DR = mvarDR
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.DR = 5
			mvarDR = Value
		End Set
	End Property
	
	
	
	
	
	Public Property HitPoints() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.HitPoints
			HitPoints = mvarHitPoints
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.HitPoints = 5
			mvarHitPoints = Value
		End Set
	End Property
	
	
	
	
	
	
	Public Property Biomechanical() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Biomechanical
			Biomechanical = mvarBiomechanical
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Biomechanical = 5
			mvarBiomechanical = Value
		End Set
	End Property
	
	
	
	
	
	Public Property FrameStrength() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.FrameStrength
			FrameStrength = mvarFrameStrength
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.FrameStrength = 5
			mvarFrameStrength = Value
		End Set
	End Property
	
	
	
	
	
	Public Property LivingMetal() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.LivingMetal
			LivingMetal = mvarLivingMetal
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.LivingMetal = 5
			mvarLivingMetal = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Materials() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Materials
			Materials = mvarMaterials
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Materials = 5
			mvarMaterials = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Responsive() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Responsive
			Responsive = mvarResponsive
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Responsive = 5
			mvarResponsive = Value
			
			
		End Set
	End Property
	
	
	
	
	
	Public Property Robotic() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Robotic
			Robotic = mvarRobotic
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Robotic = 5
			mvarRobotic = Value
		End Set
	End Property
	
	
	
	
	
	
	Public Property component() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Component
			component = mvarComponent
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Component = 5
			mvarComponent = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Volume() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Volume
			Volume = mvarVolume
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Volume = 5
			mvarVolume = Value
		End Set
	End Property
	
	
	
	
	
	
	Public Property Weight() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Weight
			Weight = mvarWeight
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Weight = 5
			mvarWeight = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Cost() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Cost
			Cost = mvarCost
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Cost = 5
			mvarCost = Value
		End Set
	End Property
	
	
	
	Public Property ImprovedSuspensionCost() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.ImprovedSuspensionCost
			ImprovedSuspensionCost = mvarImprovedSuspensionCost
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.ImprovedSuspensionCost = 5
			mvarImprovedSuspensionCost = Value
		End Set
	End Property
	
	
	
	Public Property compartmentalizationcost() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.CompartmentalizationCost
			compartmentalizationcost = mvarCompartmentalizationCost
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.CompartmentalizationCost = 5
			mvarCompartmentalizationCost = Value
		End Set
	End Property
	
	
	
	Public Property compartmentalizationWeight() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.CompartmentalizationWeight
			compartmentalizationWeight = mvarCompartmentalizationWeight
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.CompartmentalizationWeight = 5
			mvarCompartmentalizationWeight = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Compartmentalization() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Compartmentalization
			Compartmentalization = mvarCompartmentalization
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Compartmentalization = 5
			mvarCompartmentalization = Value
		End Set
	End Property
	
	
	
	
	
	
	
	Public Property SlopeB() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.SlopeB
			SlopeB = mvarSlopeB
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.SlopeB = 5
			mvarSlopeB = Value
		End Set
	End Property
	
	
	
	
	
	Public Property SlopeF() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.SlopeF
			SlopeF = mvarSlopeF
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.SlopeF = 5
			mvarSlopeF = Value
		End Set
	End Property
	
	
	
	
	
	Public Property SlopeL() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.SlopeL
			SlopeL = mvarSlopeL
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.SlopeL = 5
			mvarSlopeL = Value
		End Set
	End Property
	
	
	
	
	
	Public Property SlopeR() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.SlopeR
			SlopeR = mvarSlopeR
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.SlopeR = 5
			mvarSlopeR = Value
		End Set
	End Property
	
	
	
	
	Public Property SurfaceArea() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.SurfaceArea
			SurfaceArea = mvarSurfaceArea
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.SurfaceArea = 5
			mvarSurfaceArea = Value
		End Set
	End Property
	
	
	
	
	
	Public Property FlexibodyOption() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.FlexibodyOption
			FlexibodyOption = mvarFlexibodyOption
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.FlexibodyOption = 5
			mvarFlexibodyOption = Value
		End Set
	End Property
	
	
	
	Public Property ImprovedSuspension() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.ImprovedSuspension
			ImprovedSuspension = mvarImprovedSuspension
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.ImprovedSuspension = 5
			mvarImprovedSuspension = Value
		End Set
	End Property
	
	
	
	Public Property LiftingBody() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.LiftingBody
			LiftingBody = mvarLiftingBody
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.LiftingBody = 5
			mvarLiftingBody = Value
		End Set
	End Property
	
	
	
	
	
	Public Property TL() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.TL
			TL = mvarTL
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.TL = 5
			mvarTL = Value
			gVehicleTL = mvarTL
		End Set
	End Property
	
	
	
	
	Public Property Key() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Key
			Key = mvarKey
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Key = 5
			mvarKey = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Parent() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Parent
			Parent = mvarParent
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Parent = 5
			mvarParent = Value
		End Set
	End Property
	
	
	
	
	
	Public Property SelectedImage() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.SelectedImage
			SelectedImage = mvarSelectedImage
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.SelectedImage = 5
			mvarSelectedImage = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Image() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Image
			Image = mvarImage
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Image = 5
			mvarImage = Value
		End Set
	End Property
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'set the default dimension of the keychain to 1 element
		
		Dim mvarPowerConsumptionKeyChain(1) As Object
		Dim mvarPerformanceProfileKeychain(1) As Object
		Dim mvarLegsKeychain(1) As Object
		Dim mvarLegDrivetrainKeychain(1) As Object
		Dim mvarRotorDrivetrainKeychain(1) As Object
		Dim mvarOrnithopterDrivetrainKeychain(1) As Object
		Dim mvarOtherGroundDrivetrainKeychain(1) As Object
		Dim mvarRotorsKeychain(1) As Object
		Dim mvarWeaponLinkKeychain(1) As Object
		Dim mvarSubAssembliesKeychain(1) As Object
		
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
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
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
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarVolume = mvarVolume + Veh.Components(BATTLESUIT_KEY_SYSTEM).Volume1
		End If
		
		' Add any turret rotationspace if it exists
		mvarVolume = mvarAccessSpace + mvarVolume + CalcRotationSpace(mvarKey)
		
		' calculate the real body volume
		mvarVolume = System.Math.Round(CalcBodyVolume, 2)
		
		
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
			mvarTotalDeckArea = System.Math.Round(CalcTotalDeckArea(mvarSurfaceArea, mvarKey), 2)
			mvarFlightDeckArea = System.Math.Round(mvarTotalDeckArea * (mvarPercentFlightDeck / 100), 2)
			mvarFlightDeckLength = System.Math.Round(3 * System.Math.Sqrt(mvarFlightDeckArea), 2)
			mvarCoveredDeckArea = System.Math.Round(mvarTotalDeckArea * (mvarPercentCovered / 100), 2)
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
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		mvarHitPoints = CalcHitPoints(TypeName(Me), mvarFrameStrength, mvarSurfaceArea)
		
		'generate print
		mvarPrintOutput = mvarCustomDescription & "."
		
	End Sub
	
	
	
	Private Function CalcBodyVolume() As Single
		
		Dim SlopeMultiplier As Single ' vehicles total slope Volume multiplier
		Dim RetractsCost As Single ' Rectractable Wheels or Skids Volume Multiplier
		Dim StreamLiningCost As Single ' Streamlined Hull Volume Multiplier
		Dim HydrodynamicHullCost As Single ' total of all Special Structure Cost Modifiers
		Dim TotalOtherCost As Single ' total of all Other Volume Modifiers
		Dim sRetracts As String ' vehicle body's rectract property
		
		Const CatorTrimaranCost As Double = 1.3 ' Catamaran or Trimaran Hull Volume Multiplier
		Const SubmersibleCost As Double = 1.25 ' Submersible Hull Volume Multiplier
		
		'get the retract location
		sRetracts = GetRetractLocation
		
		' Calculate the Real Body Volume
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Veh.surface.Submersible Then TotalOtherCost = SubmersibleCost
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Veh.surface.CataTrimaran <> modConstants.EMISSION_CLOAKING.NONE Then TotalOtherCost = TotalOtherCost * CatorTrimaranCost
		' Calculate Final Volume for the Body
		CalcBodyVolume = mvarVolume * TotalOtherCost * StreamLiningCost * HydrodynamicHullCost * SlopeMultiplier * RetractsCost
	End Function
	
	Public Sub CalcCompartmentalizationStats()
		'this routine must be called AFTER the total vehicles
		'structural surface area is known.  It then can
		'compute the cost and weight associated with compartmentalizing
		'this subassembly
		Dim Divisor As Double
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Divisor = Veh.Stats.StructuralSurfaceArea
		
		If (mvarCompartmentalization <> "none") And (Divisor > 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
End Class