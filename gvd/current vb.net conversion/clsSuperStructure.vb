Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsSuperStructure_NET.clsSuperStructure")> Public Class clsSuperStructure
	Private mvarOrientation As String
	Private mvarImage As Short
	Private mvarKey As String
	Private mvarParent As String
	Private mvarSelectedImage As Short
	Private mvarTL As Short
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
	Private mvarDR As Integer
	Private mvarAccessSpace As Single
	Private mvarDatatype As Short
	Private mvarParentDatatype As Short
	Private mvarDescription As String
	Private mvarCustomDescription As String
	Private mvarCustom As Boolean
	
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
	
	Private mvarComment As String
	Private mvarCName As String
	Private mvarPrintOutput As String
	Private mvarZZInit As Byte
	Private mvarAbbrev As String
	Private mvarIndex As Integer
	Private mvarLogicalParent As String
	
	
	Public Property LogicalParent() As String
		Get
			LogicalParent = mvarLogicalParent
		End Get
		Set(ByVal Value As String)
			mvarLogicalParent = Value
		End Set
	End Property
	
	
	Public Property index() As Integer
		Get
			index = mvarIndex
		End Get
		Set(ByVal Value As Integer)
			mvarIndex = Value
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
	
	
	
	
	
	Public Property Orientation() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Orientation
			Orientation = mvarOrientation
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Orientation = 5
			
			mvarOrientation = Value
			
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
			If Value > 100 Then Value = 100
			
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
	
	
	
	
	
	Public Property AccessSpace() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.AccessSpace
			AccessSpace = mvarAccessSpace
		End Get
		Set(ByVal Value As Single)
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
	
	
	
	Public Function LocationCheck() As Boolean
		Dim TempCheck As Boolean
		Dim InstallPoint As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallPoint = Veh.Components(mvarParent).Datatype
		
		If (InstallPoint = Body) Or (InstallPoint = Turret) Then
			TempCheck = True
		Else
			modHelper.InfoPrint(1, "Superstructures must be placed on Body or Turret.")
			TempCheck = False
		End If
		
		If TempCheck Then SetLogicalParent()
		LocationCheck = TempCheck
	End Function
	
	
	Private Function GetLocation() As String
		On Error Resume Next
		If mvarLogicalParent = "" Then SetLogicalParent()
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetLocation = Veh.Components(mvarLogicalParent).Abbrev
		
	End Function
	
	Public Sub SetLogicalParent()
		mvarLogicalParent = GetLogicalParent(mvarParent)
	End Sub
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		' set the default properties
		mvarTL = gVehicleTL
		mvarSlopeR = "none"
		mvarSlopeL = "none"
		mvarSlopeF = "none"
		mvarSlopeB = "none"
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarCompartmentalization = Veh.Components(BODY_KEY).Compartmentalization
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarFrameStrength = Veh.Components(BODY_KEY).FrameStrength
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarMaterials = Veh.Components(BODY_KEY).Materials
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarResponsive = Veh.Components(BODY_KEY).Responsive
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarRobotic = Veh.Components(BODY_KEY).Robotic
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarBiomechanical = Veh.Components(BODY_KEY).Biomechanical
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarLivingMetal = Veh.Components(BODY_KEY).LivingMetal
		mvarCost = 0
		mvarWeight = 0
		mvarVolume = 0
		mvarSurfaceArea = 0
		mvarHitPoints = 0
		mvarFlightDeckOption = "none"
		mvarOrientation = "top"
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'the class is being destroyed
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
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
			mvarTotalDeckArea = System.Math.Round(CalcTotalDeckArea(mvarSurfaceArea, mvarKey), 2)
			mvarFlightDeckArea = System.Math.Round(mvarTotalDeckArea * (mvarPercentFlightDeck / 100), 2)
			mvarFlightDeckLength = System.Math.Round(3 * System.Math.Sqrt(mvarFlightDeckArea), 2)
			mvarCoveredDeckArea = System.Math.Round(mvarTotalDeckArea * (mvarPercentCovered / 100), 2)
			mvarDeckCost = CalcDeckCost(mvarFlightDeckArea, mvarCoveredDeckArea, mvarFlightDeckOption)
			mvarDeckWeight = CalcDeckWeight(mvarFlightDeckArea, mvarCoveredDeckArea, mvarFlightDeckOption)
		End If
		' Calculate the cost
		mvarCost = mvarDeckCost + BasicDesignCost(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials, mvarResponsive, mvarRobotic, mvarBiomechanical, mvarLivingMetal)
		' Calculate the weight
		mvarWeight = mvarDeckWeight + BasicDesignWeight(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials)
		' Calculate the Hit Points
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		mvarHitPoints = CalcHitPoints(TypeName(Me), mvarFrameStrength, mvarSurfaceArea)
		
		'generate print output
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarPrintOutput = CStr(CDbl(mvarCustomDescription & " (" & mvarOrientation & " of ") + Veh.Components(mvarParent).CustomDescription + CDbl(")."))
	End Sub
	
	Public Sub CalcCompartmentalizationStats()
		'this routine must be called AFTER the total vehicles
		'structural surface area is known.  It then can
		'compute the cost and weight associated with compartmentalizing
		'this subassembly
		Dim nDivisor As Single
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nDivisor = Veh.Stats.StructuralSurfaceArea
		
		If (mvarCompartmentalization <> "none") And (nDivisor > 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Veh.Components(Parent).StatsUpdate()
		End If
		
	End Sub
	
	Public Sub QueryChild()
		' if the object has children, query it and check to see if
		' more stats/property updates are needed for other objects in the collection
		' (see the StatusUpdate property for help on checking for childeren.  Can i use that one in place of this?)
		
	End Sub
End Class