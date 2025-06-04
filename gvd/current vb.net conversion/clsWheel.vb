Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsWheel_NET.clsWheel")> Public Class clsWheel
	
	'local variable(s) to hold property value(s)
	Private mvarSubType As String
	Private mvarParent As String
	Private mvarKey As String
	
	Private mvarImage As Short
	Private mvarSelectedImage As Short
	Private mvarTL As Short
	Private mvarImprovedSuspension As Boolean
	Private mvarImprovedSuspensionCost As Single
	Private mvarWheelblades As String
	Private mvarSnowTires As Boolean
	Private mvarRacingTires As Boolean
	Private mvarPunctureResistant As Boolean
	Private mvarRetractLocation As String
	Private mvarImprovedBrakes As Boolean
	Private mvarAllwheelSteering As Boolean
	Private mvarSmartwheels As Boolean
	Private mvarFrameStrength As String
	Private mvarMaterials As String
	Private mvarResponsive As Boolean
	Private mvarRobotic As Boolean
	Private mvarBiomechanical As Boolean
	Private mvarLivingMetal As Boolean
	Private mvarCost As Double
	Private mvarImprovedBrakesCost As Single
	Private mvarAllWheelSteeringCost As Single
	Private mvarSmartWheelsCost As Single
	Private mvarSnowTiresCost As Single
	Private mvarRacingTiresCost As Single
	Private mvarPunctureResistantCost As Single
	Private mvarWheelBladesCost As Single
	Private mvarWheelBladesWeight As Single
	Private mvarEmptySpace As Single 'MPJ 6/30/2000  added so users can have Monster Tires which also decrease ground pressure
	Private mvarWeight As Double
	Private mvarSurfaceArea As Double
	Private mvarVolume As Double
	Private mvarHitPoints As Double
	Private mvarLocation As String
	Private mvarDR As Integer
	Private mvarDatatype As Short
	Private mvarParentDatatype As Short
	Private mvarDescription As String
	Private mvarCustomDescription As String
	Private mvarCustom As Boolean
	Private mvarQuantity As Short
	'local variable(s) to hold property value(s)
	Private mvarComment As String
	Private mvarCName As String
	'local variable(s) to hold property value(s)
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
	
	
	
	
	Public Property Quantity() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Quantity
			Quantity = mvarQuantity
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Quantity = 5
			mvarQuantity = Value
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
	
	
	
	
	
	Public Property SubType() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			
			SubType = mvarSubType
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			mvarSubType = Value
			If mvarZZInit = 0 Then Exit Property
			
			If Value <> "retractable" Then
				If mvarRetractLocation <> "none" Then
					modHelper.InfoPrint(1, "Rectract location invalid with this wheel type.  Retract location has been reset to 'none'")
					mvarRetractLocation = "none"
				End If
			End If
			
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
	
	
	
	Public Property ImprovedSuspensionCost() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.ImprovedSuspensionCost
			ImprovedSuspensionCost = mvarImprovedSuspensionCost
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.ImprovedSuspensionCost = 5
			mvarImprovedSuspensionCost = Value
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
	
	
	
	Public Property ImprovedBrakesCost() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.ImprovedBrakesCost
			ImprovedBrakesCost = mvarImprovedBrakesCost
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.ImprovedBrakesCost = 5
			mvarImprovedBrakesCost = Value
		End Set
	End Property
	
	
	
	Public Property AllWheelSteeringCost() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.AllWheelSteeringCost
			AllWheelSteeringCost = mvarAllWheelSteeringCost
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.AllWheelSteeringCost = 5
			mvarAllWheelSteeringCost = Value
		End Set
	End Property
	
	
	
	Public Property SmartWheelsCost() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.SmartWheelscost
			SmartWheelsCost = mvarSmartWheelsCost
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.SmartWheelscost = 5
			mvarSmartWheelsCost = Value
		End Set
	End Property
	
	
	
	
	Public Property PunctureResistantCost() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PunctureResistantCost
			PunctureResistantCost = mvarPunctureResistantCost
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PunctureResistantCost = 5
			mvarPunctureResistantCost = Value
		End Set
	End Property
	
	
	
	Public Property RacingTiresCost() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.RacingTiresCost
			RacingTiresCost = mvarRacingTiresCost
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.RacingTiresCost = 5
			mvarRacingTiresCost = Value
		End Set
	End Property
	
	
	
	Public Property SnowTiresCost() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.SnowTiresCost
			SnowTiresCost = mvarSnowTiresCost
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.SnowTiresCost = 5
			mvarSnowTiresCost = Value
		End Set
	End Property
	
	
	
	Public Property WheelBladesCost() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.WheelBladesCost
			WheelBladesCost = mvarWheelBladesCost
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.WheelBladesCost = 5
			mvarWheelBladesCost = Value
		End Set
	End Property
	
	
	
	Public Property WheelBladesWeight() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.WheelBladesWeight
			WheelBladesWeight = mvarWheelBladesWeight
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.WheelBladesWeight = 5
			mvarWheelBladesWeight = Value
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
	
	
	
	
	
	Public Property Smartwheels() As Boolean
		Get
			
			
			
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Smartwheels
			Smartwheels = mvarSmartwheels
			Exit Property
			
			
		End Get
		Set(ByVal Value As Boolean)
			
			
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Smartwheels = 5
			mvarSmartwheels = Value
			Exit Property
			
		End Set
	End Property
	
	
	
	
	
	Public Property AllwheelSteering() As Boolean
		Get
			
			
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.AllwheelSteering
			AllwheelSteering = mvarAllwheelSteering
			
		End Get
		Set(ByVal Value As Boolean)
			
			
			
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.AllwheelSteering = 5
			mvarAllwheelSteering = Value
			
			
			
		End Set
	End Property
	
	
	
	Public Property ImprovedBrakes() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.ImprovedBrakes
			ImprovedBrakes = mvarImprovedBrakes
			
			
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.ImprovedBrakes = 5
			mvarImprovedBrakes = Value
		End Set
	End Property
	
	
	
	
	
	Public Property RetractLocation() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.RetractLocation
			RetractLocation = mvarRetractLocation
			
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.RetractLocation = 5
			mvarRetractLocation = Value
			If mvarZZInit = 0 Then Exit Property
			
			If Value <> "none" Then
				If mvarSubType <> "retractable" Then
					mvarSubType = "retractable"
				End If
			End If
		End Set
	End Property
	
	
	
	
	
	Public Property PunctureResistant() As Boolean
		Get
			
			
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PunctureResistant
			PunctureResistant = mvarPunctureResistant
		End Get
		Set(ByVal Value As Boolean)
			
			
			
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PunctureResistant = 5
			mvarPunctureResistant = Value
			
			
		End Set
	End Property
	
	
	
	
	
	Public Property RacingTires() As Boolean
		Get
			
			
			
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.RacingTires
			RacingTires = mvarRacingTires
		End Get
		Set(ByVal Value As Boolean)
			
			
			
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.RacingTires = 5
			mvarRacingTires = Value
		End Set
	End Property
	
	
	
	
	
	Public Property SnowTires() As Boolean
		Get
			
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.SnowTires
			SnowTires = mvarSnowTires
		End Get
		Set(ByVal Value As Boolean)
			
			
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.SnowTires = 5
			mvarSnowTires = Value
		End Set
	End Property
	
	
	
	Public Property Wheelblades() As String
		Get
			
			
			
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Wheelblades
			Wheelblades = mvarWheelblades
		End Get
		Set(ByVal Value As String)
			
			
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Wheelblades = 5
			mvarWheelblades = Value
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
	
	
	
	
	
	Public Function LocationCheck() As Boolean
		Dim TempCheck As Boolean
		Dim InstallPoint As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallPoint = Veh.Components(mvarParent).Datatype
		
		If InstallPoint = Body Then
			TempCheck = True
		Else
			modHelper.InfoPrint(1, "Wheels must be placed on hull.")
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
		mvarSubType = "standard"
		mvarQuantity = CShort("4")
		mvarImprovedSuspension = False
		mvarWheelblades = "none"
		mvarSnowTires = False
		mvarRacingTires = False
		mvarPunctureResistant = False
		mvarRetractLocation = "none"
		mvarImprovedBrakes = False
		mvarAllwheelSteering = False
		mvarSmartwheels = False
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
		mvarSurfaceArea = 0
		mvarVolume = 0
		mvarHitPoints = 0
		
		
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
		
		Dim sPrint1 As String
		Dim sPrint2 As String
		mvarZZInit = 1
		
		mvarLocation = GetLocation
		mvarAbbrev = "Wheel"
		
		' Calculate the component volume
		'TODO: NOTE- The rule is simply, we get the volume of the body. But,
		'      when using sequenced 'only necessary' object updates, we need to
		'      get the Body's volume after the body's .Update has been performed.
		'      The body is always the "second to last" then so to speak.  Since wheels/tracks
		'      and a few selected others, will update afterwards.  We should remember that
		'      these subassemblies dont influence the stats of the body (obviously) since if it
		'     did, you'd have an infinite loop of updates.
		Select Case mvarSubType
			Case "small", "retractable"
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarVolume = Veh.Components(BODY_KEY).Volume * 0.05
			Case "standard"
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarVolume = Veh.Components(BODY_KEY).Volume * 0.1
			Case "heavy", "off-road", "railway"
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarVolume = Veh.Components(BODY_KEY).Volume * 0.2
		End Select
		
		
		
		mvarVolume = System.Math.Round(mvarVolume + mvarEmptySpace, 2)
		' calculate the surface
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		
		'Improved Suspension cost
		If mvarImprovedSuspension Then mvarImprovedSuspensionCost = mvarSurfaceArea * 100 Else mvarImprovedSuspensionCost = 0
		'improvedbrakes cost
		If mvarImprovedBrakes Then
			mvarImprovedBrakesCost = 20 * mvarSurfaceArea
			If mvarTL = 8 Then
				mvarImprovedBrakesCost = mvarImprovedBrakesCost / 2
			ElseIf mvarTL >= 9 Then 
				mvarImprovedBrakesCost = mvarImprovedBrakesCost / 4
			End If
		Else
			mvarImprovedBrakesCost = 0
		End If
		'allwheel steering cost
		If mvarAllwheelSteering Then
			If mvarTL <= 7 Then
				mvarAllWheelSteeringCost = 100 * mvarSurfaceArea
				'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarAllWheelSteeringCost = modPerformance.Maximum(mvarAllWheelSteeringCost, 5000)
			ElseIf mvarTL = 8 Then 
				mvarAllWheelSteeringCost = 50 * mvarSurfaceArea
				'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarAllWheelSteeringCost = modPerformance.Maximum(mvarAllWheelSteeringCost, 2500)
			ElseIf mvarTL >= 9 Then 
				mvarAllWheelSteeringCost = 25 * mvarSurfaceArea
				'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarAllWheelSteeringCost = modPerformance.Maximum(mvarAllWheelSteeringCost, 1250)
			End If
		Else
			mvarAllWheelSteeringCost = 0
		End If
		
		'smartwheels cost
		If mvarSmartwheels Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarSmartWheelsCost = modPerformance.Maximum(80 * mvarSurfaceArea, 4000)
			If mvarTL = 9 Then
				mvarSmartWheelsCost = mvarSmartWheelsCost / 2
			ElseIf mvarTL >= 10 Then 
				mvarSmartWheelsCost = mvarSmartWheelsCost / 4
			End If
		Else
			mvarSmartWheelsCost = 0
		End If
		
		'snowtires costs
		If mvarSnowTires Then
			If mvarSurfaceArea >= 200 Then
				mvarSnowTiresCost = 200 * mvarQuantity
			Else
				mvarSnowTiresCost = 100 * mvarQuantity
			End If
		Else
			mvarSnowTiresCost = 0
		End If
		
		'racing tires cost
		If mvarRacingTires Then
			If mvarSurfaceArea >= 200 Then
				mvarRacingTiresCost = 500 * mvarQuantity
			Else
				mvarRacingTiresCost = 250 * mvarQuantity
			End If
		Else
			mvarRacingTiresCost = 0
		End If
		
		'puncture resistance tirres cost
		If mvarPunctureResistant Then
			If mvarSurfaceArea >= 200 Then
				mvarPunctureResistantCost = 500 * mvarQuantity
			Else
				mvarPunctureResistantCost = 250 * mvarQuantity
			End If
		Else
			mvarPunctureResistantCost = 0
		End If
		
		'wheelbladescost and weight
		If mvarWheelblades <> "none" Then
			If mvarWheelblades = "rectractable" Then
				mvarWheelBladesWeight = 0.2 * mvarSurfaceArea
				mvarWheelBladesCost = 100 * mvarWheelBladesWeight
			Else
				mvarWheelBladesWeight = 0.1 * mvarSurfaceArea
				mvarWheelBladesCost = 100 * mvarWheelBladesWeight
			End If
		Else
			mvarWheelBladesCost = 0
			mvarWheelBladesWeight = 0
		End If
		' Calculate the cost
		mvarCost = BasicDesignCost(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials, mvarResponsive, mvarRobotic, mvarBiomechanical, mvarLivingMetal)
		' Calculate the weight
		mvarWeight = BasicDesignWeight(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials)
		' Calculate the Hit Points
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		mvarHitPoints = CalcHitPoints(TypeName(Me), mvarFrameStrength, mvarSurfaceArea, mvarQuantity)
		
		'generate the print output
		If mvarRetractLocation <> "none" Then
			sPrint1 = ", rectract into " & mvarRetractLocation
		End If
		If mvarImprovedBrakes Then
			sPrint2 = sPrint2 & ", improved brakes"
		End If
		If mvarAllwheelSteering Then
			sPrint2 = sPrint2 & ", all-wheel steering"
		End If
		If mvarSmartwheels Then
			sPrint2 = sPrint2 & ", smart wheels"
		End If
		mvarPrintOutput = mvarSubType & " " & mvarCustomDescription & " (" & VB6.Format(mvarQuantity) & " wheels " & sPrint2 & sPrint1 & ")."
		
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