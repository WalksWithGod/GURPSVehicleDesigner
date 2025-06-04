Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsWing_NET.clsWing")> Public Class clsWing
	
	Private mvarOrientation As String
	Private mvarImage As Short
	Private mvarSelectedImage As Short
	Private mvarParent As String
	Private mvarKey As String
	Private mvarSubType As String
	Private mvarWeight As Double
	Private mvarCost As Double
	Private mvarFrameStrength As String
	Private mvarMaterials As String
	Private mvarResponsive As Boolean
	Private mvarRobotic As Boolean
	Private mvarBiomechanical As Boolean
	Private mvarLivingMetal As Boolean
	Private mvarSurfaceArea As Double
	Private mvarVolume As Double
	Private mvarTL As Short
	Private mvarControlledInstability As Boolean
	Private mvarControlledInstabilityCost As Single
	Private mvarFoldingCost As Single
	Private mvarFoldingWeight As Single
	Private mvarVariableSweepCost As Single
	Private mvarVariableSweepWeight As Single
	Private mvarFolding As Boolean
	Private mvarVariableSweep As String
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
	Private mvarSiblingKey As String
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
			Dim siblingorientation As String
			
			mvarOrientation = Value
			If mvarZZInit = 0 Then Exit Property
			
			On Error GoTo errorhandler
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			siblingorientation = Veh.Components(mvarSiblingKey).Orientation
			
			'update the sibling as well
			If mvarSiblingKey <> "" Then
				If siblingorientation <> Value Then
				ElseIf siblingorientation = "right" Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).Orientation = "left"
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).Orientation = "right"
				End If
			End If
			Exit Property
errorhandler: 
			'when loading a save vehicle, it will try to update the sibling
			'which has not yet been created
			Exit Property
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
	
	
	
	
	
	Public Property VariableSweep() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.VariableSweep
			VariableSweep = mvarVariableSweep
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.VariableSweep = 5
			mvarVariableSweep = Value
			If mvarZZInit = 0 Then Exit Property
			
			On Error GoTo errorhandler
			'update the sibling as well
			If mvarSiblingKey <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarSiblingKey).VariableSweep <> Value Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).VariableSweep = Value
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).StatsUpdate()
				End If
			End If
			Exit Property
errorhandler: 
			'when loading a save vehicle, it will try to update the sibling
			'which has not yet been created
			Exit Property
		End Set
	End Property
	
	
	
	
	
	Public Property Folding() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Folding
			Folding = mvarFolding
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Folding = 5
			mvarFolding = Value
			If mvarZZInit = 0 Then Exit Property
			
			'update the sibling as well
			If mvarSiblingKey <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarSiblingKey).Folding <> Value Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).Folding = Value
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).StatsUpdate()
				End If
			End If
		End Set
	End Property
	
	
	
	
	
	Public Property ControlledInstability() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.ControlledInstability
			ControlledInstability = mvarControlledInstability
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.ControlledInstability = 5
			mvarControlledInstability = Value
			If mvarZZInit = 0 Then Exit Property
			
			'update the sibling as well
			If mvarSiblingKey <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarSiblingKey).ControlledInstability <> Value Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).ControlledInstability = Value
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).StatsUpdate()
				End If
			End If
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
			If mvarZZInit = 0 Then Exit Property
			
			On Error GoTo errorhandler
			'update the sibling as well
			If mvarSiblingKey <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarSiblingKey).TL <> Value Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).TL = Value
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).StatsUpdate()
				End If
			End If
			Exit Property
errorhandler: 
			'when loading a save vehicle, it will try to update the sibling
			'which has not yet been created
			Exit Property
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
			If mvarZZInit = 0 Then Exit Property
			
			'update the sibling as well
			If mvarSiblingKey <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarSiblingKey).LivingMetal <> Value Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).LivingMetal = Value
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).StatsUpdate()
				End If
			End If
			
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
			If mvarZZInit = 0 Then Exit Property
			
			'update the sibling as well
			If mvarSiblingKey <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarSiblingKey).Biomechanical <> Value Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).Biomechanical = Value
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).StatsUpdate()
				End If
			End If
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
			If mvarZZInit = 0 Then Exit Property
			
			'update the sibling as well
			If mvarSiblingKey <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarSiblingKey).Robotic <> Value Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).Robotic = Value
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).StatsUpdate()
				End If
			End If
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
			If mvarZZInit = 0 Then Exit Property
			
			'update the sibling as well
			If mvarSiblingKey <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarSiblingKey).Responsive <> Value Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).Responsive = Value
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).StatsUpdate()
				End If
			End If
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
			If mvarZZInit = 0 Then Exit Property
			
			'update the sibling as well
			If mvarSiblingKey <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarSiblingKey).Materials <> Value Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).Materials = Value
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).StatsUpdate()
				End If
			End If
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
			If mvarZZInit = 0 Then Exit Property
			
			'update the sibling as well
			If mvarSiblingKey <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarSiblingKey).FrameStrength <> Value Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).FrameStrength = Value
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).StatsUpdate()
				End If
			End If
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
	
	
	
	
	Public Property FoldingCost() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.FoldingCost
			FoldingCost = mvarFoldingCost
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.FoldingCost = 5
			mvarFoldingCost = Value
		End Set
	End Property
	
	
	
	Public Property VariableSweepCost() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.VariableSweepCost
			VariableSweepCost = mvarVariableSweepCost
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.VariableSweepCost = 5
			mvarVariableSweepCost = Value
		End Set
	End Property
	
	
	
	Public Property ControlledInstabilityCost() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.ControlledInstabilityCost
			ControlledInstabilityCost = mvarControlledInstabilityCost
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.ControlledInstabilityCost = 5
			mvarControlledInstabilityCost = Value
		End Set
	End Property
	
	
	
	Public Property FoldingWeight() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.FoldingWeight
			FoldingWeight = mvarFoldingWeight
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.FoldingWeight = 5
			mvarFoldingWeight = Value
		End Set
	End Property
	
	
	
	Public Property VariableSweepWeight() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.VariableSweepWeight
			VariableSweepWeight = mvarVariableSweepWeight
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.VariableSweepWeight = 5
			mvarVariableSweepWeight = Value
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
	
	
	
	
	
	Public Property SubType() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			
			SubType = mvarSubType
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			
			mvarSubType = Value
			If mvarZZInit = 0 Then Exit Property
			
			On Error GoTo errorhandler
			
			'update the sibling as well
			If mvarSiblingKey <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarSiblingKey).SubType <> Value Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).SubType = Value
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Veh.Components(mvarSiblingKey).StatsUpdate()
				End If
			End If
			Exit Property
errorhandler: 
			'when loading a save vehicle, it will try to update the sibling
			'which has not yet been created
			Exit Property
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
	
	
	
	
	Public Property SiblingKey() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.SiblingKey
			SiblingKey = mvarSiblingKey
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.SiblingKey = 5
			mvarSiblingKey = Value
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
	
	
	
	
	Public Function LocationCheck() As Boolean
		Dim TempCheck As Boolean
		Dim InstallPoint As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallPoint = Veh.Components(mvarParent).Datatype
		
		If InstallPoint = Body Then
			TempCheck = True
		Else
			modHelper.InfoPrint(1, "Wings must be placed on hull.")
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
		mvarOrientation = "left"
		mvarFolding = False
		mvarVariableSweep = "none"
		mvarControlledInstability = False
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
		Dim sPrint2 As String
		Dim element As Object
		
		mvarZZInit = 1
		mvarLocation = GetLocation
		
		If mvarOrientation = "left" Then
			mvarAbbrev = "WiL"
		Else
			mvarAbbrev = "WiR"
		End If
		
		If mvarIndex > 0 Then mvarAbbrev = mvarAbbrev & mvarIndex
		
		'get the accessspace
		mvarAccessSpace = CalcAccessSpace(mvarKey)
		' Calculate the component volume
		mvarVolume = CalcCombinedVolume(mvarKey) + mvarEmptySpace + mvarAccessSpace
		' Calculate volume due to retracts that are of the Body & Wing variety
		If GetRetractLocation = "body & wings" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarVolume = mvarVolume + (0.025 * Veh.Components(BODY_KEY).Volume)
		End If
		mvarVolume = System.Math.Round(mvarVolume, 2)
		' calculate the surface
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		
		' apply the surface area multiplier
		Select Case mvarSubType
			Case "standard", "flarecraft"
				mvarSurfaceArea = mvarSurfaceArea * 1.5
				
			Case "high agility", "STOL"
				mvarSurfaceArea = mvarSurfaceArea * 2
			Case "biplane"
				mvarSurfaceArea = mvarSurfaceArea * 3
			Case "triplane"
				mvarSurfaceArea = mvarSurfaceArea * 4
				
			Case "stub"
				mvarSurfaceArea = mvarSurfaceArea * 1 ' no change according to rules page 18)
				
		End Select
		
		' Calculate the cost
		mvarCost = BasicDesignCost(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials, mvarResponsive, mvarRobotic, mvarBiomechanical, mvarLivingMetal)
		' Calculate the weight
		mvarWeight = BasicDesignWeight(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials)
		' Calculate the Hit Points
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		mvarHitPoints = CalcHitPoints(TypeName(Me), mvarFrameStrength, mvarSurfaceArea)
		'print output
		If mvarFolding Then
			sPrint2 = ", folding"
		End If
		If mvarVariableSweep <> "none" Then
			sPrint2 = sPrint2 & ", variable sweep"
		End If
		If mvarControlledInstability Then
			sPrint2 = sPrint2 & ", controlled instability"
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarPrintOutput = CStr(CDbl(mvarSubType & " " & mvarCustomDescription & " (" & mvarOrientation & " side of ") + Veh.Components(Parent).CustomDescription + CDbl(sPrint2) + CDbl(")."))
		
	End Sub
	
	Public Sub CalcWingRotorOptionWeightsAndCosts()
		Dim structarea As Single
		Dim structweight As Single
		Dim structcost As Single
		Dim divisor1 As Single
		Dim divisor2 As Single
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		structarea = Veh.Stats.StructuralSurfaceArea
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		structweight = Veh.Stats.StructuralWeight
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		structcost = Veh.Stats.StructuralCost
		
		divisor1 = structarea * structweight
		divisor2 = structarea * structcost
		
		'folding cost and weight
		If (mvarFolding) And (divisor1 <> 0) Then
			mvarFoldingWeight = 0.05 * mvarSurfaceArea / divisor1
			mvarFoldingCost = 25 * mvarFoldingWeight
		Else
			mvarFoldingWeight = 0
			mvarFoldingCost = 0
		End If
		
		'variable sweep cost and weight
		If (mvarVariableSweep <> "none") And (divisor1 <> 0) And (divisor2 <> 0) Then
			mvarVariableSweepWeight = 0.1 * mvarSurfaceArea / divisor1
			mvarVariableSweepCost = 0.25 * mvarSurfaceArea / divisor2
			If mvarVariableSweep = "automatic" Then mvarVariableSweepCost = mvarVariableSweepCost * 2
		Else
			mvarVariableSweepCost = 0
			mvarVariableSweepWeight = 0
		End If
		
		'controlled instability cost
		If mvarControlledInstability Then
			mvarControlledInstabilityCost = 100 * mvarSurfaceArea
		Else
			mvarControlledInstabilityCost = 0
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