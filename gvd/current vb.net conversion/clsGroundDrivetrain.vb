Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsGroundDrivetrain_NET.clsGroundDrivetrain")> Public Class clsGroundDrivetrain
	
	
	Private mvarTL As Short
	Private mvarMotivePower As Single
	Private mvarWeight As Double
	Private mvarVolume As Double
	Private mvarCost As Double
	Private mvarPowerReqt As Double
	Private mvarLocation As String
	Private mvarParent As String
	Private mvarKey As String
	Private mvarDR As Integer
	Private mvarRuggedized As Boolean
	Private mvarSurfaceArea As Double
	Private mvarHitPoints As Double
	Private mvarDatatype As Short
	Private mvarParentDatatype As Short
	Private mvarDescription As String
	Private mvarCustomDescription As String
	Private mvarCustom As Boolean
	
	'holds the index for DrivetrainMatrix
	Private mvarImage As Short
	Private mvarSelectedImage As Short
	Private mvarComment As String
	Private mvarCName As String
	Private mvarMatrixPos As Integer
	
	Private mvarPrintOutput As String
	Private mvarZZInit As Byte
	Private mvarLogicalParent As String
	
	
	Public Property LogicalParent() As String
		Get
			LogicalParent = mvarLogicalParent
		End Get
		Set(ByVal Value As String)
			mvarLogicalParent = Value
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
	
	
	
	
	Public Property MatrixPos() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.MatrixPos
			MatrixPos = mvarMatrixPos
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.MatrixPos = 5
			mvarMatrixPos = Value
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
	
	
	
	
	
	Public Property Ruggedized() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Ruggedized
			Ruggedized = mvarRuggedized
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Ruggedized = 5
			mvarRuggedized = Value
			If mvarZZInit = 0 Then Exit Property
			On Error GoTo errorhandler
			
			Dim legarray() As String
			Dim UpdateSiblings As Boolean
			Dim NumSiblings As Integer
			Dim i As Integer
			
			'get the keys for our sibling leg drivetrains
			If mvarDatatype = LegDrivetrain Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				legarray = VB6.CopyArray(Veh.KeyManager.GetCurrentLegDrivetrainKeys)
				If legarray(1) <> "" Then
					UpdateSiblings = True
					NumSiblings = UBound(legarray)
				End If
			End If
			
			'update the siblings as well
			If UpdateSiblings Then
				For i = 1 To NumSiblings
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Veh.Components(legarray(i)).Ruggedized <> Value Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Veh.Components(legarray(i)).Ruggedized = Value
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Veh.Components(legarray(i)).StatsUpdate()
					End If
				Next 
			End If
			Exit Property
errorhandler: 
			'when loading a save vehicle, it will try to update the sibling
			'which has not yet been created
			Exit Property
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
			If mvarZZInit = 0 Then Exit Property
			
			On Error GoTo errorhandler
			
			Dim legarray() As String
			Dim UpdateSiblings As Boolean
			Dim NumSiblings As Integer
			Dim i As Integer
			
			'get the keys for our sibling leg drivetrains
			If mvarDatatype = LegDrivetrain Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				legarray = VB6.CopyArray(Veh.KeyManager.GetCurrentLegDrivetrainKeys)
				If legarray(1) <> "" Then
					UpdateSiblings = True
					NumSiblings = UBound(legarray)
				End If
			End If
			
			
			'update the siblings as well
			If UpdateSiblings Then
				For i = 1 To NumSiblings
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Veh.Components(legarray(i)).DR <> Value Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Veh.Components(legarray(i)).DR = Value
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Veh.Components(legarray(i)).StatsUpdate()
					End If
				Next 
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
	
	
	
	Public Property PowerReqt() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PowerReqt
			PowerReqt = mvarPowerReqt
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PowerReqt = 5
			mvarPowerReqt = Value
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
	
	
	
	
	
	Public Property MotivePower() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.MotivePower
			MotivePower = mvarMotivePower
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.MotivePower = 5
			mvarMotivePower = Value
			If mvarZZInit = 0 Then Exit Property
			On Error GoTo errorhandler
			
			Dim legarray() As String
			Dim UpdateSiblings As Boolean
			Dim NumSiblings As Integer
			Dim i As Integer
			
			'get the keys for our sibling leg drivetrains
			If mvarDatatype = LegDrivetrain Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				legarray = VB6.CopyArray(Veh.KeyManager.GetCurrentLegDrivetrainKeys)
				If legarray(1) <> "" Then
					UpdateSiblings = True
					NumSiblings = UBound(legarray)
				End If
			End If
			
			'update the siblings as well
			If UpdateSiblings Then
				For i = 1 To NumSiblings
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Veh.Components(legarray(i)).MotivePower <> Value Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Veh.Components(legarray(i)).MotivePower = Value
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Veh.Components(legarray(i)).StatsUpdate()
					End If
				Next 
			End If
			Exit Property
errorhandler: 
			'when loading a save vehicle, it will try to update the sibling
			'which has not yet been created
			Exit Property
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
			
			Dim legarray() As String
			Dim UpdateSiblings As Boolean
			Dim NumSiblings As Integer
			Dim i As Integer
			
			On Error GoTo errorhandler
			
			mvarTL = Value
			If mvarZZInit = 0 Then Exit Property
			
			'get the keys for our sibling leg drivetrains
			If mvarDatatype = LegDrivetrain Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				legarray = VB6.CopyArray(Veh.KeyManager.GetCurrentLegDrivetrainKeys)
				If legarray(1) <> "" Then
					UpdateSiblings = True
					NumSiblings = UBound(legarray)
				End If
			End If
			
			'update the siblings as well
			If UpdateSiblings Then
				For i = 1 To NumSiblings
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Veh.Components(legarray(i)).TL <> Value Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Veh.Components(legarray(i)).TL = Value
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Veh.Components(legarray(i)).StatsUpdate()
					End If
				Next 
			End If
			
			GetMatrixIndex()
			Exit Property
errorhandler: 
			'when loading a save vehicle, it will try to update the sibling
			'which has not yet been created
			Exit Property
		End Set
	End Property
	
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		
		' set the default properties
		mvarCustom = False
		TL = gVehicleTL
		mvarRuggedized = False
		mvarMotivePower = 60
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
	
	Public Sub Init()
		Dim legarray() As String
		Dim i As Integer
		
		Select Case mvarDatatype
			Case WheeledDrivetrain
				
			Case AllWheelDriveWheeledDrivetrain
				
			Case TrackedDrivetrain
				
			Case LegDrivetrain
				
				
				'if this leg drivetrain has been added AFTER other leg drivetrains have been added
				'then this ddrivetrains stats should default to existing stats
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				legarray = VB6.CopyArray(Veh.KeyManager.GetCurrentLegDrivetrainKeys)
				If legarray(1) <> "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarTL = Veh.Components(legarray(1)).TL
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarRuggedized = Veh.Components(legarray(1)).Ruggedized
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarMotivePower = Veh.Components(legarray(1)).MotivePower
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarDR = Veh.Components(legarray(1)).DR
					'must now update all the other legs statsupdate since adding this legdrivetrain and leg
					'results in changed average volume due to increased leg count
					For i = 1 To UBound(legarray)
						'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Veh.Components(legarray(i)).StatsUpdate()
					Next 
				End If
				
			Case FlexibodyDrivetrain
				
		End Select
		
	End Sub
	
	Public Sub GetMatrixIndex()
		Dim i As Short
		If mvarDatatype = 0 Then Exit Sub
		mvarMatrixPos = 0 'init the counter
		For i = 1 To UBound(DrivetrainMatrix)
			If DrivetrainMatrix(i).ID = mvarDatatype Then
				If DrivetrainMatrix(i).TL >= mvarTL Then
					mvarMatrixPos = i
					Exit For
				Else
					mvarMatrixPos = i
				End If
			End If
		Next 
	End Sub
	
	Public Sub StatsUpdate()
		mvarZZInit = 1
		If mvarMatrixPos = 0 Then Exit Sub
		
		Dim TempWeight1 As Single
		Dim TempWeight2 As Single
		Dim NumLegs As Integer
		Dim CostMod As Short
		Dim legarray() As String
		Dim i As Integer
		Dim QRugMod As Single 'combined quantity and ruggedized multipliers
		Dim RugHitMod As Short 'ruggedized hit point multiplier
		Dim sPrint1 As String
		
		mvarLocation = GetLocation
		
		'set the ruggedized and quantity multipliers
		If mvarRuggedized Then
			QRugMod = 1.5
			RugHitMod = 2
		Else
			QRugMod = 1
			RugHitMod = 1
		End If
		
		
		' Find how many legs are on the vehicle and apply modifiers
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		legarray = VB6.CopyArray(Veh.KeyManager.GetCurrentLegKeys)
		NumLegs = UBound(legarray)
		
		' a cost modifier is applied to Legs depending on how
		' many are installed.. this is due to the need for
		' extra stabilization and control systems
		If NumLegs < 2 Then
			NumLegs = 1
			CostMod = 1 'default modifier
		End If
		If NumLegs = 3 Then
			CostMod = 2
		ElseIf NumLegs = 2 Then 
			CostMod = 4
		Else
			CostMod = 1
		End If
		
		
		If mvarDatatype = LegDrivetrain Then
			' NOTE: Here we must determine the correct >=5 formula.
			' Vehicles 2dE assumes a single drivetrain.  GVD uses
			' seperate motors so the >=5kW must take into account total
			' number of legs
			If mvarMotivePower * NumLegs >= 5 Then
				TempWeight1 = DrivetrainMatrix(mvarMatrixPos).Weight2
				TempWeight2 = DrivetrainMatrix(mvarMatrixPos).Weight3
			Else
				TempWeight1 = DrivetrainMatrix(mvarMatrixPos).Weight1
				TempWeight2 = 0
			End If
			
			' NOTE: must divide the TempWeight2 by the number of legs since GVD
			' calcs all stats for each motor as seperate components and not as
			' one single drivetrain
			TempWeight1 = (mvarMotivePower * TempWeight1) + (TempWeight2 / NumLegs)
		Else
			
			If mvarMotivePower >= 5 Then
				TempWeight1 = DrivetrainMatrix(mvarMatrixPos).Weight2
				TempWeight2 = DrivetrainMatrix(mvarMatrixPos).Weight3
			Else
				TempWeight1 = DrivetrainMatrix(mvarMatrixPos).Weight1
				TempWeight2 = 0
			End If
			
			TempWeight1 = (mvarMotivePower * TempWeight1) + TempWeight2
		End If
		
		
		
		'get base stats
		mvarWeight = TempWeight1
		mvarCost = mvarWeight * DrivetrainMatrix(mvarMatrixPos).Cost * CostMod
		mvarVolume = mvarWeight / DrivetrainMatrix(mvarMatrixPos).Volume
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
		
		'get finals
		mvarWeight = System.Math.Round(QRugMod * mvarWeight, 2)
		mvarCost = System.Math.Round(QRugMod * mvarCost, 2)
		mvarVolume = System.Math.Round(QRugMod * mvarVolume, 2)
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		
		mvarPowerReqt = System.Math.Round(mvarMotivePower, 2)
		
		'produce the print output
		If mvarRuggedized Then
			sPrint1 = "ruggedized "
		Else
			sPrint1 = ""
		End If
		
		mvarPrintOutput = "TL" & mvarTL & " " & VB6.Format(mvarMotivePower, p_sFormat) & " kW " & sPrint1 & mvarCustomDescription & " (" & mvarLocation & ", HP " & mvarHitPoints & ", " & VB6.Format(mvarWeight, p_sFormat) & " lbs., " & VB6.Format(mvarVolume, p_sFormat) & " cf., " & "$" & VB6.Format(mvarCost, p_sFormat) & ", " & VB6.Format(mvarPowerReqt, p_sFormat) & " kW)." & mvarComment
		
		
		
		
	End Sub
	
	Public Sub QueryParent()
		' if the object has a parent, query it and check to see if
		' more stats/property updates are needed for other objects in the collection
		If mvarParent <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Veh.Components(Parent).StatsUpdate()
		End If
	End Sub
	
	Public Function LocationCheck() As Boolean
		Dim TempCheck As Boolean
		Dim otherdrivetrainarray() As String
		Dim component As String
		Dim i As Integer
		
		'determine if the user is adding more drivetrains than rotors
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		otherdrivetrainarray = VB6.CopyArray(Veh.KeyManager.GetCurrentOtherGroundDrivetrainKeys)
		
		If mvarDatatype <> LegDrivetrain Then
			If otherdrivetrainarray(1) <> "" Then
				For i = 1 To UBound(otherdrivetrainarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					component = Veh.Components(otherdrivetrainarray(i)).Datatype
					If CDbl(component) = mvarDatatype Then
						TempCheck = False
						modHelper.InfoPrint(1, "Only one of these types of drivetrains can be installed onto a Vehicle.")
						LocationCheck = TempCheck
						Exit Function
						'need to check that allwheeldrivewheeledrivetrain checks regular wheeleddrivetrain
					ElseIf CDbl(component) = WheeledDrivetrain And mvarDatatype = AllWheelDriveWheeledDrivetrain Then 
						TempCheck = False
						modHelper.InfoPrint(1, "Only one of these types of drivetrains can be installed onto a Vehicle.")
						LocationCheck = TempCheck
						Exit Function
					ElseIf CDbl(component) = AllWheelDriveWheeledDrivetrain And mvarDatatype = WheeledDrivetrain Then 
						TempCheck = False
						modHelper.InfoPrint(1, "Only one of these types of drivetrains can be installed onto a Vehicle.")
						LocationCheck = TempCheck
						Exit Function
					End If
				Next 
			End If
		End If
		
		'determine if the user is placing the component in a valid location
		Select Case mvarDatatype
			
			Case LegDrivetrain
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(Parent).Datatype <> Leg Then
					modHelper.InfoPrint(1, "Leg Drivetrains can only be placed in the Vehicle's Legs")
					TempCheck = False
				Else
					TempCheck = True
				End If
			Case Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(Parent).Datatype <> Body Then
					modHelper.InfoPrint(1, "Wheeled, Flexibody and Tracked Drivetrains must be placed in the Vehicle's Body")
					TempCheck = False
				Else
					TempCheck = True
				End If
		End Select
		
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
End Class