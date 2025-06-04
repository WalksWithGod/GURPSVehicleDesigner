Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsSafetySystem_NET.clsSafetySystem")> Public Class clsSafetySystem
	
	'local variable(s) to hold property value(s)
	Private mvarTL As Short
	Private mvarWeight As Double
	Private mvarVolume As Double
	Private mvarCost As Double
	Private mvarPowerReqt As Double
	Private mvarLocation As String
	Private mvarParent As String
	Private mvarKey As String
	Private mvarDR As Integer
	Private mvarSurfaceArea As Double
	Private mvarHitPoints As Double
	Private mvarOccupancy As Integer
	Private mvarGReduction As Single
	Private mvarDatatype As Short
	Private mvarParentDatatype As Short
	Private mvarDescription As String
	Private mvarCustomDescription As String
	Private mvarCustom As Boolean
	Private mvarQuantity As Integer
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
	
	
	Public Property Occupancy() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.MatrixPos
			Occupancy = mvarOccupancy
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Occupancy = 5
			mvarOccupancy = Value
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
			'Syntax: Debug.Print X.Power
			PowerReqt = mvarPowerReqt
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Power = 5
			mvarPowerReqt = Value
		End Set
	End Property
	
	
	Public Property GReduction() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.GReduction
			GReduction = mvarGReduction
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.GReduction = 5
			mvarGReduction = Value
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
			GetMatrixIndex()
		End Set
	End Property
	
	
	
	
	Public Function LocationCheck() As Boolean
		Dim TempCheck As Boolean
		Dim InstallPoint As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallPoint = Veh.Components(mvarParent).Datatype
		
		
		Select Case mvarDatatype
			Case GravityWeb, EjectionSeat, Airbag, CrashWeb
				If (InstallPoint = Cabin) Or (InstallPoint = LuxuryCabin) Or (InstallPoint = Suite) Or (InstallPoint = LuxurySuite) Or (InstallPoint = CrampedCrewStation) Or (InstallPoint = NormalCrewStation) Or (InstallPoint = RoomyCrewStation) Or (InstallPoint = CycleCrewStation) Or (InstallPoint = HarnessCrewStation) Or (InstallPoint = CycleSeat) Or (InstallPoint = CrampedSeat) Or (InstallPoint = NormalSeat) Or (InstallPoint = RoomySeat) Then
					TempCheck = True
				Else
					modHelper.InfoPrint(1, "Ejection Seat, Airbag, Crashweb, and Gravity Webs must be added to a Crew Station, Seat, Cabin or Suite.  Note: Cabins and Suites are assumed to have the same number of seats as its Occupancy.")
					TempCheck = False
				End If
				
			Case WombTank
				If (InstallPoint = Suite) Or (InstallPoint = LuxurySuite) Or (InstallPoint = Cabin) Or (InstallPoint = LuxuryCabin) Or (InstallPoint = Hammock) Or (InstallPoint = Bunk) Or (InstallPoint = SmallGalley) Or (InstallPoint = CrampedCrewStation) Or (InstallPoint = NormalCrewStation) Or (InstallPoint = RoomyCrewStation) Or (InstallPoint = CrampedSeat) Or (InstallPoint = NormalSeat) Or (InstallPoint = RoomySeat) Or (InstallPoint = CrampedStandingRoom) Or (InstallPoint = NormalStandingRoom) Or (InstallPoint = RoomyStandingRoom) Then
					TempCheck = True
				Else
					modHelper.InfoPrint(1, "Womb Tanks must be added to a Bunks, Cabins, Suites,Hammocks, Galleys, Crew Station, Seat or Standing Room with the exception of Cycle and Harness versions.")
					TempCheck = False
				End If
			Case CrewEscapeCapsule
				If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = equipmentPod) Then
					TempCheck = True
				Else
					modHelper.InfoPrint(1, "Crew Escape Capsule must be placed in Body, Superstructure, Pod, Equipment Pod, Turret, Popturret, Arm, Wing, Open Mount or Leg.")
					TempCheck = False
				End If
				
				
			Case GravCompensator
				If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = equipmentPod) Or (InstallPoint = Module_Renamed) Then
					TempCheck = True
				Else
					modHelper.InfoPrint(1, "Grav Compensators must be placed in Body, Superstructure, Pod, Equipment Pod, Turret, Popturret, Arm, Wing, Open Mount, Leg or Module.")
					TempCheck = False
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
	
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		
		' set the default properties
		mvarCustom = False
		TL = gVehicleTL
		mvarQuantity = 1
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
		Select Case mvarDatatype
			Case EjectionSeat
				
			Case Airbag
				
			Case CrashWeb
				
			Case WombTank
				
			Case GravityWeb
				
			Case GravCompensator
			Case CrewEscapeCapsule
				
				
		End Select
		mvarOccupancy = 1
	End Sub
	
	Public Sub GetMatrixIndex()
		Dim i As Short
		If mvarDatatype = 0 Then Exit Sub
		mvarMatrixPos = 0 'init the counter
		For i = 1 To UBound(SafetyMatrix)
			If SafetyMatrix(i).ID = mvarDatatype Then
				If SafetyMatrix(i).TL >= mvarTL Then
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
		Dim ParentComponent As Short
		Dim OC As Integer
		Dim lngNumberofSystems As Integer
		Dim sPrint1 As String
		Dim sPrint2 As String
		Dim sPrintPlural As String
		Dim sPrintPlural2 As String
		Dim sPrintPlural3 As String
		Dim sPrintPlural4 As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ParentComponent = Veh.Components(mvarParent).Datatype
		
		mvarLocation = GetLocation
		
		If mvarMatrixPos = 0 Then Exit Sub
		
		If mvarDatatype = WombTank Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarWeight = SafetyMatrix(mvarMatrixPos).Weight * Veh.Components(Parent).Volume
			mvarCost = SafetyMatrix(mvarMatrixPos).Cost * mvarWeight
			mvarVolume = Weight / SafetyMatrix(mvarMatrixPos).Volume
			mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)
		ElseIf (ParentComponent = Cabin) Or (ParentComponent = LuxuryCabin) Or (ParentComponent = Suite) Or (ParentComponent = LuxurySuite) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarQuantity = Veh.Components(mvarParent).Occupancy * Veh.Components(mvarParent).Quantity
			mvarWeight = SafetyMatrix(mvarMatrixPos).Weight
			mvarCost = SafetyMatrix(mvarMatrixPos).Cost
			mvarVolume = SafetyMatrix(mvarMatrixPos).Volume
			mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)
		ElseIf mvarDatatype = CrewEscapeCapsule Then 
			mvarWeight = SafetyMatrix(mvarMatrixPos).Weight * mvarOccupancy
			mvarCost = SafetyMatrix(mvarMatrixPos).Cost * mvarOccupancy
			mvarVolume = SafetyMatrix(mvarMatrixPos).Volume * mvarOccupancy
			mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)
		ElseIf mvarDatatype = GravCompensator Then 
			mvarWeight = SafetyMatrix(mvarMatrixPos).Weight
			mvarCost = SafetyMatrix(mvarMatrixPos).Cost
			mvarVolume = SafetyMatrix(mvarMatrixPos).Volume
			mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarQuantity = Veh.Components(Parent).Quantity
			mvarWeight = SafetyMatrix(mvarMatrixPos).Weight
			mvarCost = SafetyMatrix(mvarMatrixPos).Cost
			mvarVolume = SafetyMatrix(mvarMatrixPos).Volume
			mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)
		End If
		
		'do final rounding
		mvarWeight = System.Math.Round(mvarWeight * mvarQuantity, 2)
		mvarCost = System.Math.Round(mvarCost * mvarQuantity, 2)
		mvarVolume = System.Math.Round(mvarVolume * mvarQuantity, 2)
		mvarSurfaceArea = System.Math.Round(mvarSurfaceArea * mvarQuantity, 2)
		
		mvarPowerReqt = SafetyMatrix(mvarMatrixPos).Power * mvarQuantity
		
		'find out G Reduction if gravcompensator
		Dim MaxG As Integer
		Dim BaseReduction As Double
		Dim NumCompensators As Integer
		Dim element As Object
		Dim TempReduction As Double
		Dim i As Short
		Dim Lweight As Double 'MPJ 07/07/00 increasd from single to double 'MPJ 07/07/00 increased from single to double 'MPJ 07/07/00 increased from long to double
		If mvarDatatype = GravCompensator Then
			
			'//for backward compatibility, make sure quantity is at least 1
			If mvarQuantity < 1 Then mvarQuantity = 1
			
			'determine maximum g reduction
			If mvarTL > 12 Then
				i = mvarTL - 12
				MaxG = 2 * 2 ^ i
				BaseReduction = 4000000 * 2 ^ i
			Else
				MaxG = 2
				BaseReduction = 4000000
			End If
			
			'find the number of compensators
			'todo: assuming i used recursion, i would send an iterator object to bring me
			'back a reference via:
			'set o = Body.Itterator.getObject(GravCompesator) <-- with that being a class ID
			
			For	Each element In Veh
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.Datatype = GravCompensator Then
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Quantity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					NumCompensators = NumCompensators + element.Quantity
				End If
			Next element
			
			'find the reduction
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Lweight = Veh.Stats.LoadedWeight
			If Lweight <= 0 Then Lweight = mvarWeight
			TempReduction = (BaseReduction * NumCompensators) / Lweight
			
			If TempReduction > MaxG Then
				mvarGReduction = MaxG
			Else
				mvarGReduction = System.Math.Round(TempReduction, 2)
			End If
		End If
		
		If mvarQuantity > 1 Then
			sPrintPlural = "s"
			sPrintPlural2 = " with total "
			sPrintPlural3 = " each"
			sPrintPlural4 = " total of "
		Else
			sPrintPlural = ""
			sPrintPlural2 = " with "
			sPrintPlural3 = ""
			sPrintPlural4 = ""
		End If
		
		'print output
		If mvarGReduction <> 0 Then
			sPrint1 = ", " & VB6.Format(GReduction) & " G reduction"
		End If
		
		mvarPrintOutput = NumericToString(mvarQuantity) & " TL" & mvarTL & " " & mvarCustomDescription & sPrintPlural & sPrint1 & " (" & mvarLocation & ", HP " & mvarHitPoints & sPrintPlural3 & ", " & sPrintPlural4 & VB6.Format(mvarWeight, p_sFormat) & " lbs., " & VB6.Format(mvarVolume, p_sFormat) & " cf., " & "$" & VB6.Format(mvarCost, p_sFormat) & ", " & VB6.Format(mvarPowerReqt, p_sFormat) & " kW " & sPrint2 & ")." & mvarComment
	End Sub
	
	Public Sub QueryParent()
		' if the object has a parent, query it and check to see if
		' more stats/property updates are needed for other objects in the collection
		If mvarParent <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Veh.Components(Parent).StatsUpdate()
		End If
	End Sub
End Class