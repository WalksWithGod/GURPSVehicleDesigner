Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsFuelTank_NET.clsFuelTank")> Public Class clsFuelTank
	
	Private mvarTL As Short
	Private mvarWeight As Double
	Private mvarVolume As Double
	Private mvarCost As Double
	Private mvarFire As Short
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
	Private mvarCapacity As Single
	Private mvarFailSafePoints As Integer
	Private mvarFuelWeight As Single
	Private mvarFuelCost As Single
	Private mvarFuelType As Short
	Private mvarFuelFire As Short
	Private mvarFuel As String 'this is the string that appears in property's dialog for user
	
	Private mvarImage As Short
	Private mvarSelectedImage As Short
	Private mvarComment As String
	Private mvarCName As String
	Private mvarMatrixPos As Integer
	Private mvarMatrixPos2 As Integer
	
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
	
	
	
	
	Public Property MatrixPos2() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.MatrixPos2
			MatrixPos2 = mvarMatrixPos2
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.MatrixPos2 = 5
			mvarMatrixPos2 = Value
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
	
	
	
	Public Property Fuel() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Fuel
			Fuel = mvarFuel
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Fuel = 5
			mvarFuel = Value
			If mvarZZInit = 0 Then Exit Property
			
			If mvarFuel = "ethanol" Then
				mvarFuelType = EthanolAlchohol
			ElseIf mvarFuel = "methanol" Then 
				mvarFuelType = MethanolAlchohol
			ElseIf mvarFuel = "aviation gas" Then 
				mvarFuelType = AviationGas
			ElseIf mvarFuel = "cadmium" Then 
				mvarFuelType = Cadmium
			ElseIf mvarFuel = "diesel" Then 
				mvarFuelType = Diesel
			ElseIf mvarFuel = "gasoline" Then 
				mvarFuelType = Gasoline
			ElseIf mvarFuel = "jet fuel" Then 
				mvarFuelType = JetFuel
			ElseIf mvarFuel = "rocket fuel" Then 
				mvarFuelType = RocketFuel
			ElseIf mvarFuel = "water" Then 
				mvarFuelType = Water
			ElseIf mvarFuel = "hydrogen" Then 
				mvarFuelType = LiquidHydrogen
			ElseIf mvarFuel = "metal/LOX" Then 
				mvarFuelType = MetalLOX
			ElseIf mvarFuel = "oxygen (LOX)" Then 
				mvarFuelType = LiquidOxygen
			ElseIf mvarFuel = "propane/LNG" Then 
				mvarFuelType = Propane
			End If
			
			GetMatrixIndex()
		End Set
	End Property
	
	
	
	
	Public Property FuelType() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.FuelType
			FuelType = mvarFuelType
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.FuelType = 5
			mvarFuelType = Value
			GetMatrixIndex()
		End Set
	End Property
	
	
	
	Public Property Fire() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Fire
			Fire = mvarFire
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Fire = 5
			mvarFire = Value
		End Set
	End Property
	
	
	
	Public Property FuelFire() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.FuelFire
			FuelFire = mvarFuelFire
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.FuelFire = 5
			mvarFuelFire = Value
		End Set
	End Property
	
	
	
	Public Property FuelCost() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.FuelCost
			FuelCost = mvarFuelCost
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.FuelCost = 5
			mvarFuelCost = Value
		End Set
	End Property
	
	
	
	
	Public Property FuelWeight() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.FuelWeight
			FuelWeight = mvarFuelWeight
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.FuelWeight = 5
			mvarFuelWeight = Value
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
	
	
	
	Public Property capacity() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Capacity
			capacity = mvarCapacity
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Capacity = 5
			mvarCapacity = Value
		End Set
	End Property
	
	
	
	Public Property FailSafePoints() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.FailSafePoints
			FailSafePoints = mvarFailSafePoints
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.FailSafePoints = 5
			mvarFailSafePoints = Value
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
		
		If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = equipmentPod) Then
			TempCheck = True
		Else
			modHelper.InfoPrint(1, "Fuel Tanks must be placed in Body, Superstructure, Pod, equipment Pod, Turret, Popturret, Arm, Wing, Open Mount or Leg.")
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
		Dim mvarFuelUsingSystemKeyChain(1) As Object
		
		' set the default properties
		mvarCustom = False
		TL = gVehicleTL
		mvarRuggedized = False
		mvarCapacity = 1000
		
		
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
			Case CoalBunker
				
				mvarCapacity = 100
				mvarFuelType = Coal
			Case WoodBunker
				
				mvarCapacity = 100
				mvarFuelType = Wood
			Case StandardTank
				
				mvarFuel = "gasoline"
				mvarFuelType = Gasoline
				mvarCapacity = 10
			Case lightTank
				
				mvarFuel = "gasoline"
				mvarFuelType = Gasoline
				mvarCapacity = 20
			Case StandardSelfSealingTank
				
				mvarFuel = "gasoline"
				mvarFuelType = Gasoline
				mvarCapacity = 20
			Case UltralightTank
				
				mvarFuel = "gasoline"
				mvarFuelType = Gasoline
				mvarCapacity = 20
			Case lightSelfSealingTank
				
				mvarFuel = "gasoline"
				mvarFuelType = Gasoline
				mvarCapacity = 20
			Case UltralightSelfSealingTank
				
				mvarCapacity = 20
				mvarFuel = "gasoline"
				mvarFuelType = Gasoline
			Case AntiMatterBay
				
				mvarFailSafePoints = 0
				mvarCapacity = 1
				mvarFuelType = AntiMatter
		End Select
		
		
	End Sub
	
	Public Sub GetMatrixIndex()
		Dim i As Short
		If mvarDatatype = 0 Then Exit Sub
		'first load the matrix for the Storage Tanks
		mvarMatrixPos = 0 'init the counter
		For i = 1 To UBound(FuelTankMatrix)
			If FuelTankMatrix(i).ID = mvarDatatype Then
				If FuelTankMatrix(i).TL >= mvarTL Then
					mvarMatrixPos = i
					Exit For
				Else
					mvarMatrixPos = i
				End If
			End If
		Next 
		
		
		'Now load the matrix for the fuel itself
		mvarMatrixPos2 = 0 'init the counter
		For i = 1 To UBound(FuelMatrix)
			If FuelMatrix(i).ID = mvarFuelType Then
				If FuelMatrix(i).TL >= mvarTL Then
					mvarMatrixPos2 = i
					Exit For
				Else
					mvarMatrixPos2 = i
				End If
			End If
		Next 
		
	End Sub
	
	
	Public Sub StatsUpdate()
		mvarZZInit = 1
		If mvarMatrixPos = 0 Then Exit Sub
		
		Dim TrueWeight As Single
		Dim TempCost As Single
		Dim TempFuel As Single
		Dim TempVolume As Single
		Dim QRugMod As Single 'combined quantity and ruggedized multipliers
		Dim RugHitMod As Short 'ruggedized hit point multiplier
		Dim sPrint1 As String
		Dim sPrint2 As String
		Dim sPrint3 As String
		Dim sPrintPlural As String
		Dim sPrintPlural2 As String
		Dim sPrintPlural3 As String
		Dim sPrintPlural4 As String
		
		mvarLocation = GetLocation
		
		'set the ruggedized and quantity multipliers
		If mvarRuggedized Then
			QRugMod = 1.5
			RugHitMod = 2
		Else
			QRugMod = 1
			RugHitMod = 1
		End If
		
		'determine if the weight is above or below 5kw and then make adjustments
		TrueWeight = mvarCapacity * FuelTankMatrix(mvarMatrixPos).Weight
		
		'Find the volume
		'If (mvarDatatype = CoalBunker) Or (mvarDatatype = WoodBunker) Then
		'    TempVolume = mvarCapacity
		'Else 'MPJ 07/07/00  Coal bunkers and Wood bunkers also need to use the
		'data files.  Volume is not a 1:1 relationship
		
		TempVolume = mvarCapacity * FuelTankMatrix(mvarMatrixPos).Volume
		'End If
		
		'find the cost
		TempCost = mvarCapacity * FuelTankMatrix(mvarMatrixPos).Cost
		
		'calc stats for Failsafes NOTE: ive placed failsafes before ruggedized calcs
		If (mvarDatatype = AntiMatterBay) And (mvarFailSafePoints > 0) Then
			TempVolume = TempVolume * mvarFailSafePoints
			TempCost = TempCost * mvarFailSafePoints
			TrueWeight = TrueWeight * mvarFailSafePoints
		End If
		
		'get base stats
		mvarWeight = TrueWeight
		mvarCost = TempCost
		mvarVolume = TempVolume
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
		
		'get finals
		mvarWeight = System.Math.Round(QRugMod * mvarWeight, 2)
		mvarCost = System.Math.Round(QRugMod * mvarCost, 2)
		mvarVolume = System.Math.Round(QRugMod * mvarVolume, 2)
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		
		'Get fire
		mvarFire = FuelTankMatrix(mvarMatrixPos).Fire
		
		'Get the cost for the fuel based on the fueltype
		'NOTE: These should not be added to the actual cost of the tank. Thats why its last!
		mvarFuelWeight = FuelMatrix(mvarMatrixPos2).Weight * mvarCapacity
		mvarFuelCost = FuelMatrix(mvarMatrixPos2).Cost * mvarCapacity
		mvarFuelFire = FuelMatrix(mvarMatrixPos2).Fire
		
		'produce the print output
		If mvarRuggedized Then
			sPrint1 = "ruggedized "
		Else
			sPrint1 = ""
		End If
		
		Select Case mvarDatatype
			Case CoalBunker
				sPrint1 = VB6.Format(mvarCapacity, p_sFormat) & " cf. " & sPrint1
				sPrint2 = " Holds " & VB6.Format(mvarCapacity, p_sFormat) & " cf. coal" & " (" & VB6.Format(mvarFuelWeight, p_sFormat) & " lbs)"
			Case WoodBunker
				sPrint1 = VB6.Format(mvarCapacity, p_sFormat) & " cf. " & sPrint1
				sPrint2 = " Holds " & VB6.Format(mvarCapacity, p_sFormat) & " cf. wood" & " (" & VB6.Format(mvarFuelWeight, p_sFormat) & " lbs)"
			Case StandardTank, lightTank, StandardSelfSealingTank, UltralightTank, lightSelfSealingTank, UltralightSelfSealingTank
				sPrint1 = VB6.Format(mvarCapacity, p_sFormat) & " gal. " & sPrint1
				sPrint2 = " Holds " & VB6.Format(mvarCapacity, p_sFormat) & " gal. " & mvarFuel & " (" & VB6.Format(mvarFuelWeight, p_sFormat) & " lbs, fire +" & VB6.Format(mvarFuelFire) & ")"
				sPrint3 = ", fire " & VB6.Format(mvarFire)
			Case AntiMatterBay
				sPrint1 = VB6.Format(mvarCapacity, p_sFormat) & " gram " & sPrint1
				sPrint3 = ", fail safe " & VB6.Format(mvarFailSafePoints)
		End Select
		
		mvarPrintOutput = "TL" & mvarTL & " " & sPrint1 & mvarCustomDescription & " (" & mvarLocation & ", HP " & mvarHitPoints & ", " & VB6.Format(mvarWeight, p_sFormat) & " lbs., " & VB6.Format(mvarVolume, p_sFormat) & " cf., " & "$" & VB6.Format(mvarCost, p_sFormat) & sPrint3 & ")." & sPrint2 & mvarComment
		
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