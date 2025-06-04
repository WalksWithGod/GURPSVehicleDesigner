Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsWeaponLauncher_NET.clsWeaponLauncher")> Public Class clsWeaponLauncher
	
	'local variable(s) to hold property value(s)
	Private mvarHitPoints As Double
	Private mvarDR As Integer
	Private mvarCustom As Boolean
	Private mvarCost As Double
	Private mvarDatatype As Short
	Private mvarDescription As String
	Private mvarImage As Short
	Private mvarSelectedImage As Short
	Private mvarKey As String
	Private mvarParent As String
	Private mvarQuantity As Short
	Private mvarSurfaceArea As Double
	Private mvarTL As Short
	Private mvarVolume As Double
	Private mvarRuggedized As Boolean
	Private mvarWeight As Double
	Private mvarCustomDescription As String
	Private mvarQuality As String
	Private mvarDiameter As Single
	Private mvarMaxLoad As Single
	Private mvarSnapShot As Integer
	Private mvarLoaders As Integer
	Private mvarRoF As String
	Private mvarShots As String
	Private mvarCylinders As Integer
	Private mvarMount As String
	Private mvarDirection As String
	Private mvarLocation As String
	Private mvarComment As String
	Private mvarCName As String
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
	
	
	
	
	Public Property Mount() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Mount
			Mount = mvarMount
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Mount = 5
			mvarMount = Value
		End Set
	End Property
	
	
	
	Public Property Cylinders() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Cylinders
			Cylinders = mvarCylinders
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Cylinders = 5
			mvarCylinders = Value
		End Set
	End Property
	
	
	
	Public Property Shots() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Shots
			Shots = mvarShots
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Shots = 5
			mvarShots = Value
		End Set
	End Property
	
	
	
	
	
	Public Property rof() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.RoF
			rof = mvarRoF
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.RoF = 5
			mvarRoF = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Loaders() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Loaders
			Loaders = mvarLoaders
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Loaders = 5
			mvarLoaders = Value
		End Set
	End Property
	
	
	
	
	
	Public Property SnapShot() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.SnapShot
			SnapShot = mvarSnapShot
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.SnapShot = 5
			mvarSnapShot = Value
		End Set
	End Property
	
	
	
	
	
	Public Property MaxLoad() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.MaxLoad
			MaxLoad = mvarMaxLoad
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.MaxLoad = 5
			mvarMaxLoad = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Diameter() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Diameter
			Diameter = mvarDiameter
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Diameter = 5
			
			mvarDiameter = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Quality() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Quality
			Quality = mvarQuality
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Quality = 5
			mvarQuality = Value
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
	
	
	
	Public Property Direction() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Direction
			Direction = mvarDirection
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Direction = 5
			mvarDirection = Value
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
	
	
	
	
	Public Function LocationCheck() As Boolean
		Dim TempCheck As Boolean
		Dim InstallPoint As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallPoint = Veh.Components(mvarParent).Datatype
		
		
		If (InstallPoint = HardPoint) Or (InstallPoint = WeaponBay) Or (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = equipmentPod) Or (InstallPoint = Module_Renamed) Then
			TempCheck = True
		Else
			modHelper.InfoPrint(1, "Launchers must be placed in Body, Superstructure, Pod, equipment Pod,Turret, Popturret, Arm, Wing, Open Mount, Leg, Hardpoint, Weaponbay or Module.")
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
		mvarCustom = False
		TL = gVehicleTL
		mvarQuantity = 1
		mvarQuality = "normal"
		mvarMount = "normal"
		mvarMaxLoad = 1000
		mvarDirection = "front"
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
			Case MuzzleloadingLauncher
				
				mvarDiameter = 30
				mvarCylinders = 1
			Case BreechloadingLauncher
				
				mvarDiameter = 30
				mvarCylinders = 1
			Case ManualRepeaterLauncher
				
				mvarDiameter = 30
				mvarCylinders = 1
			Case RevolverLauncher
				
				mvarCylinders = 5
				mvarDiameter = 30
			Case SlowAutoLoaderLauncher
				
				mvarDiameter = 30
				mvarCylinders = 1
			Case FastAutoLoaderLauncher
				
				mvarDiameter = 30
				mvarCylinders = 1
			Case lightAutomaticLauncher
				
				mvarDiameter = 30
				mvarCylinders = 1
			Case HeavyAutomaticLauncher
				
				mvarDiameter = 30
				mvarCylinders = 1
			Case DisposableLauncher
				
				mvarDiameter = 30
				mvarCylinders = 1
		End Select
		
	End Sub
	
	Public Sub StatsUpdate()
		Dim sPrint1 As String
		Dim sPrint2 As String
		Dim sPrintPlural As String
		Dim sPrintPlural2 As String
		Dim sPrintPlural3 As String
		Dim sPrintDirection As String
		Dim QRugMod As Single
		Dim RugHitMod As Integer
		
		
		mvarZZInit = 1
		
		mvarLocation = GetLocation
		
		'set the ruggedized and quantity multipliers
		If mvarQuantity < 1 Then mvarQuantity = 1
		If mvarRuggedized Then
			QRugMod = 1.5 * mvarQuantity
			RugHitMod = 2
		Else
			QRugMod = 1 * mvarQuantity
			RugHitMod = 1
		End If
		
		'calculate statistics
		mvarWeight = GetWeight
		mvarVolume = GetVolume
		mvarSnapShot = GetSnapShot
		mvarRoF = GetRoF
		mvarCost = GetCost
		mvarShots = GetShots
		mvarLoaders = GetLoaders
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)
		
		'cost, malf, and accuracy modifiers for Cheap, Fine and Very Fine quality are calced in the functions below
		'todo check that "mounting" option is taken into account
		'todo check that "recoiless" option is taken into account
		'todo check that advancedoption is taken into account
		
		'//update the cost,weight,volume, surface area and volume based on quantity and ruggedized options
		mvarCost = mvarCost * QRugMod
		mvarWeight = mvarWeight * QRugMod
		mvarVolume = mvarVolume * QRugMod
		
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		mvarLoaders = mvarLoaders * mvarQuantity
		
		'produce the print output
		If mvarRuggedized Then
			sPrint1 = "ruggedized "
		Else
			sPrint1 = ""
		End If
		
		
		'produce the print output
		If mvarMount <> "normal" Then
			sPrint2 = sPrint2 & ", " & mvarMount
		End If
		
		sPrint2 = sPrint2 & ", " & VB6.Format(mvarMaxLoad, p_sFormat) & " lbs max load"
		
		If mvarQuality <> "normal" Then sPrint2 = sPrint2 & ", " & mvarQuality & " construction"
		
		
		sPrintDirection = StrConv(Left(mvarDirection, 1), VbStrConv.UpperCase)
		
		If mvarQuantity > 1 Then
			sPrintPlural = "s"
			sPrintPlural2 = " each"
			sPrintPlural3 = " total of "
		Else
			sPrintPlural = ""
			sPrintPlural2 = ""
			sPrintPlural3 = ""
		End If
		
		mvarPrintOutput = NumericToString(mvarQuantity) & " TL" & mvarTL & " " & sPrint1 & " " & mvarDiameter & "mm " & mvarCustomDescription & sPrintPlural & sPrint2 & " (" & mvarLocation & sPrintDirection & ", HP " & mvarHitPoints & sPrintPlural2 & ", " & sPrintPlural3 & VB6.Format(mvarWeight, p_sFormat) & " lbs., " & VB6.Format(mvarVolume, p_sFormat) & " cf., " & "$" & VB6.Format(mvarCost, p_sFormat) & ")." & mvarComment
		
	End Sub
	
	Public Sub QueryParent()
		' if the object has a parent, query it and check to see if
		' more stats/property updates are needed for other objects in the collection
		If mvarParent <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Veh.Components(mvarParent).StatsUpdate()
		End If
	End Sub
	
	
	Private Function GetWeight() As Single
		Dim mWPS As Single
		Dim R As Single
		Dim T As Single
		Dim M As Single
		Dim TempWeight As Single
		
		'get mWPS
		mWPS = mvarMaxLoad
		
		'get R
		If mvarDatatype = DisposableLauncher Then
			R = 1
		ElseIf mvarDatatype = MuzzleloadingLauncher Then 
			R = 1.5
		ElseIf mvarDatatype = BreechloadingLauncher Then 
			R = 2
		ElseIf (mvarDatatype = ManualRepeaterLauncher) Or (mvarDatatype = SlowAutoLoaderLauncher) Then 
			R = 2.5
		ElseIf mvarDatatype = RevolverLauncher Then 
			R = 1.9 + (0.1 * mvarCylinders)
		ElseIf mvarDatatype = FastAutoLoaderLauncher Then 
			R = 3
		ElseIf mvarDatatype = lightAutomaticLauncher Then 
			R = 3.5
		ElseIf mvarDatatype = HeavyAutomaticLauncher Then 
			R = 4.5
		End If
		
		' get T
		If mvarTL <= 7 Then
			T = 0
		Else
			T = 0.5
		End If
		
		'get M
		If mvarCylinders = 1 Then
			M = 1
		ElseIf mvarCylinders > 1 Then 
			M = 1 + ((mvarCylinders - 1) * 0.4)
		End If
		
		TempWeight = (R - T) * mWPS * M
		
		GetWeight = System.Math.Round(TempWeight, 2)
		
	End Function
	
	Private Function GetVolume() As Single
		If mvarMount = "normal" Then
			GetVolume = mvarWeight / 50
		Else
			GetVolume = mvarWeight / 20 'concealed weapons take up more space
		End If
		
	End Function
	
	Private Function GetSnapShot() As Integer
		Dim TSS As Integer
		
		If mvarWeight < 2.5 Then
			TSS = 11
		ElseIf mvarWeight < 10 Then 
			TSS = 12
		ElseIf mvarWeight < 15 Then 
			TSS = 14
		ElseIf mvarWeight < 26 Then 
			TSS = 17
		ElseIf mvarWeight < 401 Then 
			TSS = 20
		ElseIf mvarWeight < 2001 Then 
			TSS = 25
		Else
			TSS = 30
		End If
		
		GetSnapShot = TSS
	End Function
	
	Private Function GetRoF() As String
		Dim TempRoF As String
		Dim iRof As Single
		Dim Numerator As String
		
		iRof = 0
		
		Select Case mvarDatatype
			Case MuzzleloadingLauncher
				If mvarCylinders > 1 Then
					Numerator = Str(mvarCylinders) & ":"
				Else
					Numerator = "1/"
				End If
				
				If mvarDiameter < 120 Then
					If mvarTL <= 3 Then
						TempRoF = Numerator & "45"
					ElseIf mvarTL = 4 Then 
						TempRoF = Numerator & "30"
					ElseIf mvarTL = 5 Then 
						TempRoF = Numerator & "10"
					ElseIf mvarTL >= 6 Then 
						TempRoF = Numerator & "6"
					End If
				Else
					If mvarTL <= 3 Then
						iRof = System.Math.Round(mvarDiameter / 2.66, 0) 'note: i round the iROF here
					ElseIf mvarTL = 4 Then 
						iRof = mvarDiameter / 4
					ElseIf mvarTL = 5 Then 
						iRof = System.Math.Round(mvarDiameter / 12, 0)
					Else
						iRof = System.Math.Round(mvarDiameter / 20, 0)
					End If
					TempRoF = Numerator & Str(iRof)
					iRof = 0
				End If
				
			Case DisposableLauncher
				If mvarCylinders = 1 Then
					TempRoF = "1NR"
				Else
					TempRoF = mvarCylinders & ":1NR"
				End If
				iRof = 0
			Case BreechloadingLauncher
				If mvarCylinders > 1 Then
					Numerator = Str(mvarCylinders) & ":"
				Else
					Numerator = "1/"
				End If
				
				If mvarDiameter <= 60 Then
					If mvarTL <= 3 Then
						TempRoF = Numerator & "20"
					ElseIf mvarTL = 4 Then 
						TempRoF = Numerator & "10"
					ElseIf mvarTL = 5 Then 
						TempRoF = Numerator & "5"
					ElseIf mvarTL >= 6 Then 
						TempRoF = Numerator & "2"
					End If
				Else
					If mvarTL <= 3 Then
						iRof = System.Math.Round(mvarDiameter / 3, 0)
					ElseIf mvarTL = 4 Then 
						iRof = System.Math.Round(mvarDiameter / 6, 0)
					ElseIf mvarTL = 5 Then 
						iRof = System.Math.Round(mvarDiameter / 12, 0)
					Else
						iRof = System.Math.Round(mvarDiameter / 30, 0)
					End If
					TempRoF = Numerator & Str(iRof)
					iRof = 0
				End If
				
			Case RevolverLauncher
				
				If mvarDiameter < 40 Then
					If mvarTL <= 5 Then
						TempRoF = "1"
					ElseIf mvarTL >= 6 Then 
						TempRoF = "3"
					End If
				Else
					If mvarTL <= 5 Then
						TempRoF = "1/2"
					Else
						TempRoF = CStr(1)
					End If
				End If
				
			Case ManualRepeaterLauncher
				If mvarTL <= 5 Then
					TempRoF = "1"
				Else
					TempRoF = "2"
				End If
				
			Case SlowAutoLoaderLauncher
				If mvarDiameter <= 40 Then
					TempRoF = "1"
				Else
					iRof = System.Math.Round(mvarDiameter / 40, 0)
					TempRoF = CStr(iRof)
					If iRof = 1 Then
						iRof = 0
					End If
					
				End If
				
			Case FastAutoLoaderLauncher
				If mvarDiameter <= 15 Then
					TempRoF = "3"
				ElseIf mvarDiameter <= 20 Then 
					TempRoF = "2"
				ElseIf mvarDiameter <= 60 Then 
					TempRoF = "1"
				Else
					iRof = System.Math.Round(mvarDiameter / 60, 0)
					TempRoF = CStr(iRof)
					If iRof = 1 Then
						iRof = 0
					End If
				End If
				
			Case lightAutomaticLauncher
				TempRoF = "Up to " & Str(CInt(160 / mvarDiameter))
				
			Case HeavyAutomaticLauncher
				If mvarDiameter <= 20 Then
					TempRoF = "3 to 20"
				Else
					TempRoF = "Up to " & Str(System.Math.Round(400 / mvarDiameter, 0))
				End If
		End Select
		
		
		'pass final results
		If iRof <> 0 Then
			GetRoF = "1/" & TempRoF
		Else
			GetRoF = TempRoF
		End If
	End Function
	
	Private Function GetCost() As Single
		Dim TempCost As Single
		
		
		'get weight
		If mvarWeight < 10 Then
			TempCost = (50 * mvarWeight) + 250
		ElseIf mvarWeight <= 100 Then 
			TempCost = 75 * mvarWeight
		ElseIf mvarWeight > 100 Then 
			TempCost = (25 * mvarWeight) + 5000
		End If
		
		'get tl  modifier
		If mvarTL <= 5 Then
			TempCost = TempCost / 10
		ElseIf mvarTL = 6 Then 
			TempCost = TempCost / 5
		ElseIf mvarTL = 9 Then 
			TempCost = TempCost / 2
		ElseIf mvarTL >= 10 Then 
			TempCost = TempCost / 4
		End If
		
		'double cost for automatic or divide by 10 if disposable
		If (mvarDatatype = HeavyAutomaticLauncher) Or (mvarDatatype = lightAutomaticLauncher) Then
			TempCost = TempCost * 2
		ElseIf mvarDatatype = DisposableLauncher Then 
			TempCost = TempCost / 10
		End If
		
		'get modifier for Cheap, Fine and Very Fine quality
		If mvarQuality = "cheap" Then
			TempCost = TempCost / 2
		ElseIf mvarQuality = "fine (accurate)" Then 
			TempCost = TempCost * 5
		ElseIf mvarQuality = "very fine (accurate)" Then 
			TempCost = TempCost * 30
		ElseIf mvarQuality = "fine (reliable)" Then 
			TempCost = TempCost * 5
		End If
		
		GetCost = TempCost
	End Function
	
	
	Private Function GetShots() As String
		'number of shots the weapon has ready to fire.
		If (mvarDatatype = DisposableLauncher) Or (mvarDatatype = RevolverLauncher) Or (mvarDatatype = BreechloadingLauncher) Or (mvarDatatype = MuzzleloadingLauncher) Then
			GetShots = Str(mvarCylinders)
		Else
			GetShots = "var."
		End If
	End Function
	
	Private Function GetLoaders() As Integer
		Dim TempLoaders As Single
		
		If (mvarDatatype = MuzzleloadingLauncher) Or (mvarDatatype = BreechloadingLauncher) Or (mvarDatatype = RevolverLauncher) Or (mvarDatatype = ManualRepeaterLauncher) Then
			TempLoaders = (mvarDiameter / 250) - 1
			'TODO rules (page 122 top) say round to nearest whole  number.  Im rounding up, need to decide if that's final
			GetLoaders = RoundUP(TempLoaders)
		Else
			GetLoaders = 0
		End If
		
	End Function
End Class