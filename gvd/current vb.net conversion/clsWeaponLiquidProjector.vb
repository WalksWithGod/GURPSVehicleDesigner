Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsWeaponLiquidProjector_NET.clsWeaponLiquidProjector")> Public Class clsWeaponLiquidProjector
	
	Private mvarHitPoints As Double
	Private mvarDR As Integer
	Private mvarQuality As String
	Private mvarCustom As Boolean
	Private mvarCost As Double
	Private mvarDatatype As Short
	Private mvarDescription As String
	Private mvarImage As Short
	Private mvarSelectedImage As Short
	Private mvarRuggedized As Boolean
	Private mvarKey As String
	Private mvarParent As String
	Private mvarQuantity As Short
	Private mvarSurfaceArea As Double
	Private mvarTL As Short
	Private mvarVolume As Double
	Private mvarWeight As Double
	Private mvarWPS As Single
	Private mvarCPS As Single
	Private mvarCustomDescription As String
	Private mvarDamage As String
	Private mvarTypeDamage As String
	Private mvarhalfDamage As Double
	Private mvarMaxRange As Double
	Private mvarAccuracy As Integer
	Private mvarSnapShot As Integer
	Private mvarShots As Integer
	Private mvarRoF As String
	Private mvarPowerReqt As Double
	Private mvarMount As String
	Private mvarStyle As String
	Private mvarAmmunitionType As String
	Private mvarLoaders As Integer
	Private mvarMalfunction As String
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
	
	
	Public Property Style() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Style
			Style = mvarStyle
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Style = 5
			mvarStyle = Value
		End Set
	End Property
	
	
	
	Public Property AmmunitionType() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.AmmunitionType
			AmmunitionType = mvarAmmunitionType
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.AmmunitionType = 5
			mvarAmmunitionType = Value
		End Set
	End Property
	
	
	Public Property Malfunction() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Malfunction
			Malfunction = mvarMalfunction
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Malfunction = 5
			mvarMalfunction = Value
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
	
	
	
	
	
	Public Property Shots() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Shots
			Shots = mvarShots
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Shots = 5
			mvarShots = Value
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
	
	
	
	
	
	Public Property Accuracy() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Accuracy
			Accuracy = mvarAccuracy
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Accuracy = 5
			mvarAccuracy = Value
		End Set
	End Property
	
	
	
	
	
	Public Property MaxRange() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.MaxRange
			MaxRange = mvarMaxRange
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.MaxRange = 5
			mvarMaxRange = Value
		End Set
	End Property
	
	
	
	
	
	Public Property halfDamage() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.halfDamage
			halfDamage = mvarhalfDamage
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.halfDamage = 5
			mvarhalfDamage = Value
		End Set
	End Property
	
	
	
	
	
	Public Property TypeDamage() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.TypeDamage
			TypeDamage = mvarTypeDamage
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.TypeDamage = 5
			mvarTypeDamage = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Damage() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Damage
			Damage = mvarDamage
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Damage = 5
			mvarDamage = Value
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
	
	
	
	
	
	Public Property CPS() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.CPS
			CPS = mvarCPS
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.CPS = 5
			mvarCPS = Value
		End Set
	End Property
	
	
	
	Public Property WPS() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.WPS
			WPS = mvarWPS
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.WPS = 5
			mvarWPS = Value
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
		
		
		If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = equipmentPod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = Module_Renamed) Then
			TempCheck = True
		Else
			modHelper.InfoPrint(1, "Liquid Projectors must be placed in Body, Superstructure, Pod, equipment Pod,Turret, Popturret, Arm, Wing, Open Mount, Leg or Module.")
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
		mvarTL = gVehicleTL
		mvarQuantity = 1
		mvarQuality = "normal"
		mvarMount = "normal"
		mvarStyle = "medium"
		mvarTypeDamage = "Spcl."
		mvarShots = 10
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
			Case FlameThrower
				
			Case WaterCannon
				
				mvarAmmunitionType = "water"
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
		'UPGRADE_WARNING: Couldn't resolve default property of object ConvertDamage(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarDamage = ConvertDamage(GetDamage)
		mvarMalfunction = GetMalfunction
		mvarLoaders = GetLoaders
		mvarSnapShot = GetSnapShot
		mvarAccuracy = GetAccuracy
		mvarhalfDamage = GetHalfDamage
		mvarMaxRange = GetMaxRange
		mvarWeight = GetWeight
		mvarVolume = GetVolume
		mvarRoF = GetRoF
		mvarCost = GetCost
		mvarWPS = GetWPS
		mvarCPS = GetCPS
		mvarSurfaceArea = CalcSurfaceArea(Volume)
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarHitPoints = CalcComponentHitpoints(SurfaceArea) * RugHitMod
		'cost, malf, and accuracy modifiers for Cheap, Fine and Very Fine quality are calced in the functions below
		
		'//update the cost,weight,volume, surface area and volume based on quantity and ruggedized options
		mvarCost = mvarCost * QRugMod
		mvarWeight = mvarWeight * QRugMod
		mvarVolume = mvarVolume * QRugMod
		
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		mvarPowerReqt = mvarPowerReqt * mvarQuantity
		
		'produce the print output
		If mvarRuggedized Then
			sPrint1 = "ruggedized "
		Else
			sPrint1 = ""
		End If
		
		sPrint1 = sPrint1 & mvarStyle & " "
		
		If mvarMount <> "normal" Then
			sPrint2 = sPrint2 & ", " & mvarMount
		End If
		
		If mvarDatatype = WaterCannon Then
			sPrint2 = sPrint2 & ", fires " & mvarAmmunitionType
		End If
		
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
		
		mvarPrintOutput = NumericToString(mvarQuantity) & " TL" & mvarTL & " " & sPrint1 & mvarCustomDescription & sPrintPlural & sPrint2 & " (" & mvarLocation & sPrintDirection & ", HP " & mvarHitPoints & sPrintPlural2 & ", " & sPrintPlural3 & VB6.Format(mvarWeight, p_sFormat) & " lbs., " & VB6.Format(mvarVolume, p_sFormat) & " cf., " & "$" & VB6.Format(mvarCost, p_sFormat) & ", " & VB6.Format(mvarPowerReqt, p_sFormat) & " kW)." & mvarComment
		
		
	End Sub
	
	Public Sub QueryParent()
		' if the object has a parent, query it and check to see if
		' more stats/property updates are needed for other objects in the collection
		If mvarParent <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Veh.Components(Parent).StatsUpdate()
		End If
	End Sub
	
	
	
	Private Function GetMalfunction() As String
		Dim TempMalf As String
		
		
		If mvarTL <= 3 Then
			TempMalf = CStr(14)
		ElseIf mvarTL = 4 Then 
			TempMalf = CStr(15)
		ElseIf mvarTL = 5 Then 
			TempMalf = CStr(16)
		ElseIf mvarTL >= 6 Then 
			TempMalf = "Crit."
		End If
		
		'get modifier for Cheap, Fine and Very Fine quality
		If mvarQuality = "cheap" Then
			TempMalf = DecreaseMalf(TempMalf)
		ElseIf mvarQuality = "fine (reliable)" Then 
			TempMalf = IncreaseMalf(TempMalf)
		End If
		
		GetMalfunction = TempMalf
	End Function
	
	
	Private Function GetDamage() As Single
		Dim TDam As Single
		
		If mvarDatatype = FlameThrower Then
			If mvarTL <= 4 Then
				TDam = 1
			ElseIf mvarTL = 5 Then 
				TDam = 2
			ElseIf mvarTL <= 8 Then 
				TDam = 3
			ElseIf mvarTL >= 9 Then 
				TDam = 5
			End If
		ElseIf mvarDatatype = WaterCannon Then 
			If mvarAmmunitionType = "acid" Then
				TDam = 0.8 'this translates into 1d-1 damage
			ElseIf mvarTL <= 4 Then 
				TDam = 2
			ElseIf mvarTL = 5 Then 
				TDam = 3
			ElseIf mvarTL <= 8 Then 
				TDam = 4
			ElseIf mvarTL >= 9 Then 
				TDam = 4
			End If
		End If
		
		GetDamage = TDam
		
	End Function
	
	Private Function GetHalfDamage() As Single 'in yards
		Dim TempHD As Single
		
		If mvarStyle = "light" Then
			TempHD = mvarTL * 5
		ElseIf mvarStyle = "medium" Then 
			TempHD = mvarTL * 7
		ElseIf mvarStyle = "heavy" Then 
			TempHD = mvarTL * 10
		End If
		
		If mvarTL <= 4 Then TempHD = CInt(TempHD / 10)
		
		GetHalfDamage = TempHD
	End Function
	
	Private Function GetMaxRange() As Double 'in yards
		Dim TempMax As Single
		
		If mvarStyle = "light" Then
			TempMax = mvarTL * 7
		ElseIf mvarStyle = "medium" Then 
			TempMax = mvarTL * 10
		ElseIf mvarStyle = "heavy" Then 
			TempMax = mvarTL * 15
		End If
		
		If mvarTL <= 4 Then TempMax = TempMax / 10
		
		GetMaxRange = TempMax
	End Function
	
	Private Function GetAccuracy() As Integer
		
		Dim i As Short 'ammomatrix array position
		Dim R As Single
		Dim Acc As Short
		
		
		Acc = mvarTL
		'TODO find out if this max of 8 should be done first or last
		If Acc > 8 Then Acc = 8
		
		If mvarStyle = "medium" Then
			Acc = Acc + 1
		ElseIf mvarStyle = "heavy" Then 
			Acc = Acc + 2
		End If
		
		'get modifier for Cheap, Fine and Very Fine quality
		If mvarQuality = "cheap" Then
			Acc = Acc - 1
		ElseIf mvarQuality = "fine (accurate)" Then 
			Acc = Acc + 1
		ElseIf mvarQuality = "very fine (accurate)" Then 
			Acc = Acc + 2
		End If
		
		GetAccuracy = Acc
	End Function
	
	Private Function GetWeight() As Single
		Dim l As Single
		Dim W As Single
		Dim T As Single
		Dim TempWeight As Single
		
		'get L
		If mvarDatatype = WaterCannon Then
			l = 160
		ElseIf mvarDatatype = FlameThrower Then 
			l = 200
		End If
		
		'get W
		If mvarStyle = "heavy" Then
			W = 2
		ElseIf mvarStyle = "medium" Then 
			W = 1
		ElseIf mvarStyle = "light" Then 
			W = 0.5
		End If
		
		'get T
		If mvarTL <= 5 Then
			T = 1.25
		ElseIf mvarTL = 6 Then 
			T = 0.75
		ElseIf mvarTL = 7 Then 
			T = 0.5
		ElseIf mvarTL >= 8 Then 
			T = 0.25
		End If
		
		
		TempWeight = l * W * T + (T * mvarShots)
		
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
		
		If mvarTL <= 5 Then
			TSS = 10
		ElseIf mvarTL >= 6 Then 
			TSS = 5
		End If
		
		GetSnapShot = TSS
	End Function
	
	Private Function GetRoF() As String
		
		If mvarDatatype = WaterCannon Then
			If mvarTL <= 5 Then
				GetRoF = CStr(1)
			ElseIf mvarTL = 6 Then 
				GetRoF = CStr(3)
			ElseIf mvarTL >= 7 Then 
				GetRoF = CStr(4)
			End If
		ElseIf mvarDatatype = FlameThrower Then 
			If mvarTL <= 5 Then
				GetRoF = CStr(4)
			ElseIf mvarTL = 6 Then 
				GetRoF = CStr(8)
			ElseIf mvarTL >= 7 Then 
				GetRoF = CStr(8)
			End If
		End If
		
	End Function
	
	Private Function GetCost() As Single
		Dim TempCost As Single
		
		
		If mvarTL <= 6 Then
			TempCost = mvarWeight * 5
		ElseIf mvarTL >= 7 Then 
			TempCost = mvarWeight * 25
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
	
	Private Function GetLoaders() As Integer
		
		If mvarTL <= 4 Then
			If mvarStyle = "heavy" Then
				GetLoaders = 2
			ElseIf mvarStyle = "medium" Then 
				GetLoaders = 1
			Else
				GetLoaders = 0
			End If
		Else
			GetLoaders = 0
		End If
		
	End Function
	
	Private Function GetWPS() As Single
		If mvarDatatype = WaterCannon Then
			GetWPS = 4.25
		ElseIf mvarDatatype = FlameThrower Then 
			GetWPS = 3
		End If
	End Function
	
	Private Function GetCPS() As Single
		If mvarDatatype = FlameThrower Then
			GetCPS = 0.5
		ElseIf mvarDatatype = WaterCannon Then 
			If mvarAmmunitionType = "water" Then
				GetCPS = 0
			ElseIf mvarAmmunitionType = "acid" Then 
				GetCPS = 0.5
			ElseIf mvarAmmunitionType = "foam" Then 
				GetCPS = 0.5
			End If
		End If
		
	End Function
End Class