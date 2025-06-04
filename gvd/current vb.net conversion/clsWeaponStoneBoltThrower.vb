Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsWeaponStoneBoltThrower_NET.clsWeaponStoneBoltThrower")> Public Class clsWeaponStoneBoltThrower
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
	Private mvarRuggedized As Boolean
	Private mvarTL As Short
	Private mvarVolume As Double
	Private mvarWeight As Double
	Private mvarStrength As Integer
	Private mvarMagazineCapacity As Integer
	Private mvarAmmunitionType As String
	Private mvarMalfunction As String
	Private mvarTypeDamage As String
	Private mvarDamage As String
	Private mvarhalfDamage As Double
	Private mvarMaxRange As Double
	Private mvarMinRange As Single
	Private mvarAccuracy As Integer
	Private mvarSnapShot As Integer
	Private mvarRoF As String
	Private mvarWPS As Single
	Private mvarVPS As Single
	Private mvarCPS As Single
	Private mvarShots As Integer
	Private mvarLoaders As Integer
	Private mvarQuality As String
	Private mvarMechanism As String
	Private mvarCustomDescription As String
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
	
	
	
	Public Property Mechanism() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Mechanism
			Mechanism = mvarMechanism
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Mechanism = 5
			mvarMechanism = Value
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
	
	
	
	
	
	Public Property VPS() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.VPS
			VPS = mvarVPS
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.VPS = 5
			mvarVPS = Value
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
	
	
	
	
	
	Public Property MinRange() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.MinRange
			MinRange = mvarMinRange
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.MinRange = 5
			mvarMinRange = Value
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
	
	
	
	
	
	Public Property MagazineCapacity() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.MagazineCapacity
			MagazineCapacity = mvarMagazineCapacity
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.MagazineCapacity = 5
			mvarMagazineCapacity = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Strength() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Strength
			Strength = mvarStrength
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Strength = 5
			mvarStrength = Value
			If mvarZZInit = 0 Then Exit Property
			
			If mvarDatatype = BoltThrower Then
				If mvarMechanism = "spring-powered" Then
					If mvarStrength > 75 Then
						mvarStrength = 75
						modHelper.InfoPrint(1, "Max Practical Strength for Spring-Powered Bolt Thrower is 75")
					End If
				Else
					If mvarStrength > 150 Then
						mvarStrength = 150
						modHelper.InfoPrint(1, "Max Practical Strength for Torsion-Powered Bolt Thrower is 150")
					End If
				End If
			ElseIf mvarDatatype = RepeatingBoltThrower Then 
				If mvarStrength > Val(CStr(mvarTL)) * 5 Then
					mvarStrength = Val(CStr(mvarTL)) * 5
					modHelper.InfoPrint(1, "Repeating Bolt Throwers Strength cannot exceed TL * 5")
				End If
			Else
				If mvarMechanism = "spring-powered" Then
					If mvarStrength > 50 Then
						mvarStrength = 50
						modHelper.InfoPrint(1, "Maximum Strength for Spring-Powered Stone Thrower is 50")
					End If
				ElseIf mvarMechanism = "torsion-powered" Then 
					If mvarStrength > 500 Then
						mvarStrength = 500
						modHelper.InfoPrint(1, "Maximum Strength for Torsion-Powered Stone Throwers = 500")
					End If
				End If
			End If
			
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
		
		
		If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = equipmentPod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = Module_Renamed) Then
			TempCheck = True
		Else
			modHelper.InfoPrint(1, "Weapons must be placed in Body, Superstructure, Pod, equipment Pod,Turret, Popturret, Arm, Wing, Open Mount, Leg or Module.")
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
		mvarStrength = 50
		mvarMechanism = "spring-powered" 'others are Torsion-Powered and Counterweight"
		mvarQuality = "normal"
		mvarMount = "normal"
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
			Case StoneThrower
				
				mvarAmmunitionType = "stones"
			Case BoltThrower
				
				mvarAmmunitionType = "bolts"
			Case RepeatingBoltThrower
				
				mvarMagazineCapacity = 5
				mvarAmmunitionType = "bolts"
		End Select
		
	End Sub
	
	
	
	Public Sub StatsUpdate()
		mvarZZInit = 1
		Dim sPrint1 As String
		Dim sPrint2 As String
		Dim sPrintPlural As String
		Dim sPrintPlural2 As String
		Dim sPrintPlural3 As String
		Dim sPrintDirection As String
		Dim QRugMod As Single
		Dim RugHitMod As Integer
		
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
		
		mvarMalfunction = GetMalfunction
		mvarTypeDamage = GetTypeDamage
		mvarDamage = GetDamage
		mvarhalfDamage = GetHalfDamage
		mvarMaxRange = GetMaxRange
		mvarMinRange = GetMinRange
		mvarAccuracy = GetAccuracy
		mvarWeight = GetWeight
		mvarVolume = GetVolume
		mvarSnapShot = GetSnapShot
		mvarRoF = GetRoF
		mvarCost = GetCost
		mvarWPS = GetWeightPerShot
		mvarVPS = GetVolumePerShot
		mvarCPS = GetCostPerShot
		mvarShots = GetShots
		mvarLoaders = GetLoaders
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea) * RugHitMod
		
		
		'//update the cost,weight,volume, surface area and volume based on quantity and ruggedized options
		mvarCost = mvarCost * QRugMod
		mvarWeight = mvarWeight * QRugMod
		mvarVolume = mvarVolume * QRugMod
		
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		
		'produce the print output
		'produce the print output
		If mvarRuggedized Then
			sPrint1 = "ruggedized "
		Else
			sPrint1 = ""
		End If
		sPrint1 = sPrint1 & "ST " & mvarStrength & " "
		
		If mvarMount <> "normal" Then
			sPrint2 = sPrint2 & ", " & mvarMount
		End If
		If mvarQuality <> "normal" Then sPrint2 = sPrint2 & ", " & mvarQuality & " construction"
		
		sPrint2 = sPrint2 & ", " & mvarMechanism
		
		If mvarMagazineCapacity <> 0 Then
			sPrint2 = sPrint2 & ", " & VB6.Format(mvarMagazineCapacity) & " round magazine"
		End If
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
		
		mvarPrintOutput = NumericToString(mvarQuantity) & " TL" & mvarTL & " " & sPrint1 & mvarCustomDescription & sPrintPlural & sPrint2 & " (" & mvarLocation & sPrintDirection & ", HP " & mvarHitPoints & sPrintPlural2 & ", " & sPrintPlural3 & VB6.Format(mvarWeight, p_sFormat) & " lbs., " & VB6.Format(mvarVolume, p_sFormat) & " cf., " & "$" & VB6.Format(mvarCost, p_sFormat) & ")." & mvarComment
		
		'note: cost, malf, and accuracy modifiers for Cheap, Fine and Very Fine quality are calced in the functions below
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
		
		If mvarDatatype = RepeatingBoltThrower Then
			TempMalf = "16"
		Else
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
	
	Private Function GetTypeDamage() As String
		
		If mvarDatatype = StoneThrower Then
			GetTypeDamage = "Cr."
		Else
			GetTypeDamage = "Imp."
		End If
		
	End Function
	
	Private Function GetDamage() As String
		Dim dam1 As Integer
		Dim dam2 As Integer
		Dim i As Integer
		
		Dim WeaponST As String
		
		WeaponST = Str(mvarStrength)
		
		'find base damage using the lookup table
		If mvarDatatype = StoneThrower Then
			If CDbl(WeaponST) <= 100 Then
				dam1 = StoneBoltDamageMatrix(CInt(WeaponST)).Swing1
				dam2 = StoneBoltDamageMatrix(CInt(WeaponST)).Swing2
			Else
				i = Int(CDbl(WeaponST) / 10)
				dam1 = 3 + i
			End If
		Else
			If CDbl(WeaponST) <= 100 Then
				dam1 = StoneBoltDamageMatrix(CInt(WeaponST)).Thrust1
				dam2 = StoneBoltDamageMatrix(CInt(WeaponST)).Thrust2
			Else
				i = Int(CDbl(WeaponST) / 10)
				dam1 = 1 + i
			End If
		End If
		
		'apply modifiers depending on Stone or Bolt thrower
		If mvarDatatype = StoneThrower Then
			If mvarMechanism = "counterweight" Then
				dam2 = dam2 + (1 * dam1)
			End If
		Else
			If mvarStrength <= 20 Then
				dam2 = dam2 + 4
			Else
				dam1 = dam1 + 1
			End If
		End If
		
		'create merged damage string
		If dam2 = 0 Then
			GetDamage = dam1 & "d"
		ElseIf dam2 < 0 Then 
			GetDamage = dam1 & "d" & " -" & System.Math.Abs(dam2)
		Else
			GetDamage = dam1 & "d" & " +" & dam2
		End If
		
	End Function
	
	Private Function GetHalfDamage() As Single 'in yards
		Dim TempHalfDamage As Single
		
		If mvarDatatype = StoneThrower Then
			If mvarStrength <= 20 Then
				TempHalfDamage = mvarStrength * 10
			Else
				TempHalfDamage = 190 + (mvarStrength / 2)
			End If
			'counterweight stone throwers must be modified to have shorter range
			If mvarMechanism = "counterweight" Then TempHalfDamage = TempHalfDamage * 0.75
		Else
			If mvarStrength <= 20 Then
				TempHalfDamage = mvarStrength * 20
			Else
				TempHalfDamage = 380 + mvarStrength
			End If
		End If
		
		GetHalfDamage = TempHalfDamage
	End Function
	
	Private Function GetMaxRange() As Double 'in yards
		GetMaxRange = 1.25 * mvarhalfDamage
	End Function
	
	Private Function GetMinRange() As Single 'in yards
		If mvarMechanism = "counterweight" Then
			GetMinRange = mvarMaxRange / 4
		Else
			GetMinRange = 0
		End If
	End Function
	
	Private Function GetAccuracy() As Integer
		Dim TempAccuracy As Integer
		
		If mvarDatatype = RepeatingBoltThrower Then
			TempAccuracy = 5
		ElseIf mvarDatatype = BoltThrower Then 
			TempAccuracy = 6
		Else
			If mvarMechanism = "counterweight" Then
				TempAccuracy = 1
			Else
				TempAccuracy = 2
			End If
		End If
		
		'get modifier for Cheap, Fine and Very Fine quality
		If mvarQuality = "cheap" Then
			TempAccuracy = TempAccuracy - 1
		ElseIf mvarQuality = "fine (accurate)" Then 
			TempAccuracy = TempAccuracy + 1
		ElseIf mvarQuality = "very fine (accurate)" Then 
			TempAccuracy = TempAccuracy + 2
		End If
		
		GetAccuracy = TempAccuracy
	End Function
	
	Private Function GetWeight() As Double
		Dim ST As Integer ' strength
		Dim P As Single 'power mechanism modifier
		Dim T As Single 'tech level modifier
		Dim M As Single 'magazine capacity modifier
		Dim D As Single 'datatype modifier
		
		ST = mvarStrength
		M = 1 + (0.05 * mvarMagazineCapacity)
		
		If mvarMechanism = "spring-powered" Then
			P = 1
		ElseIf mvarMechanism = "torsion-powered" Then 
			P = 0.8
		Else
			P = 0.5
		End If
		
		If mvarTL <= 4 Then
			T = 1
		ElseIf mvarTL = 5 Then 
			T = 0.75
		ElseIf mvarTL = 6 Then 
			T = 0.6
		Else
			T = 0.5
		End If
		
		If mvarDatatype = StoneThrower Then
			D = 0.25
		Else
			D = 0.1
		End If
		
		GetWeight = System.Math.Round(ST * ST * P * T * M * D, 5)
	End Function
	
	Private Function GetVolume() As Double
		If mvarMount = "normal" Then
			GetVolume = mvarWeight / 50
		Else
			GetVolume = mvarWeight / 20 'concealed weapons take up more space
		End If
		
	End Function
	
	Private Function GetSnapShot() As Integer
		Dim TSS As Integer
		
		Const Min As Short = 25 'minimum value for counterweight stone throwers
		
		If mvarWeight < 15 Then
			TSS = 12
		ElseIf mvarWeight <= 25 Then 
			TSS = 15
		ElseIf mvarWeight <= 400 Then 
			TSS = 20
		Else
			TSS = 25
		End If
		
		
		If mvarDatatype = StoneThrower Then
			TSS = TSS + 5 'add five to all stonethrowers
			If mvarMechanism = "counterweight" Then
				If TSS < Min Then TSS = Min
			End If
		End If
		
		GetSnapShot = TSS
	End Function
	
	Private Function GetRoF() As String
		Dim TempRoF As Double
		
		If mvarDatatype = RepeatingBoltThrower Then
			TempRoF = System.Math.Round(2 * System.Math.Sqrt(mvarStrength), 0)
		ElseIf mvarMechanism = "counterweight" Then 
			TempRoF = System.Math.Round(10 * System.Math.Sqrt(mvarStrength), 0)
		Else
			TempRoF = System.Math.Round(5 * System.Math.Sqrt(mvarStrength), 0)
		End If
		
		GetRoF = "1/" & TempRoF
		
	End Function
	
	Private Function GetCost() As Double
		Dim M As Short 'weapon type modifier
		Dim P As Short 'power mechanism modifier
		Dim TempCost As Double
		
		If mvarMechanism = "spring-powered" Then
			M = 1
		Else
			M = 2
		End If
		
		If mvarDatatype = RepeatingBoltThrower Then
			P = 2
		Else
			P = 1
		End If
		
		If mvarWeight < 100 Then
			TempCost = 25 * mvarWeight * M * P
		Else
			TempCost = (2400 + mvarWeight) * M * P
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
		
		GetCost = System.Math.Round(TempCost, 2)
	End Function
	
	Private Function GetWeightPerShot() As Single
		Dim sngTemp As Single
		
		If mvarDatatype = StoneThrower Then
			sngTemp = mvarStrength / 10
		Else
			sngTemp = mvarStrength * mvarStrength / 1600
		End If
		
		GetWeightPerShot = System.Math.Round(sngTemp, 5)
	End Function
	
	Private Function GetVolumePerShot() As Single
		
		GetVolumePerShot = mvarWPS / 50
	End Function
	
	Private Function GetCostPerShot() As Single
		Const Min As Short = 2
		Dim TempCost As Single
		
		If mvarDatatype = StoneThrower Then
			GetCostPerShot = mvarWPS * 0.5
		Else
			TempCost = mvarWPS * 2
			If TempCost < Min Then TempCost = Min
			GetCostPerShot = TempCost
		End If
	End Function
	
	Private Function GetShots() As Integer
		'number of shots the weapon has ready to fire.
		'stone throwers and bolt throwers have 1.  Repeating bolt throwers use their
		'magaizine capacity
		If mvarDatatype = RepeatingBoltThrower Then
			GetShots = mvarMagazineCapacity
		Else
			GetShots = 1
		End If
		
	End Function
	
	Private Function GetLoaders() As Integer
		'TODO inquire to Pulver.  There is a discrepancy in Crew versus Loaders
		' of mech weapons and guns.  One asks for totalcrew while other asks for loaders. Why?
		Dim Divisor As Short
		Dim TempCrew As Single
		Const GunnerST As Short = 12 ' the arbitary value for the gunner's strength
		
		If mvarMechanism = "counterweight" Then
			Divisor = 20
		Else
			Divisor = 40
		End If
		
		If GunnerST + 4 <= mvarStrength Then
			TempCrew = (mvarStrength / Divisor) - 1
			GetLoaders = RoundUP(TempCrew)
		Else 'no loaders required
			GetLoaders = 0
		End If
		
	End Function
End Class