Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsSonar_NET.clsSonar")> Public Class clsSonar
	
	Private mvarDirection As String
	Private mvarDepthFinding As Boolean
	Private mvarDippingSonar As Boolean
	Private mvarTowedArray As Boolean
	Private mvarNoTargeting As Boolean
	Private mvarRange As Single
	Private mvarScanRating As Integer
	Private mvarTL As Short
	Private mvarWeight As Double
	Private mvarVolume As Double
	Private mvarCost As Double
	Private mvarPowerReqt As Double
	Private mvarActivePassive As Boolean
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
	Private mvarQuantity As Short
	
	
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
	
	
	
	
	Public Property ActivePassive() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.ActivePassive
			ActivePassive = mvarActivePassive
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.ActivePassive = 5
			mvarActivePassive = Value
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
	
	
	
	
	
	Public Property ScanRating() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.ScanRating
			ScanRating = mvarScanRating
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.ScanRating = 5
			mvarScanRating = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Range() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Range
			Range = mvarRange
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Range = 5
			mvarRange = Value
			If mvarZZInit = 0 Then Exit Property
			
			If mvarRange < 0.1 Then
				mvarRange = 0.1
				modHelper.InfoPrint(1, "Range cannot be less than .1")
			End If
			
			If mvarDatatype = ActiveSonar Then
				If mvarRange > mvarTL * 3 Then
					mvarRange = mvarTL * 3
					modHelper.InfoPrint(1, "Maximum range for Active Sonars is Tech Level x 3")
				End If
			Else
				If mvarRange > mvarTL * 8 Then
					mvarRange = mvarTL * 8
					modHelper.InfoPrint(1, "Maximum range for Passive Sonars is TL x 8")
				End If
			End If
			
		End Set
	End Property
	
	
	
	
	
	Public Property NoTargeting() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.NoTargeting
			NoTargeting = mvarNoTargeting
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.NoTargeting = 5
			mvarNoTargeting = Value
		End Set
	End Property
	
	
	
	Public Property TowedArray() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.TowedArray
			TowedArray = mvarTowedArray
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.TowedArray = 5
			mvarTowedArray = Value
		End Set
	End Property
	
	
	
	
	
	Public Property DippingSonar() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.DippingSonar
			DippingSonar = mvarDippingSonar
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.DippingSonar = 5
			mvarDippingSonar = Value
		End Set
	End Property
	
	
	
	
	
	Public Property DepthFinding() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.DepthFinding
			DepthFinding = mvarDepthFinding
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.DepthFinding = 5
			mvarDepthFinding = Value
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
	
	
	
	
	Public Function LocationCheck() As Boolean
		Dim TempCheck As Boolean
		Dim InstallPoint As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallPoint = Veh.Components(mvarParent).Datatype
		
		If (InstallPoint = ExtendableSensorPeriscope) Then
			'If NoPriorPeriscopeChildren(Parent) Then
			TempCheck = True
			'Else
			'    InfoPrint 1, "There is already a sensor attached to this periscope. Only one sensor device allowed!"
			'    TempCheck = False
			'End If
		ElseIf (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = equipmentPod) Or (InstallPoint = Module_Renamed) Then 
			TempCheck = True
		Else
			modHelper.InfoPrint(1, "Sonar Devices must be placed in Periscope, Body, Superstructure, Pod, equipment Pod, Turret, Popturret, Arm, Wing, Open Mount, Leg or Module.")
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
		mvarRuggedized = False
		mvarQuantity = 1
		mvarNoTargeting = False
		mvarDepthFinding = False
		mvarDippingSonar = False
		mvarActivePassive = False
		mvarTowedArray = False
		
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
			Case ActiveSonar
				
				Range = 18
				Direction = "front"
			Case PassiveSonar
				
				Range = 10
				
		End Select
		
	End Sub
	
	Public Sub GetMatrixIndex()
		Dim i As Short
		If mvarDatatype = 0 Then Exit Sub
		mvarMatrixPos = 0 'init the counter
		For i = 1 To UBound(SonarMatrix)
			If SonarMatrix(i).ID = mvarDatatype Then
				If SonarMatrix(i).TL >= mvarTL Then
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
		
		
		Dim WeightMod As Single
		Dim CostMod As Single
		Dim VolumeMod As Single
		Dim PowerMod As Single
		Dim QRugMod As Single 'combined quantity and ruggedized multipliers
		Dim RugHitMod As Short 'ruggedized hit point multiplier
		Dim sPrint1 As String
		Dim sPrint2 As String
		Dim sPrintPlural As String
		Dim sPrintPlural2 As String
		Dim sPrintPlural3 As String
		Dim sPrintDirection As String
		
		mvarLocation = GetLocation
		
		'set the ruggedized and quantity multipliers
		If mvarRuggedized Then
			QRugMod = 1.5 * mvarQuantity
			RugHitMod = 2
		Else
			QRugMod = 1 * mvarQuantity
			RugHitMod = 1
		End If
		
		
		WeightMod = 1
		VolumeMod = 1
		CostMod = 1
		PowerMod = 1
		If mvarActivePassive Then
			WeightMod = 1.5
			VolumeMod = 1.5
			CostMod = 1.5
			PowerMod = 1
		End If
		
		If mvarDepthFinding Then
			CostMod = CostMod * 0.25
		End If
		
		If mvarDippingSonar Then
			WeightMod = WeightMod * 1.5
			VolumeMod = VolumeMod * 1.5
			CostMod = CostMod * 1.5
		End If
		
		If mvarTowedArray Then
			WeightMod = WeightMod * 1.5
			CostMod = CostMod * 4
			VolumeMod = VolumeMod * 1.5
			PowerMod = PowerMod * 1
		End If
		
		If mvarNoTargeting Then
			CostMod = CostMod * 0.5
		End If
		
		'get base
		mvarWeight = WeightMod * SonarMatrix(mvarMatrixPos).Weight * mvarRange
		mvarVolume = mvarRange * VolumeMod * SonarMatrix(mvarMatrixPos).Volume
		mvarCost = mvarRange * CostMod * SonarMatrix(mvarMatrixPos).Cost
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
		mvarScanRating = GetScanRating(mvarRange)
		mvarPowerReqt = SonarMatrix(mvarMatrixPos).Power * PowerMod * mvarRange
		
		'get finals
		mvarWeight = System.Math.Round(QRugMod * mvarWeight, 2)
		mvarCost = System.Math.Round(QRugMod * mvarCost, 2)
		mvarVolume = System.Math.Round(QRugMod * mvarVolume, 2)
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		mvarPowerReqt = System.Math.Round(mvarQuantity * mvarPowerReqt, 2)
		
		'produce the print output
		If mvarRuggedized Then
			sPrint1 = "ruggedized "
		Else
			sPrint1 = ""
		End If
		
		If mvarActivePassive Then
			sPrint2 = ", active/passive option"
		End If
		If mvarDepthFinding Then
			If sPrint2 = "" Then
				sPrint2 = ", depth finding option"
			Else
				sPrint2 = sPrint2 & ", " & "depth finding option"
			End If
		End If
		If mvarDippingSonar Then
			If sPrint2 = "" Then
				sPrint2 = ", dipping sonar"
			Else
				sPrint2 = sPrint2 & ", " & "dipping sonar"
			End If
		End If
		If mvarTowedArray Then
			If sPrint2 = "" Then
				sPrint2 = ", towed array"
			Else
				sPrint2 = sPrint2 & ", towed array"
			End If
		End If
		If mvarNoTargeting Then
			If sPrint2 = "" Then
				sPrint2 = ", no targeting option"
			Else
				sPrint2 = sPrint2 & ", " & "no targeting option"
			End If
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
		
		mvarPrintOutput = NumericToString(mvarQuantity) & " TL" & mvarTL & " " & sPrint1 & mvarCustomDescription & sPrintPlural & sPrint2 & ", " & VB6.Format(mvarRange, p_sFormat) & " mile range" & " (" & mvarLocation & sPrintDirection & ", HP " & mvarHitPoints & sPrintPlural2 & ", " & sPrintPlural3 & VB6.Format(mvarWeight, p_sFormat) & " lbs., " & VB6.Format(mvarVolume, p_sFormat) & " cf., " & "$" & VB6.Format(mvarCost, p_sFormat) & ", " & VB6.Format(mvarPowerReqt, p_sFormat) & " kW" & ", Scan " & VB6.Format(mvarScanRating) & ")." & mvarComment
		
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