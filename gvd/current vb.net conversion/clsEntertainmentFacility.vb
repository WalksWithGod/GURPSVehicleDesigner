Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsEntertainmentFacility_NET.clsEntertainmentFacility")> Public Class clsEntertainmentFacility
	
	Private mvarTL As Short
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
	
	Private mvarFloorArea As Single
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
			PrintOutput = mvarPrintOutput
		End Get
		Set(ByVal Value As String)
			mvarPrintOutput = Value
		End Set
	End Property
	
	
	Public Property MatrixPos() As Integer
		Get
			MatrixPos = mvarMatrixPos
		End Get
		Set(ByVal Value As Integer)
			mvarMatrixPos = Value
		End Set
	End Property
	
	
	Public Property CName() As String
		Get
			CName = mvarCName
		End Get
		Set(ByVal Value As String)
			mvarCName = Value
		End Set
	End Property
	
	
	Public Property Comment() As String
		Get
			Comment = mvarComment
		End Get
		Set(ByVal Value As String)
			mvarComment = Value
		End Set
	End Property
	
	
	Public Property SelectedImage() As Short
		Get
			SelectedImage = mvarSelectedImage
		End Get
		Set(ByVal Value As Short)
			mvarSelectedImage = Value
		End Set
	End Property
	
	
	Public Property Image() As Short
		Get
			Image = mvarImage
		End Get
		Set(ByVal Value As Short)
			mvarImage = Value
		End Set
	End Property
	
	
	Public Property Quantity() As Short
		Get
			Quantity = mvarQuantity
		End Get
		Set(ByVal Value As Short)
			mvarQuantity = Value
		End Set
	End Property
	
	
	Public Property Custom() As Boolean
		Get
			Custom = mvarCustom
		End Get
		Set(ByVal Value As Boolean)
			mvarCustom = Value
		End Set
	End Property
	
	
	Public Property CustomDescription() As String
		Get
			CustomDescription = mvarCustomDescription
		End Get
		Set(ByVal Value As String)
			mvarCustomDescription = Value
		End Set
	End Property
	
	
	Public Property Description() As String
		Get
			Description = mvarDescription
		End Get
		Set(ByVal Value As String)
			mvarDescription = Value
		End Set
	End Property
	
	
	Public Property ParentDatatype() As Short
		Get
			ParentDatatype = mvarParentDatatype
		End Get
		Set(ByVal Value As Short)
			mvarParentDatatype = Value
		End Set
	End Property
	
	
	Public Property Datatype() As Short
		Get
			Datatype = mvarDatatype
		End Get
		Set(ByVal Value As Short)
			mvarDatatype = Value
		End Set
	End Property
	
	
	Public Property HitPoints() As Double
		Get
			HitPoints = mvarHitPoints
		End Get
		Set(ByVal Value As Double)
			mvarHitPoints = Value
		End Set
	End Property
	
	
	Public Property FloorArea() As Single
		Get
			FloorArea = mvarFloorArea
		End Get
		Set(ByVal Value As Single)
			mvarFloorArea = Value
		End Set
	End Property
	
	
	Public Property SurfaceArea() As Double
		Get
			SurfaceArea = mvarSurfaceArea
		End Get
		Set(ByVal Value As Double)
			mvarSurfaceArea = Value
		End Set
	End Property
	
	
	Public Property Ruggedized() As Boolean
		Get
			Ruggedized = mvarRuggedized
		End Get
		Set(ByVal Value As Boolean)
			mvarRuggedized = Value
		End Set
	End Property
	
	
	Public Property DR() As Integer
		Get
			DR = mvarDR
		End Get
		Set(ByVal Value As Integer)
			mvarDR = Value
		End Set
	End Property
	
	
	Public Property Key() As String
		Get
			Key = mvarKey
		End Get
		Set(ByVal Value As String)
			mvarKey = Value
		End Set
	End Property
	
	
	Public Property Parent() As String
		Get
			Parent = mvarParent
		End Get
		Set(ByVal Value As String)
			mvarParent = Value
		End Set
	End Property
	
	
	Public Property Location() As String
		Get
			Location = mvarLocation
		End Get
		Set(ByVal Value As String)
			mvarLocation = Value
		End Set
	End Property
	
	
	Public Property PowerReqt() As Double
		Get
			PowerReqt = mvarPowerReqt
		End Get
		Set(ByVal Value As Double)
			mvarPowerReqt = Value
		End Set
	End Property
	
	
	Public Property Cost() As Double
		Get
			Cost = mvarCost
		End Get
		Set(ByVal Value As Double)
			mvarCost = Value
		End Set
	End Property
	
	
	Public Property Volume() As Double
		Get
			Volume = mvarVolume
		End Get
		Set(ByVal Value As Double)
			mvarVolume = Value
		End Set
	End Property
	
	
	Public Property Weight() As Double
		Get
			Weight = mvarWeight
		End Get
		Set(ByVal Value As Double)
			mvarWeight = Value
		End Set
	End Property
	
	
	Public Property TL() As Short
		Get
			TL = mvarTL
		End Get
		Set(ByVal Value As Short)
			mvarTL = Value
			GetMatrixIndex()
		End Set
	End Property
	
	Public Function LocationCheck() As Boolean
		Dim TempCheck As Boolean
		Dim InstallPoint As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallPoint = Veh.Components(mvarParent).Datatype
		
		If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Pod) Then
			TempCheck = True
		Else
			modHelper.InfoPrint(1, "Entertainment Facilities must be placed in Body, Superstructure, Turret, Popturret or Pod.")
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
			Case Stage
				
				mvarFloorArea = 100
			Case Hall
				
				mvarFloorArea = 600
			Case BarRoom
				
				mvarFloorArea = 500
			Case ConferenceRoom
				
				mvarFloorArea = 400
			Case MovieScreenandProjector
				
			Case MovieScreenandProjectorSmall
				
			Case HoloventureZone
				mvarFloorArea = 100
				
		End Select
		
	End Sub
	
	Public Sub GetMatrixIndex()
		Dim i As Short
		If mvarDatatype = 0 Then Exit Sub
		mvarMatrixPos = 0 'init the counter
		For i = 1 To UBound(EntertainmentMatrix)
			If EntertainmentMatrix(i).ID = mvarDatatype Then
				If EntertainmentMatrix(i).TL >= mvarTL Then
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
		
		Dim AreaMod As Short
		Dim QRugMod As Single 'combined quantity and ruggedized multipliers
		Dim RugHitMod As Short 'ruggedized hit point multiplier
		Dim sPrint1 As String
		Dim sPrintPlural As String
		Dim sPrintPlural2 As String
		Dim sPrintPlural3 As String
		
		mvarLocation = GetLocation
		
		'set the ruggedized and quantity multipliers
		If mvarRuggedized Then
			QRugMod = 1.5 * mvarQuantity
			RugHitMod = 2
		Else
			QRugMod = 1 * mvarQuantity
			RugHitMod = 1
		End If
		
		Select Case mvarDatatype
			Case Stage, BarRoom, Hall, ConferenceRoom, HoloventureZone
				AreaMod = mvarFloorArea / 100
				
				mvarWeight = AreaMod * EntertainmentMatrix(mvarMatrixPos).Weight
				mvarCost = AreaMod * EntertainmentMatrix(mvarMatrixPos).Cost
				mvarVolume = AreaMod * EntertainmentMatrix(mvarMatrixPos).Volume
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarFloorArea)
				
				If EntertainmentMatrix(mvarMatrixPos).Power = 0 Then
					mvarPowerReqt = 0
				Else
					mvarPowerReqt = AreaMod * EntertainmentMatrix(mvarMatrixPos).Power
				End If
				'get finals
				mvarWeight = System.Math.Round(mvarQuantity * mvarWeight, 2)
				mvarCost = System.Math.Round(mvarQuantity * mvarCost, 2)
				mvarVolume = System.Math.Round(mvarQuantity * mvarVolume, 2)
				mvarSurfaceArea = System.Math.Round(mvarQuantity * mvarFloorArea, 2)
				mvarPowerReqt = System.Math.Round(mvarQuantity * mvarPowerReqt, 2)
			Case Else
				
				mvarWeight = EntertainmentMatrix(mvarMatrixPos).Weight
				mvarCost = EntertainmentMatrix(mvarMatrixPos).Cost
				mvarVolume = EntertainmentMatrix(mvarMatrixPos).Volume
				mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
				'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
				
				mvarPowerReqt = EntertainmentMatrix(mvarMatrixPos).Power
				
				'get finals
				mvarWeight = System.Math.Round(QRugMod * mvarWeight, 2)
				mvarCost = System.Math.Round(QRugMod * mvarCost, 2)
				mvarVolume = System.Math.Round(QRugMod * mvarVolume, 2)
				mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
				mvarPowerReqt = System.Math.Round(mvarQuantity * mvarPowerReqt, 2)
		End Select
		
		'print output
		If mvarRuggedized Then
			sPrint1 = "ruggedized "
		Else
			sPrint1 = ""
		End If
		
		If mvarFloorArea <> 0 Then
			sPrint1 = sPrint1 & VB6.Format(mvarFloorArea) & " sq ft "
		End If
		
		If mvarQuantity > 1 Then
			sPrintPlural = "s"
			sPrintPlural2 = " each"
			sPrintPlural3 = " total of "
		Else
			sPrintPlural = ""
			sPrintPlural2 = ""
			sPrintPlural3 = ""
		End If
		mvarPrintOutput = NumericToString(mvarQuantity) & " TL" & mvarTL & " " & sPrint1 & mvarCustomDescription & sPrintPlural & " (" & mvarLocation & ", HP " & mvarHitPoints & sPrintPlural2 & ", " & sPrintPlural3 & VB6.Format(mvarWeight, p_sFormat) & " lbs., " & VB6.Format(mvarVolume, p_sFormat) & " cf., " & "$" & VB6.Format(mvarCost, p_sFormat) & ", " & VB6.Format(mvarPowerReqt, p_sFormat) & " kW)." & mvarComment
		
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