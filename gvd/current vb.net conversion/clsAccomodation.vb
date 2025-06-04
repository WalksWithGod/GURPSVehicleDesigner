Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsAccommodation_NET.clsAccommodation")> Public Class clsAccommodation
	
	'local variable(s) to hold property value(s)
	Private mvarWeight As Double
	Private mvarVolume As Double
	Private mvarCost As Double
	Private mvarLocation As String
	'local variable(s) to hold property value(s)
	Private mvarParent As String
	Private mvarKey As String
	Private mvarDR As Integer
	Private mvarRuggedized As Boolean
	Private mvarSurfaceArea As Double
	Private mvarHitPoints As Double
	Private mvarTL As Short
	
	Private mvarDatatype As Short
	Private mvarParentDatatype As Short
	Private mvarDescription As String
	Private mvarCustomDescription As String
	Private mvarCustom As Boolean
	Private mvarQuantity As Short
	Private mvarExposed As Boolean
	Private mvarAddedVolume As Single
	Private mvarGSeat As Boolean
	Private mvarOccupancy As Integer
	
	Private mvarPrintOutput As String
	'local variable(s) to hold property value(s)
	Private mvarImage As Short
	Private mvarSelectedImage As Short
	Private mvarComment As String
	Private mvarCName As String
	'local variable(s) to hold property value(s)
	Private mvarMatrixPos As Integer
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
			'Syntax: Debug.Print X.Occupancy
			Occupancy = mvarOccupancy
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Occupancy = 5
			mvarOccupancy = Value
			If mvarZZInit = 0 Then Exit Property
			
			If mvarOccupancy < 1 Then
				modHelper.InfoPrint(1, "Occupancy can be no less than 1")
				mvarOccupancy = 1
			End If
			
			If (mvarDatatype = Cabin) Or (mvarDatatype = LuxuryCabin) Then
				If mvarOccupancy > 2 Then
					modHelper.InfoPrint(1, "Cabins can have an occupancy of 1 or 2.  For larger accomodations use a Suite.")
					mvarOccupancy = 2
				End If
			End If
			
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
	
	
	
	Public Property Exposed() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Exposed
			Exposed = mvarExposed
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Exposed = 5
			mvarExposed = Value
		End Set
	End Property
	
	
	
	Public Property GSeat() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.GSeat
			GSeat = mvarGSeat
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.GSeat = 5
			mvarGSeat = Value
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
	
	
	
	Public Property AddedVolume() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.AddedVolume
			AddedVolume = mvarAddedVolume
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.AddedVolume = 5
			mvarAddedVolume = Value
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
	
	
	
	
	Public Function LocationCheck() As Boolean
		Dim InstallPoint As Short
		Dim TempCheck As Boolean
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallPoint = Veh.Components(mvarParent).Datatype
		
		Select Case mvarDatatype
			
			Case CycleSeat, CrampedSeat, NormalSeat, RoomySeat, CrampedStandingRoom, NormalStandingRoom, RoomyStandingRoom
				If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Pod) Or (InstallPoint = Arm) Or (InstallPoint = Leg) Or (InstallPoint = Wing) Or (InstallPoint = equipmentPod) Then
					TempCheck = True
				Else
					modHelper.InfoPrint(1, "Cramped, Normal, and Roomy Seats and Standing Room must be placed in Body, Superstructure, Turret, Popturret, Pod, equipment Pod, Wing, Arm, or Leg.")
					TempCheck = False
				End If
			Case Hammock, Bunk, Cabin, LuxuryCabin, Suite, LuxurySuite, SmallGalley
				If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Pod) Or (InstallPoint = Arm) Or (InstallPoint = Leg) Or (InstallPoint = Wing) Or (InstallPoint = equipmentPod) Then
					TempCheck = True
				Else
					modHelper.InfoPrint(1, "Hammocks, Bunks, Cabins, Suites and Galleys must be placed in Body, Superstructure, Turret, Popturret, Pod, equipment Pod, Wing, Arm, or Leg.")
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
		mvarTL = gVehicleTL
		mvarQuantity = 1
		mvarExposed = False
		mvarGSeat = False
		
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
			Case CrampedSeat
				
			Case NormalSeat
				
			Case RoomySeat
				
			Case NormalStandingRoom
				
			Case CrampedStandingRoom
				
			Case RoomyStandingRoom
				
			Case CycleSeat
				
			Case Hammock
				
			Case Bunk
				
			Case Cabin
				mvarOccupancy = 1
			Case LuxuryCabin
				mvarOccupancy = 1
				
			Case SmallGalley
				
			Case Suite
				mvarOccupancy = 1
			Case LuxurySuite
				mvarOccupancy = 1
				
		End Select
		
	End Sub
	
	Public Sub GetMatrixIndex()
		Dim i As Short
		
		If mvarDatatype = 0 Then Exit Sub
		
		mvarMatrixPos = 0 'init the counter
		For i = 1 To UBound(AccommodationsMatrix)
			If AccommodationsMatrix(i).ID = mvarDatatype Then
				If AccommodationsMatrix(i).TL >= mvarTL Then
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
		Dim VolumeMod As Single
		Dim CostMod As Single
		Dim sPrint1 As String
		Dim sPrint2 As String
		Dim sPrintPlural As String
		Dim sPrintPlural2 As String
		Dim sPrintPlural3 As String
		Dim sPrintPlural4 As String
		
		mvarLocation = GetLocation
		
		Select Case mvarDatatype
			'note: i dont think these can be ruggedized so i dont include
			' the modifier calcs for that
			Case Hammock, Bunk, Cabin, LuxuryCabin, SmallGalley
				mvarWeight = AccommodationsMatrix(mvarMatrixPos).Weight
				mvarCost = AccommodationsMatrix(mvarMatrixPos).Cost
				mvarVolume = AccommodationsMatrix(mvarMatrixPos).Volume + AddedVolume
				If mvarGSeat Then mvarCost = mvarCost + (500 * mvarOccupancy)
				
			Case Suite, LuxurySuite
				mvarWeight = mvarOccupancy / 2 * AccommodationsMatrix(mvarMatrixPos).Weight
				mvarCost = mvarOccupancy / 2 * AccommodationsMatrix(mvarMatrixPos).Cost
				mvarVolume = (mvarOccupancy / 2) * (AccommodationsMatrix(mvarMatrixPos).Volume + AddedVolume)
				If mvarGSeat Then mvarCost = mvarCost + (500 * mvarOccupancy)
				
			Case Else
				If mvarExposed Then 'todo: how to handle stats modifiers like these?  They cant be handled in oStats
					'since its not a general modifier... it only applies to this seat component types.
					' same with GSeat modifier...  how to apply these to the stats?
					
					' I could have volmodifer variable that gets set before the calc is done in oStats.
					' E.G oStats.SetVolumeModifier = 0.5
					
					'perhaps these too MUST be set in the IDL file for this item.  Then, instead of
					' exposed being an option, its simply a TYPE of seat.  Same with gseat, it must be a
					' type of seat.  Actually, its is "type" then, you still need to be able to check
					' for exposed seats in performance profiles for instance... hrm.  Rule should be'
					' either something is a TYPE or its a BASE property and all objects based on this
					' class ID share it.
					
					VolumeMod = 0.5
				Else
					VolumeMod = 1
				End If
				If mvarGSeat Then
					CostMod = 500 * mvarOccupancy
				Else
					CostMod = 0
				End If
				mvarWeight = AccommodationsMatrix(mvarMatrixPos).Weight
				mvarCost = AccommodationsMatrix(mvarMatrixPos).Cost + CostMod
				mvarVolume = VolumeMod * AccommodationsMatrix(mvarMatrixPos).Volume
		End Select
		
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)
		
		mvarWeight = System.Math.Round(mvarWeight * mvarQuantity, 2)
		mvarCost = System.Math.Round(mvarCost * mvarQuantity, 2)
		mvarVolume = System.Math.Round(mvarVolume * mvarQuantity, 2)
		mvarSurfaceArea = System.Math.Round(mvarSurfaceArea * mvarQuantity, 2)
		
		'produce the print output
		If mvarRuggedized Then
			sPrint1 = "ruggedized "
		Else
			sPrint1 = ""
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
		
		
		If mvarGSeat Then
			Select Case mvarDatatype
				Case Cabin, LuxuryCabin, Suite, LuxurySuite
					If mvarOccupancy > 1 Then
						sPrint2 = sPrint2 & ", with g-seats"
					Else
						sPrint2 = sPrint2 & ", with g-seat"
					End If
				Case Else
					sPrint2 = sPrint2 & ", g-seat" & sPrintPlural
			End Select
		End If
		If mvarExposed Then
			sPrint2 = ", exposed"
		End If
		If mvarOccupancy <> 0 Then
			sPrint2 = sPrint2 & ", " & VB6.Format(mvarOccupancy) & " person occupancy" & sPrintPlural3
		End If
		
		mvarPrintOutput = NumericToString(mvarQuantity) & " TL" & mvarTL & " " & sPrint1 & mvarCustomDescription & sPrintPlural & sPrint2 & " (" & mvarLocation & ", HP " & mvarHitPoints & sPrintPlural3 & ", " & sPrintPlural4 & VB6.Format(mvarWeight, p_sFormat) & " lbs., " & VB6.Format(mvarVolume, p_sFormat) & " cf., " & "$" & VB6.Format(mvarCost, p_sFormat) & ")." & mvarComment
		
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