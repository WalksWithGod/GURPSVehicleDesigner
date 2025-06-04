Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsSail_NET.clsSail")> Public Class clsSail
	
	Private mvarWind As String
	Private mvarMaterial As String
	Private mvarWeight As Double
	Private mvarCost As Double
	Private mvarMotiveThrust As Single
	Private mvarLocation As String
	Private mvarParent As String
	Private mvarKey As String
	Private mvarVolume As Double
	Private mvarDR As Integer
	Private mvarSurfaceArea As Double
	Private mvarHitPoints As Double
	Private mvarTL As Short
	
	
	Private mvarDatatype As Short
	Private mvarParentDatatype As Short
	Private mvarDescription As String
	Private mvarCustomDescription As String
	Private mvarCustom As Boolean
	
	Private mvarWindForce As Single 'holds the value of the wind force
	' holds the index of the SailMatrix
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
	
	
	
	
	Public Property MotiveThrust() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.MotiveThrust
			MotiveThrust = mvarMotiveThrust
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.MotiveThrust = 5
			mvarMotiveThrust = Value
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
	
	
	
	
	
	Public Property Material() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Material
			Material = mvarMaterial
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Material = 5
			mvarMaterial = Value
		End Set
	End Property
	
	
	
	
	Public Property Wind() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Wind
			Wind = mvarWind
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Wind = 5
			mvarWind = Value
			If mvarZZInit = 0 Then Exit Property
			
			Select Case mvarWind
				Case "calm"
					mvarWindForce = 0
				Case "light air"
					mvarWindForce = 1
				Case "light breeze"
					mvarWindForce = 2
				Case "gentle breeze"
					mvarWindForce = 3
				Case "moderate breeze"
					mvarWindForce = 4
				Case "fresh breeze"
					mvarWindForce = 5
				Case "strong breeze"
					mvarWindForce = 6
				Case "gale force winds"
					mvarWindForce = 7
			End Select
			
		End Set
	End Property
	
	
	
	Public Property WindForce() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.WindForce
			WindForce = CStr(mvarWindForce)
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.WindForce = 5
			mvarWindForce = CSng(Value)
			
		End Set
	End Property
	
	
	
	
	Function LocationCheck() As Boolean
		Dim TempCheck As Boolean
		Select Case mvarDatatype
			Case SquareRig, FullRig, ForeandAftRig, AerialSail, AerialSailForeAftRig
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(Parent).Datatype <> Mast Then
					TempCheck = False
					modHelper.InfoPrint(1, "Sail must be attached to Masts")
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
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		' set the default properties
		
		'mvarCustom = False
		
		TL = gVehicleTL
		mvarDR = 3
		mvarMaterial = "cloth"
		mvarWind = "fresh breeze"
		WindForce = CStr(5) 'this is determined by mvarWind!! They go hand in hand
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
		'set default descirption values for Datatype
		If mvarDatatype = SquareRig Then
			
		ElseIf mvarDatatype = ForeandAftRig Then 
			
		ElseIf mvarDatatype = FullRig Then 
			
		ElseIf mvarDatatype = AerialSail Then 
			
		ElseIf mvarDatatype = AerialSailForeAftRig Then 
			
		End If
		
		
	End Sub
	
	Public Sub GetMatrixIndex()
		'This must be called each time the tech level changes!!!
		Dim i As Short
		If mvarDatatype = 0 Then Exit Sub
		mvarMatrixPos = 0 'init the counter
		For i = 1 To UBound(SailMatrix)
			If SailMatrix(i).ID = mvarDatatype Then
				If SailMatrix(i).TL >= mvarTL Then
					mvarMatrixPos = i
					Exit For
				Else
					mvarMatrixPos = i
				End If
			End If
		Next 
		StatsUpdate()
	End Sub
	
	
	Public Sub StatsUpdate()
		mvarZZInit = 1
		If mvarMatrixPos = 0 Then Exit Sub
		
		Dim element As Object
		Dim TotalMastHeight As Single
		Dim NumberofMasts As Short
		Dim arrSubs() As String
		Dim j As Integer
		
		mvarLocation = GetLocation
		
		On Error Resume Next
		'get the mast height and number of masts
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		arrSubs = VB6.CopyArray(Veh.KeyManager.GetCurrentSubAssembliesKeys)
		If arrSubs(1) <> "" Then
			For j = 1 To UBound(arrSubs)
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(arrSubs(j)).Datatype = Mast Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TotalMastHeight = TotalMastHeight + (Veh.Components(arrSubs(j)).Quantity * Veh.Components(arrSubs(j)).Height)
					'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					NumberofMasts = NumberofMasts + Veh.Components(arrSubs(j)).Quantity
				End If
			Next 
		End If
		'get the surface area.
		If mvarDatatype = FullRig Then
			mvarSurfaceArea = ((TotalMastHeight / NumberofMasts) ^ 2) * (NumberofMasts / 2)
		ElseIf (mvarDatatype = AerialSail) Then 
			mvarSurfaceArea = ((TotalMastHeight / NumberofMasts) ^ 2) * (NumberofMasts / 2)
			
		Else 'square rigs, fore and aft rigs and aerial fore and aft rigs are .8
			mvarSurfaceArea = 0.8 * ((TotalMastHeight / NumberofMasts) ^ 2) * (NumberofMasts / 2)
		End If
		
		mvarSurfaceArea = System.Math.Round(mvarSurfaceArea, 2)
		mvarHitPoints = System.Math.Round(mvarSurfaceArea, 0)
		' Calculate the weight and cost statistics
		If mvarMaterial = "cloth" Then
			mvarWeight = mvarSurfaceArea * SailMatrix(mvarMatrixPos).WeightCloth
			mvarCost = mvarSurfaceArea * SailMatrix(mvarMatrixPos).CostCloth
		ElseIf mvarMaterial = "synthetic" Then 
			mvarWeight = mvarSurfaceArea * SailMatrix(mvarMatrixPos).WeightSynthetic
			mvarCost = mvarSurfaceArea * SailMatrix(mvarMatrixPos).CostSynthetic
		ElseIf mvarMaterial = "bioplas" Then 
			mvarWeight = mvarSurfaceArea * SailMatrix(mvarMatrixPos).WeightBioplas
			mvarCost = mvarSurfaceArea * SailMatrix(mvarMatrixPos).CostBioplas
		End If
		'apply modifier for Aerial Sails
		'If mvarDatatype = AerialSail Then
		'    mvarWeight = mvarWeight * 1.5
		'    mvarCost = mvarCost * 2
		'End If
		
		mvarWeight = System.Math.Round(mvarWeight, 2)
		mvarCost = System.Math.Round(mvarCost, 2)
		
		'calculate the motive thrust
		If mvarWind = "calm" Then
			mvarMotiveThrust = 0
		Else
			mvarMotiveThrust = System.Math.Round(SailMatrix(mvarMatrixPos).MotiveThrust * mvarWindForce * mvarSurfaceArea, 2)
		End If
		
		'produce the print output
		mvarPrintOutput = "TL" & mvarTL & " " & mvarCustomDescription & ", " & VB6.Format(mvarSurfaceArea, p_sFormat) & " sf of " & mvarMaterial & " sails, average thrust " & VB6.Format(mvarMotiveThrust, p_sFormat) & " lbs." & " (" & mvarLocation & ", HP " & mvarHitPoints & ", " & VB6.Format(mvarWeight, p_sFormat) & " lbs., " & "$" & VB6.Format(mvarCost, p_sFormat) & ")." & mvarComment
		
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