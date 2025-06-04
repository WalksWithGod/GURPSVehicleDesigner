Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsArmMotor_NET.clsArmMotor")> Public Class clsArmMotor
	
	Private mvarWeight As Double
	Private mvarVolume As Double
	Private mvarCost As Double
	Private mvarBadGrip As Boolean
	Private mvarPoorCoordination As Boolean
	Private mvarCheap As Boolean
	Private mvarExtendable As Boolean
	Private mvarStriker As Boolean
	Private mvarST As Integer
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
	
	
	
	
	
	Public Property ST() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.ST
			ST = mvarST
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.ST = 5
			mvarST = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Striker() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Striker
			Striker = mvarStriker
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Striker = 5
			mvarStriker = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Extendable() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Extendable
			Extendable = mvarExtendable
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Extendable = 5
			mvarExtendable = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Cheap() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Cheap
			Cheap = mvarCheap
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Cheap = 5
			mvarCheap = Value
		End Set
	End Property
	
	
	
	
	
	Public Property PoorCoordination() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PoorCoordination
			PoorCoordination = mvarPoorCoordination
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PoorCoordination = 5
			mvarPoorCoordination = Value
		End Set
	End Property
	
	
	
	
	
	Public Property BadGrip() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.BadGrip
			BadGrip = mvarBadGrip
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.BadGrip = 5
			mvarBadGrip = Value
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
	
	
	Public Function LocationCheck() As Boolean
		Dim TempCheck As Boolean
		Dim InstallPoint As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallPoint = Veh.Components(mvarParent).Datatype
		
		If InstallPoint = Arm Then
			TempCheck = True
		Else
			modHelper.InfoPrint(1, "Arm Motors must be placed in Arms.")
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
		mvarRuggedized = False
		
		mvarST = 12
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
		
		
	End Sub
	
	Public Sub GetMatrixIndex()
		Dim i As Short
		If mvarDatatype = 0 Then Exit Sub
		
		mvarMatrixPos = 0 'init the counter
		
		mvarLocation = GetLocation
		
		For i = 1 To UBound(ArmMotorMatrix)
			If ArmMotorMatrix(i).ID = mvarDatatype Then
				' arm motors are dependant on the ARM's (the parents) TL and NOT the vehicles TL.
				' arm motors do not have their own TL either.
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ArmMotorMatrix(i).TL >= Veh.Components(Parent).TL Then
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
		'Note:  Must call Getmatrixindex here because we need to update
		'the Counter in case the parent TL has changed
		GetMatrixIndex()
		
		Dim WeightMod As Single
		Dim VolumeMod As Single
		Dim CostMod As Single
		Dim RugMod As Single
		Dim RugHitMod As Short
		Dim sPrint1 As String
		Dim sPrint2 As String
		
		'set the ruggedized modifiers
		If mvarRuggedized Then
			RugMod = 1.5
			RugHitMod = 2
		Else
			RugMod = 1
			RugHitMod = 1
		End If
		
		'inti variables
		WeightMod = 1
		VolumeMod = 1
		CostMod = 1
		
		'bad grip option modifier
		If mvarBadGrip Then
			CostMod = 0.5
		End If
		
		'Cheap option modifier
		If mvarCheap Then
			WeightMod = 2
			VolumeMod = 2
			CostMod = CostMod * 0.5
		End If
		
		'Extendable option modifier
		If mvarExtendable Then
			WeightMod = WeightMod * 2
			VolumeMod = VolumeMod * 2
			CostMod = CostMod * 2
		End If
		
		'poor coordination option modifier
		If mvarPoorCoordination Then
			CostMod = CostMod * 0.5
		End If
		
		'Striker option modifier
		If mvarStriker Then
			WeightMod = WeightMod * 0.5
			VolumeMod = VolumeMod * 0.5
			CostMod = CostMod * 0.2
		End If
		
		'calc final stats
		mvarWeight = System.Math.Round(mvarST * RugMod * WeightMod * ArmMotorMatrix(mvarMatrixPos).Weight, 2)
		mvarCost = System.Math.Round(mvarST * RugMod * CostMod * ArmMotorMatrix(mvarMatrixPos).Cost, 2)
		mvarVolume = System.Math.Round(mvarST * RugMod * VolumeMod * ArmMotorMatrix(mvarMatrixPos).Volume, 2)
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
		
		mvarPowerReqt = System.Math.Round(mvarST / ArmMotorMatrix(mvarMatrixPos).Power, 2)
		
		'update the stats for its parent
		QueryParent()
		
		'produce the print output
		If mvarRuggedized Then
			sPrint1 = "ruggedized "
		Else
			sPrint1 = ""
		End If
		
		If mvarBadGrip Then
			sPrint2 = ", " & "bad grip"
		End If
		
		If mvarCheap Then
			sPrint2 = sPrint2 & ", " & "cheap"
		End If
		
		If mvarExtendable Then
			sPrint2 = sPrint2 & ", extendable"
		End If
		
		If mvarPoorCoordination Then
			sPrint2 = sPrint2 & ", poor coordination"
		End If
		
		If mvarStriker Then
			sPrint2 = sPrint2 & ", striker"
		End If
		
		mvarPrintOutput = "ST " & VB6.Format(mvarST) & " " & sPrint1 & mvarCustomDescription & sPrint2 & " (" & mvarLocation & ", HP " & mvarHitPoints & ", " & VB6.Format(mvarWeight, p_sFormat) & " lbs., " & VB6.Format(mvarVolume, p_sFormat) & " cf., " & "$" & VB6.Format(mvarCost, p_sFormat) & ", " & VB6.Format(mvarPowerReqt, p_sFormat) & " kW)." & mvarComment
		
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