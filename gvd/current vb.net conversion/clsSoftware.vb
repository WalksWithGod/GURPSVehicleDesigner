Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsSoftware_NET.clsSoftware")> Public Class clsSoftware
	
	Private mvarTL As Short
	Private mvarCost As Double
	Private mvarComplexity As Integer
	Private mvarGigabytes As Single
	Private mvarSkillPoints As Integer
	Private mvarLocation As String
	Private mvarParent As String
	Private mvarKey As String
	Private mvarBonusSkillPoints As Integer
	
	Private mvarDatatype As Short
	Private mvarParentDatatype As Short
	Private mvarDescription As String
	Private mvarCustomDescription As String
	Private mvarCustom As Boolean
	Private mvarQuantity As Integer
	'Note:  quantity is disabled for this.  If users request, i can re-enable so leave this property here for now
	
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
	
	
	
	Public Property SkillPoints() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.SkillPoints
			SkillPoints = mvarSkillPoints
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.SkillPoints = 5
			mvarSkillPoints = Value
		End Set
	End Property
	
	
	
	Public Property BonusSkillPoints() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.BonusSkillPoints
			BonusSkillPoints = mvarBonusSkillPoints
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.BonusSkillPoints = 5
			mvarBonusSkillPoints = Value
		End Set
	End Property
	
	
	
	Public Property Gigabytes() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Gigabytes
			Gigabytes = mvarGigabytes
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Gigabytes = 5
			mvarGigabytes = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Complexity() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Complexity
			Complexity = mvarComplexity
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Complexity = 5
			mvarComplexity = Value
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
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If TypeOf Veh.Components(Parent) Is clsComputer Then
			TempCheck = True
		Else
			modHelper.InfoPrint(1, "Software must be installed in a Computer.")
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
		mvarSkillPoints = 1
		mvarGigabytes = 1
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
			Case CartographySoftware
				
			Case ComputerNavigationSoftware
				
			Case DamageControlSoftware
				
			Case DatalinkSoftware
				
			Case FireDirectionSoftware
				
			Case DatabaseSoftware
				
			Case GunnerSoftware
				
			Case PersonalitySimulationSoftwareFull
				
			Case PersonalitySimulationLimited
				
			Case RobotSkillProgramsPhysical
				
			Case RobotSkillProgramsMental
				
			Case RoutineVehicleOperationSoftwarePilot
				
			Case RoutineVehicleOperationSoftwareOther
				
			Case TargetingSoftware
				
			Case TransmissionProfilingSoftware
				
			Case HoloventureProgram
				
				
		End Select
		
	End Sub
	
	Public Sub GetMatrixIndex()
		Dim i As Short
		If mvarDatatype = 0 Then Exit Sub
		mvarMatrixPos = 0 'init the counter
		For i = 1 To UBound(SoftwareMatrix)
			If SoftwareMatrix(i).ID = mvarDatatype Then
				If SoftwareMatrix(i).TL >= mvarTL Then
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
		
		Dim basecomplex As Integer
		Dim TempCost As Single
		Dim sPrint1 As String
		
		mvarLocation = GetLocation
		
		'store the base complexity and skillpoint value
		basecomplex = SoftwareMatrix(mvarMatrixPos).Complexity
		mvarSkillPoints = SoftwareMatrix(mvarMatrixPos).BonusSkill
		
		Select Case mvarDatatype
			
			Case DatabaseSoftware
				TempCost = mvarGigabytes * SoftwareMatrix(mvarMatrixPos).Cost
				
			Case CartographySoftware, ComputerNavigationSoftware, DatalinkSoftware, TransmissionProfilingSoftware, HoloventureProgram, PersonalitySimulationSoftwareFull, PersonalitySimulationLimited, RoutineVehicleOperationSoftwarePilot, RoutineVehicleOperationSoftwareOther
				
				TempCost = SoftwareMatrix(mvarMatrixPos).Cost
				
			Case FireDirectionSoftware, TargetingSoftware, DamageControlSoftware, GunnerSoftware
				
				basecomplex = basecomplex + mvarBonusSkillPoints 'increase complexity by 1 point for each user added bonus skill
				
				TempCost = SoftwareMatrix(mvarMatrixPos).Cost
				If mvarBonusSkillPoints > 0 Then TempCost = TempCost * 2 ^ mvarBonusSkillPoints 'each +1 in complexity doubles cost
				
				
			Case RobotSkillProgramsPhysical, RobotSkillProgramsMental
				
				'get the complexity based on the skill points
				If mvarBonusSkillPoints < 1 Then
					basecomplex = 1
				ElseIf mvarBonusSkillPoints = 1 Then 
					basecomplex = 2
				ElseIf mvarBonusSkillPoints = 2 Then 
					basecomplex = 3
				ElseIf mvarBonusSkillPoints <= 4 Then 
					basecomplex = 4
				ElseIf mvarBonusSkillPoints <= 8 Then 
					basecomplex = 5
				Else
					basecomplex = Fix((mvarBonusSkillPoints - 8) / 8)
					If basecomplex < (mvarBonusSkillPoints - 8 / 8) Then basecomplex = basecomplex + 1
					basecomplex = basecomplex + 5
				End If
				
				TempCost = SoftwareMatrix(mvarMatrixPos).Cost * mvarBonusSkillPoints
				If mvarBonusSkillPoints > 8 And mvarBonusSkillPoints < 20 Then
					TempCost = TempCost * 2.5
				ElseIf mvarSkillPoints > 20 Then 
					TempCost = TempCost * 5
				End If
				
		End Select
		mvarSkillPoints = mvarSkillPoints + mvarBonusSkillPoints
		mvarComplexity = basecomplex
		mvarCost = System.Math.Round(TempCost, 2)
		
		'produce the print output
		
		If mvarSkillPoints <> 0 Then
			sPrint1 = ", skill bonus +" & VB6.Format(mvarSkillPoints)
		ElseIf mvarDatatype = DatabaseSoftware Then 
			sPrint1 = ", " & mvarGigabytes & " gig"
		End If
		
		mvarPrintOutput = " TL" & mvarTL & " " & mvarCustomDescription & " (" & "$" & VB6.Format(mvarCost, p_sFormat) & ", complexity " & VB6.Format(mvarComplexity) & sPrint1 & ")." & mvarComment
		
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