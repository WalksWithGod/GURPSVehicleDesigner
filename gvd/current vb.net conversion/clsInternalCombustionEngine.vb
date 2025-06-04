Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsInternalCombustionEngine_NET.clsInternalCombustionEngine")> Public Class clsInternalCombustionEngine
	
	Private mvarTL As Short
	Private mvarWeight As Double
	Private mvarCost As Double
	Private mvarFuelConsumption As Single
	Private mvarVolume As Double
	Private mvarLocation As String
	Private mvarParent As String
	Private mvarKey As String
	Private mvarDR As Integer
	Private mvarRuggedized As Boolean
	Private mvarSurfaceArea As Double
	Private mvarHitPoints As Double
	Private mvarPowerConsumed As Single
	
	Private mvarDatatype As Short
	Private mvarParentDatatype As Short
	Private mvarDescription As String
	Private mvarCustomDescription As String
	Private mvarCustom As Boolean
	Private mvarQuantity As Short
	Private mvarFuelType As String
	Private mvarUnModifiedOutput As Single
	Private mvarOutput As Single
	Private mvarDesiredOutput As Single
	Private mvarEndurance As Single
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
	
	
	
	
	Public Property Endurance() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Endurance
			Endurance = mvarEndurance
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Endurance = 5
			mvarEndurance = Value
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
	
	
	
	
	Public Property UnModifiedOutput() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.UnModifiedOutput
			UnModifiedOutput = mvarUnModifiedOutput
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.UnModifiedOutput = 5
			mvarUnModifiedOutput = Value
		End Set
	End Property
	
	
	
	Public Property PowerConsumed() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PowerConsumed
			PowerConsumed = mvarPowerConsumed
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PowerConsumed = 5
			mvarPowerConsumed = Value
		End Set
	End Property
	
	
	
	
	
	Public Property FuelType() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.FuelType
			FuelType = mvarFuelType
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.FuelType = 5
			mvarFuelType = Value
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
	
	
	
	
	
	Public Property FuelConsumption() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.FuelConsumption
			FuelConsumption = mvarFuelConsumption
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.FuelConsumption = 5
			mvarFuelConsumption = Value
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
	
	
	
	Public Property Output() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Output
			Output = mvarOutput
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Output = 5
			mvarOutput = Value
		End Set
	End Property
	
	
	Public Property DesiredOutput() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.DesiredOutput
			DesiredOutput = mvarDesiredOutput
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.DesiredOutput = 5
			mvarDesiredOutput = Value
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
	
	
	
	
	
	Public Function LocationCheck() As Boolean
		Dim TempCheck As Boolean
		Dim InstallPoint As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallPoint = Veh.Components(mvarParent).Datatype
		
		If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Pod) Or (InstallPoint = Wing) Or (InstallPoint = equipmentPod) Then
			TempCheck = True
		Else
			modHelper.InfoPrint(1, "Internal Combustion Engines must go in Body, Pod, equipment Pod or Wing.")
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
		'set the default dimension of the keychain to 1 element
		Dim mvarPowerConsumptionKeyChain(1) As Object
		Dim mvarFuelStorageKeyChain(1) As Object
		
		' set the default properties
		mvarCustom = False
		TL = gVehicleTL
		mvarRuggedized = False
		mvarQuantity = 1
		mvarUnModifiedOutput = 40
		mvarOutput = 40
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
			Case HPCeramicEngine
				
				mvarFuelType = "aviation gas"
			Case TurboHPCeramicEngine
				
				mvarFuelType = "aviation gas"
			Case SuperHPCeramicEngine
				
				mvarFuelType = "aviation gas"
			Case HPGasolineEngine
				
				mvarFuelType = "gasoline"
			Case TurboHPGasolineEngine
				
				mvarFuelType = "gasoline"
			Case SuperHPGasolineEngine
				
				mvarFuelType = "gasoline"
			Case CeramicEngine
				
				mvarFuelType = "gasoline"
			Case TurboCeramicEngine
				
				mvarFuelType = "gasoline"
			Case SuperCeramicEngine
				
				mvarFuelType = "gasoline"
			Case TurboStandardDieselEngine
				
				mvarFuelType = "diesel"
			Case TurboHPDieselEngine
				
				mvarFuelType = "diesel"
			Case MarineDieselEngine
				
				mvarFuelType = "diesel"
			Case StandardDieselEngine
				
				mvarFuelType = "diesel"
			Case HPDieselEngine
				
				mvarFuelType = "diesel"
			Case GasolineEngine
				
				mvarFuelType = "gasoline"
			Case TurboGasolineEngine
				
				mvarFuelType = "gasoline"
			Case SuperGasolineEngine
				
				mvarFuelType = "gasoline"
				
			Case HydrogenCombustionEngine
				
				mvarFuelType = "hydrogen"
				
		End Select
		
		
	End Sub
	
	Public Sub GetMatrixIndex()
		Dim i As Short
		If mvarDatatype = 0 Then Exit Sub
		mvarMatrixPos = 0 'init the counter
		For i = 1 To UBound(CombustionEngineMatrix)
			If CombustionEngineMatrix(i).ID = mvarDatatype Then
				If CombustionEngineMatrix(i).TL >= mvarTL Then
					mvarMatrixPos = i
					Exit For
				Else
					mvarMatrixPos = i
				End If
			End If
		Next 
	End Sub
	
	Public Sub StatsUpdate()
		
		Dim WeightMod As Single
		Dim TrueWeight As Single
		Dim TempCost As Single
		Dim TempVolume As Single
		Dim element As Object
		Dim TempPowerConsumption As Single
		Dim TempUnmodifiedOutput As Single
		Dim i As Integer
		Dim QRugMod As Single 'combined quantity and ruggedized multipliers
		Dim RugHitMod As Short 'ruggedized hit point multiplier
		Dim sPrint1 As String
		Dim sPrint2 As String
		Dim sPrintPlural As String
		Dim sPrintPlural2 As String
		Dim sPrintPlural3 As String
		Dim sPrintPlural4 As String
		
		mvarZZInit = 1
		
		mvarLocation = GetLocation
		
		'set the ruggedized and quantity multipliers
		If mvarRuggedized Then
			QRugMod = 1.5 * mvarQuantity
			RugHitMod = 2
		Else
			QRugMod = 1 * mvarQuantity
			RugHitMod = 1
		End If
		
		
		'Dont run this StatsUpdate procedure after the component
		'has just initialized because the Datatype is still empty
		'which results in a 0 for Counter
		If mvarMatrixPos = 0 Then Exit Sub
		
		'determine if the weight is above or below 5kw and then make adjustments
		If UnModifiedOutput < 5 Then
			WeightMod = 0
			TrueWeight = CombustionEngineMatrix(mvarMatrixPos).Weight1
		Else
			WeightMod = CombustionEngineMatrix(mvarMatrixPos).Weight3
			TrueWeight = CombustionEngineMatrix(mvarMatrixPos).Weight2
		End If
		
		TrueWeight = (mvarUnModifiedOutput * TrueWeight) + WeightMod
		
		'Get Cost
		TempCost = TrueWeight * CombustionEngineMatrix(mvarMatrixPos).Cost
		
		'Get Volume
		TempVolume = TrueWeight / CombustionEngineMatrix(mvarMatrixPos).Volume
		
		' Set Effective Power before any modifiers
		TempUnmodifiedOutput = mvarUnModifiedOutput
		mvarDesiredOutput = System.Math.Round(TempUnmodifiedOutput, 2)
		
		'only propane fuel causes change in fuel consumption (i.e. alcohol conversion does not)
		If mvarFuelType = "propane" Then
			mvarFuelConsumption = System.Math.Round(1.5 * CombustionEngineMatrix(mvarMatrixPos).FuelUsed * TempUnmodifiedOutput, 2)
			TrueWeight = TrueWeight * 1.1
			TempCost = TempCost * 1.1
			TempVolume = TempVolume * 1.1
		ElseIf (mvarFuelType = "ethanol") Or (mvarFuelType = "methanol") Then  'this is for multi fuels
			If (mvarDatatype = CeramicEngine) Or (mvarDatatype = TurboCeramicEngine) Or (mvarDatatype = SuperCeramicEngine) Then
				mvarFuelConsumption = System.Math.Round(CombustionEngineMatrix(mvarMatrixPos).AlcoholFuelMod * CombustionEngineMatrix(mvarMatrixPos).FuelUsed * TempUnmodifiedOutput, 2)
			Else ' this is for alcohol conversions
				mvarFuelConsumption = System.Math.Round(CombustionEngineMatrix(mvarMatrixPos).FuelUsed * TempUnmodifiedOutput, 2)
				mvarDesiredOutput = System.Math.Round(TempUnmodifiedOutput * CombustionEngineMatrix(mvarMatrixPos).AlcoholModifier, 2)
			End If
		Else 'for all other engine types
			mvarFuelConsumption = System.Math.Round(CombustionEngineMatrix(mvarMatrixPos).FuelUsed * TempUnmodifiedOutput, 2)
		End If
		
		' check for a Nitrous Oxide booster child component and update effective motive power to 20%
		' nitrous only boosts internal combustion engines other than Hydrogen!
		If mvarDatatype = HydrogenCombustionEngine Then
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For	Each element In Veh.Components
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.Datatype = NitrousOxideBooster Then
					'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.LogicalParent = mvarKey Then
						mvarDesiredOutput = System.Math.Round(mvarDesiredOutput * 1.2, 2)
						Exit For
					End If
				End If
			Next element
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
		mvarOutput = mvarDesiredOutput * mvarQuantity
		mvarFuelConsumption = mvarFuelConsumption * mvarQuantity
		
		
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
		
		sPrint2 = ", uses " & sPrintPlural4 & VB6.Format(mvarFuelConsumption, p_sFormat) & " gph " & mvarFuelType
		
		
		mvarPrintOutput = NumericToString(mvarQuantity) & " TL" & mvarTL & " " & VB6.Format(mvarDesiredOutput, p_sFormat) & " kW " & sPrint1 & mvarCustomDescription & sPrintPlural & " (" & mvarLocation & ", HP " & mvarHitPoints & sPrintPlural3 & ", " & sPrintPlural4 & VB6.Format(mvarWeight, p_sFormat) & " lbs., " & VB6.Format(mvarVolume, p_sFormat) & " cf., " & "$" & VB6.Format(mvarCost, p_sFormat) & sPrint2 & ")." & mvarComment
		
		
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