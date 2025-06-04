Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsArmor_NET.clsArmor")> Public Class clsArmor
	
	
	
	
	
	'local variable(s) to hold property value(s)
	Private mvarMaterial As String
	Private mvarQuality As String
	Private mvarDR As Integer
	Private mvarRAP As Boolean
	Private mvarElectrified As Boolean
	Private mvarThermal As Boolean
	Private mvarRadiation As Boolean
	Private mvarCoating As String
	Private mvarPD As Short
	Private mvarWeight As Double
	Private mvarCost As Double
	Private mvarTL As Short
	'Private mvarCost1 as single
	'Private mvarCost2 as single
	'Private mvarCost3 as single
	'Private mvarCost4 as single
	'Private mvarCost5 as single
	'Private mvarCost6 as single
	'Private mvarWeight1 as single
	'Private mvarWeight2 as single
	'Private mvarWeight3 as single
	'Private mvarWeight4 as single
	'Private mvarWeight5 as single
	'Private mvarWeight6 as single
	Private mvarMaterial1 As String
	Private mvarMaterial2 As String
	Private mvarMaterial3 As String
	Private mvarMaterial4 As String
	Private mvarMaterial5 As String
	Private mvarMaterial6 As String
	Private mvarQuality1 As String
	Private mvarQuality2 As String
	Private mvarQuality3 As String
	Private mvarQuality4 As String
	Private mvarQuality5 As String
	Private mvarQuality6 As String
	Private mvarEffectiveDR1 As Integer
	Private mvarEffectiveDR2 As Integer
	Private mvarEffectiveDR3 As Integer
	Private mvarEffectiveDR4 As Integer
	Private mvarEffectiveDR5 As Integer
	Private mvarEffectiveDR6 As Integer
	Private mvarDR1 As Integer
	Private mvarDR2 As Integer
	Private mvarDR3 As Integer
	Private mvarDR4 As Integer
	Private mvarDR5 As Integer
	Private mvarDR6 As Integer
	Private mvarPD1 As Short
	Private mvarPD2 As Short
	Private mvarPD3 As Short
	Private mvarPD4 As Short
	Private mvarPD5 As Short
	Private mvarPD6 As Short
	Private mvarParent As String
	Private mvarLocation As String
	Private mvarKey As String
	Private mvarImage As Short
	Private mvarSelectedImage As Short
	Private mvarDatatype As Short
	Private mvarParentDatatype As Short
	Private mvarDescription As String
	Private mvarCustomDescription As String
	Private mvarCustom As Boolean
	Private mvarAverageDR As Single
	Private mvarAveragePD As Single
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
	
	
	
	
	Public Property Coating() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Coating
			Coating = mvarCoating
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Coating = 5
			mvarCoating = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Radiation() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Radiation
			Radiation = mvarRadiation
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Radiation = 5
			mvarRadiation = Value
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
	
	
	
	
	
	Public Property Key() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.key
			Key = mvarKey
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.key = 5
			mvarKey = Value
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
	
	
	
	
	Public Property PD() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PD
			PD = mvarPD
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PD = 5
			mvarPD = Value
		End Set
	End Property
	
	
	
	Public Property PD1() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PD1
			PD1 = mvarPD1
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PD1 = 5
			mvarPD1 = Value
		End Set
	End Property
	
	
	Public Property PD2() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PD2
			PD2 = mvarPD2
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PD2 = 5
			mvarPD2 = Value
		End Set
	End Property
	
	
	Public Property PD3() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PD3
			PD3 = mvarPD3
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PD3 = 5
			mvarPD3 = Value
		End Set
	End Property
	
	
	Public Property PD4() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PD4
			PD4 = mvarPD4
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PD4 = 5
			mvarPD4 = Value
		End Set
	End Property
	
	
	Public Property PD5() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PD5
			PD5 = mvarPD5
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PD5 = 5
			mvarPD5 = Value
		End Set
	End Property
	
	
	Public Property PD6() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.PD6
			PD6 = mvarPD6
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.PD6 = 5
			mvarPD6 = Value
		End Set
	End Property
	
	
	
	Public Property AveragePD() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.AveragePD
			AveragePD = mvarAveragePD
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.AveragePD = 5
			mvarAveragePD = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Thermal() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Thermal
			Thermal = mvarThermal
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Thermal = 5
			mvarThermal = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Electrified() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Electrified
			Electrified = mvarElectrified
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Electrified = 5
			mvarElectrified = Value
		End Set
	End Property
	
	
	
	
	
	Public Property RAP() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.RAP
			RAP = mvarRAP
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.RAP = 5
			mvarRAP = Value
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
	
	
	
	Public Property DR1() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.DR1
			DR1 = mvarDR1
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.DR1 = 5
			mvarDR1 = Value
		End Set
	End Property
	
	
	
	Public Property DR2() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.DR2
			DR2 = mvarDR2
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.DR2 = 5
			mvarDR2 = Value
		End Set
	End Property
	
	
	
	Public Property DR3() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.DR3
			DR3 = mvarDR3
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.DR3= 5
			mvarDR3 = Value
		End Set
	End Property
	
	
	
	Public Property DR4() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.DR4
			DR4 = mvarDR4
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.DR4 = 5
			mvarDR4 = Value
		End Set
	End Property
	
	
	
	Public Property DR5() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.DR5
			DR5 = mvarDR5
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.DR5 = 5
			mvarDR5 = Value
		End Set
	End Property
	
	
	
	Public Property DR6() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.DR6
			DR6 = mvarDR6
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.DR6 = 5
			mvarDR6 = Value
		End Set
	End Property
	
	
	
	Public Property AverageDR() As Single
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.AverageDR
			AverageDR = mvarAverageDR
		End Get
		Set(ByVal Value As Single)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.AverageDR = 5
			mvarAverageDR = Value
		End Set
	End Property
	
	
	
	Public Property EffectiveDR1() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.EffectiveDR1
			EffectiveDR1 = mvarEffectiveDR1
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.EffectiveDR1 = 5
			mvarEffectiveDR1 = Value
		End Set
	End Property
	
	
	
	Public Property EffectiveDR2() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.EffectiveDR2
			EffectiveDR2 = mvarEffectiveDR2
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.EffectiveDR2 = 5
			mvarEffectiveDR2 = Value
		End Set
	End Property
	
	
	
	Public Property EffectiveDR3() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.EffectiveDR3
			EffectiveDR3 = mvarEffectiveDR3
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.EffectiveDR3= 5
			mvarEffectiveDR3 = Value
		End Set
	End Property
	
	
	
	Public Property EffectiveDR4() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.EffectiveDR4
			EffectiveDR4 = mvarEffectiveDR4
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.EffectiveDR4 = 5
			mvarEffectiveDR4 = Value
		End Set
	End Property
	
	
	
	Public Property EffectiveDR5() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.EffectiveDR5
			EffectiveDR5 = mvarEffectiveDR5
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.EffectiveDR5 = 5
			mvarEffectiveDR5 = Value
		End Set
	End Property
	
	
	
	Public Property EffectiveDR6() As Integer
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.EffectiveDR6
			EffectiveDR6 = mvarEffectiveDR6
		End Get
		Set(ByVal Value As Integer)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.EffectiveDR6 = 5
			mvarEffectiveDR6 = Value
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
	
	
	
	Public Property Quality1() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Quality1
			Quality1 = mvarQuality1
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Quality1 = 5
			mvarQuality1 = Value
		End Set
	End Property
	
	
	Public Property Quality2() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Quality2
			Quality2 = mvarQuality2
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Quality2 = 5
			mvarQuality2 = Value
		End Set
	End Property
	
	
	Public Property Quality3() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Quality3
			Quality3 = mvarQuality3
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Quality3 = 5
			mvarQuality3 = Value
		End Set
	End Property
	
	
	Public Property Quality4() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Quality4
			Quality4 = mvarQuality4
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Quality4 = 5
			mvarQuality4 = Value
		End Set
	End Property
	
	
	Public Property Quality5() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Quality5
			Quality5 = mvarQuality5
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Quality5 = 5
			mvarQuality5 = Value
		End Set
	End Property
	
	
	Public Property Quality6() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Quality6
			Quality6 = mvarQuality6
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Quality6 = 5
			mvarQuality6 = Value
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
	
	
	
	
	Public Property Material1() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Material1
			Material1 = mvarMaterial1
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Material1 = 5
			mvarMaterial1 = Value
		End Set
	End Property
	
	
	
	Public Property Material2() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Material2
			Material2 = mvarMaterial2
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Material2 = 5
			mvarMaterial2 = Value
		End Set
	End Property
	
	
	
	Public Property Material3() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Material3
			Material3 = mvarMaterial3
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Material3 = 5
			mvarMaterial3 = Value
		End Set
	End Property
	
	
	
	Public Property Material4() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Material4
			Material4 = mvarMaterial4
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Material4 = 5
			mvarMaterial4 = Value
		End Set
	End Property
	
	
	
	Public Property Material5() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Material5
			Material5 = mvarMaterial5
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Material5 = 5
			mvarMaterial5 = Value
		End Set
	End Property
	
	
	
	Public Property Material6() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Material6
			Material6 = mvarMaterial6
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Material = 5
			mvarMaterial6 = Value
		End Set
	End Property
	
	Public Function LocationCheck() As Boolean
		Dim TempCheck As Boolean
		Dim InstallPoint As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallPoint = Veh.Components(mvarParent).Datatype
		
		'NOTE: Because Layering of armor is allowed,  user can have Overall armor along with location armor for instance on individual components!
		'So there is no need to check for existance of other armor types
		Select Case mvarDatatype
			
			Case ArmorComplexFacing, ArmorBasicFacing
				If (InstallPoint = Body) Or (InstallPoint = Superstructure) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Then
					TempCheck = True
				Else
					modHelper.InfoPrint(1, "Armor by Facing can only be applied to a Body, Superstructure, Turret or Popturret.")
					TempCheck = False
				End If
			Case ArmorOverall
				If InstallPoint = Body Then
					TempCheck = True
				Else
					modHelper.InfoPrint(1, "Overall Armor can only be placed on the hull.")
					TempCheck = False
				End If
			Case ArmorLocation
				If (InstallPoint = Body) Or (InstallPoint = Mast) Or (InstallPoint = Skid) Or (InstallPoint = Pod) Or (InstallPoint = Hovercraft) Or (InstallPoint = Hydrofoil) Or (InstallPoint = Wheel) Or (InstallPoint = Track) Or (InstallPoint = Leg) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = AutogyroRotor) Or (InstallPoint = CARotor) Or (InstallPoint = TTRotor) Or (InstallPoint = MMRotor) Or (InstallPoint = Gasbag) Or (InstallPoint = Superstructure) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Then
					TempCheck = True
				Else
					modHelper.InfoPrint(1, "Location armor can only be applied to valid subassemblies.")
					TempCheck = False
				End If
				
			Case ArmorOpenFrame
				Select Case InstallPoint
					Case Body, Mast, Skid, Pod, Hovercraft, Hydrofoil, Wheel, Track, Leg, Arm, Wing, AutogyroRotor, CARotor, TTRotor, MMRotor, Gasbag, Superstructure, Turret, Popturret
						
						TempCheck = True
						
					Case Else
						modHelper.InfoPrint(1, "Location armor can only be applied to valid subassemblies.")
						TempCheck = False
				End Select
				
			Case ArmorComponent
				Select Case InstallPoint
					Case StoneThrower, BoltThrower, RepeatingBoltThrower, MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling, BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam, FlameThrower, WaterCannon, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, RevolverLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher, PartialStabilizationGear, FullStabilizationGear, UniversalMount, CasemateMount, DoorMount, Cyberslave, AntiBlastMagazine, HardPoint, WeaponBay, Cargo, equipmentPod
						
						TempCheck = True
						
					Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, FlexibodyDrivetrain, TrackedDrivetrain, LegDrivetrain, CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, OrnithopterDrivetrain, AerialPropeller, DuctedFan, PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel, MagLevLifter, Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam, StandardThruster, SuperThruster, MegaThruster, LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion, SolidRocketEngine, OrionEngine, TeleportationDrive, Hyperdrive, JumpDrive, WarpDrive, QuantumConveyor, SubQuantumConveyor, TwoQuantumConveyor, ContraGravGenerator, RadioDirectionFinder, RadioCommunicator, TightBeamRadio, VLFRadio, CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator, Headlight, Searchlight, InfraredSearchlight, AstronomicalInstruments, Telescope, lightAmplification, LowlightTV, ExtendableSensorPeriscope
						
						TempCheck = True
						
					Case Radar, Ladar, NavigationalRadar, AntiCollisionRadar, AESA, LowResImagingRadar, HiResImagingRadar, ActiveSonar, PassiveSonar, PassiveInfrared, Thermograph, PassiveRadar, PESA, Geophone, MAD, MultiScanner, ChemScanner, RadScanner, BioScanner, GravScanner, RangingSoundDetector, SurveillanceSoundDetector, MeteorologicalInstruments, LowResPlanetarySurveyArray, MedResPlanetarySurveyArray, HighResPlanetarySurveyArray, SoundSystem, FlightRecorder, VehicleCamera, DigitalVehicleCamera, ReconCamera, DigitalReconCamera, NavigationInstruments, AutoPilot, IFF, Transponder, INS, GPS, MilitaryGPS, TFR, ImprovedOpticalBombSight, AdvancedOpticalBombSight, OpticalBombSight, FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, LaserDesignator, LaserSpotTracker, RadarDetector, LaserSensor, LaserRadarDetector, AreaRadarJammer, DeceptiveRadarJammer, InfraredJammer, ChaffDecoyDischarger, SmokeDecoyDischarger, FlareDecoyDischarger, SonarDecoyDischarger, HotSmokeDecoyDischarger, PrismDecoyDischarger, BlackOutGasDecoyDischarger, RadarReflector, BlipEnhancer, TEMPEST, MacroFrame, MainFrame, MicroFrame, MiniComputer, SmallComputer, Terminal, SurgicalInterface, InterfaceWeb, AutoInterfaceWeb, SocketInterface, NeuralInductionField
						
						TempCheck = True
						
					Case ArmMotor, FireExtinguisherSystem, FullFireSuppressionSystem, CompactFireSuppressionSystem, BilgePump, CompleteWorkshop, MechanicWorkshop, EngineeringWorkshop, ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, MiniEngineeringWorkshop, MiniArmouryWorkshop, ExtendableLadder, Crane, Winch, PowerShovel, WreckingCrane, ForkLift, VehicularBridge, LaunchCatapult, SkyHook, Bore, SuperBore, TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet, OperatingRoom, StretcherPallet, EmergencySupportUnit, EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable, Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, MovieScreenandProjectorSmall, HoloventureZone
						
						TempCheck = True
						
					Case CargoRamp, Airlock, MembraneAirlock, Forcelock, PassageTube, ArmoredPassageTube, TeleportProjector, BrigsandRestraints, BurglarAlarm, HighSecurityAlarm, MutableLicensePlate, OilSprayer, PaintSprayer, SmokeScreen, SpikeDropper, VehicleBay, HangerBay, DryDock, SpaceDock, ExternalCradle, ArrestorHook, VehicularParachute, RefuellingProbe, RefuellingDrogue, FuelElectrolysisSystem, HydrogenFuelScoop, AtmosphereProcessor, NuclearDamper, SmallRealityStabilizer, MediumRealityStabilizer, HeavyRealityStabilizer, ModularSocket, Module_Renamed
						
						TempCheck = True
						
					Case PrimitiveManeuverControl, ElectronicDivingControl, ComputerizedDivingControl, MechanicalManeuverControl, ElectronicManeuverControl, ComputerizedManeuverControl, MechanicalDivingControl, CrampedCrewStation, NormalCrewStation, RoomyCrewStation, CycleCrewStation, HarnessCrewStation, CrampedSeat, NormalSeat, RoomySeat, CrampedStandingRoom, NormalStandingRoom, RoomyStandingRoom, CycleSeat, Hammock, Bunk, Cabin, LuxuryCabin, Suite, LuxurySuite, SmallGalley, TotalLifeSystem, ArtificialGravityUnit, EnvironmentalControl, NBCKit, LimitedLifeSystem, FullLifeSystem, EjectionSeat, CrewEscapeCapsule, Airbag, CrashWeb, WombTank, GravityWeb, GravCompensator
						
						
						TempCheck = True
						
						
					Case MuscleEngine, GasolineEngine, HPGasolineEngine, TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, SuperHPGasolineEngine, StandardDieselEngine, TurboStandardDieselEngine, MarineDieselEngine, HPDieselEngine, TurboHPDieselEngine, CeramicEngine, TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, TurboHPCeramicEngine, SuperHPCeramicEngine, HydrogenCombustionEngine, EarlySteamEngine, ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine, StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, StandardMHDTurbine, HPMHDTurbine, FuelCell, FissionReactor, RTGReactor, NPU, FusionReactor, AntimatterReactor, TotalConversionPowerPlant, CosmicPowerPlant, Soulburner, ElementalFurnace, ManaEngine, Carnivore, Herbivore, Omnivore, Vampire, ClockWork, LeadAcidBattery, AdvancedBattery, Flywheel, RechargeablePowerCell, PowerCell, AntiMatterBay, CoalBunker, WoodBunker, StandardTank, lightTank, UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, UltralightSelfSealingTank, Snorkel, ElectricContactPower, LaserBeamedPowerReceiver, MaserBeamedPowerReceiver, NitrousOxideBooster
						
						TempCheck = True
					Case Else
						TempCheck = False
						modHelper.InfoPrint(1, "Component armor cannot be applied to this vehicle component.")
						
				End Select
				
			Case ArmorGunShield
				If InstallPoint = OpenMount Then
					TempCheck = True
				Else
					modHelper.InfoPrint(1, "Gun Shields can only be applied to Open Mounts.")
					TempCheck = False
				End If
				
			Case ArmorWheelGuard
				If (InstallPoint = Wheel) Or (InstallPoint = Track) Or (InstallPoint = Hovercraft) Then
					TempCheck = True
				Else
					modHelper.InfoPrint(1, "Wheel Guards and Armored Skirts can only be applied to Wheels, Tracks, and Hovercraft subassemblies.")
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
		' in the case of Armor, its logical parent is always the same as its parent
		' armor cannot be attached to GROUP components
		mvarLogicalParent = mvarParent
	End Sub
	
	Public Sub Init()
		
		
		Select Case mvarDatatype
			
			Case ArmorComplexFacing
				Material = "wood"
				Quality = "standard"
				DR = 1
				
				Material1 = "wood"
				Material2 = "wood"
				Material3 = "wood"
				Material4 = "wood"
				Material5 = "wood"
				Material6 = "wood"
				Quality1 = "standard"
				Quality2 = "standard"
				Quality3 = "standard"
				Quality4 = "standard"
				Quality5 = "standard"
				Quality6 = "standard"
				DR1 = 1
				DR2 = 1
				DR3 = 1
				DR4 = 1
				DR5 = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarParent).Datatype = Body Then
					DR6 = 1
				Else
					DR6 = 0
				End If
			Case ArmorBasicFacing
				
				Material = "wood"
				Quality = "standard"
				DR1 = 1
				DR2 = 1
				DR3 = 1
				DR4 = 1
				DR5 = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(mvarParent).Datatype = Body Then
					DR6 = 1
				Else
					DR6 = 0
				End If
			Case ArmorOpenFrame
				
				Material = "wood"
				Quality = "standard"
				DR = 1
				
			Case ArmorGunShield
				
				Material = "wood"
				Quality = "standard"
				DR = 1
				
			Case ArmorLocation
				
				Material = "wood"
				Quality = "standard"
				DR = 1
				
			Case ArmorComponent
				
				Material = "wood"
				Quality = "standard"
				DR = 1
				
			Case ArmorOverall
				
				Material = "wood"
				Quality = "standard"
				DR = 1
				
			Case ArmorWheelGuard
				Material = "wood"
				Quality = "standard"
				DR = 1
				
		End Select
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarTL = Veh.Components(mvarParent).TL
		
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		mvarCoating = "none"
		mvarRadiation = False
		mvarThermal = False
		mvarRAP = False
		mvarElectrified = False
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Public Sub StatsUpdate()
		Dim component As Short
		Dim SlopeR As String
		Dim SlopeL As String
		Dim SlopeF As String
		Dim SlopeB As String
		'UPGRADE_WARNING: Lower bound of array sCompare was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim sCompare(6) As String
		Dim sSides() As String
		Dim i As Integer
		Dim j As Integer
		Dim count As Integer
		
		mvarZZInit = 1
		mvarPrintOutput = "" ' reinit this var
		
		mvarLocation = GetLocation
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		component = Veh.Components(mvarParent).Datatype
		
		If (mvarDatatype = ArmorBasicFacing) Or (mvarDatatype = ArmorComplexFacing) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SlopeR = Veh.Components(mvarParent).SlopeR
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SlopeL = Veh.Components(mvarParent).SlopeL
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SlopeF = Veh.Components(mvarParent).SlopeF
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SlopeB = Veh.Components(mvarParent).SlopeB
			
			
			CalcByFacingArmorWeightCost()
			'get the PD
			' Get the PD for each side (used for both Complex and Basic)
			mvarPD1 = CalcPD(mvarDR1, SlopeR, mvarMaterial1)
			mvarPD2 = CalcPD(mvarDR2, SlopeL, mvarMaterial2)
			mvarPD3 = CalcPD(mvarDR3, SlopeF, mvarMaterial3)
			mvarPD4 = CalcPD(mvarDR4, SlopeB, mvarMaterial4)
			mvarPD5 = CalcPD(mvarDR5, "none", mvarMaterial5) 'the top and underside dont have slope
			mvarPD6 = CalcPD(mvarDR6, "none", mvarMaterial6)
			
			'get the effective DR
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcEffectiveDR(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEffectiveDR1 = CalcEffectiveDR(mvarDR1, SlopeR)
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcEffectiveDR(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEffectiveDR2 = CalcEffectiveDR(mvarDR2, SlopeL)
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcEffectiveDR(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEffectiveDR3 = CalcEffectiveDR(mvarDR3, SlopeF)
			'UPGRADE_WARNING: Couldn't resolve default property of object CalcEffectiveDR(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEffectiveDR4 = CalcEffectiveDR(mvarDR4, SlopeB)
			mvarEffectiveDR5 = mvarDR5
			mvarEffectiveDR6 = mvarDR6
			
			
		ElseIf mvarDatatype = ArmorLocation Then 
			
			CalcArmorWeightCost()
			If component = Body Or component = Superstructure Or component = Turret Or component = Popturret Then
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SlopeR = Veh.Components(mvarParent).SlopeR
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SlopeL = Veh.Components(mvarParent).SlopeL
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SlopeF = Veh.Components(mvarParent).SlopeF
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SlopeB = Veh.Components(mvarParent).SlopeB
				' Get the PD for each side (used for both Complex and Basic)
				mvarPD1 = CalcPD(mvarDR, SlopeR, mvarMaterial)
				mvarPD2 = CalcPD(mvarDR, SlopeL, mvarMaterial)
				mvarPD3 = CalcPD(mvarDR, SlopeF, mvarMaterial)
				mvarPD4 = CalcPD(mvarDR, SlopeB, mvarMaterial)
				mvarPD5 = CalcPD(mvarDR, "none", mvarMaterial)
				mvarPD6 = CalcPD(mvarDR, "none", mvarMaterial)
			Else
				mvarPD = CalcPD(mvarDR, "none", mvarMaterial)
			End If
		Else
			CalcArmorWeightCost()
			mvarPD = CalcPD(mvarDR, "none", mvarMaterial)
		End If
		
		'print output
		If mvarRAP Then
		End If
		If mvarElectrified Then
		End If
		If mvarThermal Then
		End If
		If mvarRadiation Then
		End If
		If mvarCoating <> "none" Then
		End If
		If mvarPD Then
		End If
		
		Dim counter As Short
		Select Case mvarDatatype
			Case ArmorComplexFacing
				'can have different everything
				'get each side
				sCompare(1) = " PD " & VB6.Format(mvarPD1) & ", DR " & VB6.Format(mvarDR1) & " " & mvarQuality1 & " " & mvarMaterial1 & ". "
				sCompare(2) = " PD " & VB6.Format(mvarPD2) & ", DR " & VB6.Format(mvarDR2) & " " & mvarQuality2 & " " & mvarMaterial2 & ". "
				sCompare(3) = " PD " & VB6.Format(mvarPD3) & ", DR " & VB6.Format(mvarDR3) & " " & mvarQuality3 & " " & mvarMaterial3 & ". "
				sCompare(4) = " PD " & VB6.Format(mvarPD4) & ", DR " & VB6.Format(mvarDR4) & " " & mvarQuality4 & " " & mvarMaterial4 & ". "
				sCompare(5) = " PD " & VB6.Format(mvarPD5) & ", DR " & VB6.Format(mvarDR5) & " " & mvarQuality5 & " " & mvarMaterial5 & ". "
				sCompare(6) = " PD " & VB6.Format(mvarPD6) & ", DR " & VB6.Format(mvarDR6) & " " & mvarQuality6 & " " & mvarMaterial6 & ". "
				'move the first side (the Right) into the sSides array
				ReDim sSides(2, 1)
				sSides(1, 1) = sCompare(1)
				sSides(2, 1) = "R"
				count = 1
				'compare each side to see if it can be grouped with another
				For j = 2 To 6
					counter = count
					For i = 1 To counter
						If sCompare(j) = sSides(1, i) Then
							sSides(1, i) = sCompare(j)
							sSides(2, i) = sSides(2, i) & "," & GetSideLetterFromNumber(j)
						ElseIf i = count Then 
							ReDim Preserve sSides(2, i + 1)
							count = count + 1
							sSides(1, i + 1) = sCompare(j)
							sSides(2, i + 1) = GetSideLetterFromNumber(j)
						End If
					Next 
				Next 
				'get final string and include the surface options to the armor
				For i = 1 To count
					mvarPrintOutput = mvarPrintOutput & " " & sSides(2, i) & ": " & sSides(1, i)
				Next 
				mvarPrintOutput = mvarPrintOutput & " (" & VB6.Format(mvarWeight, p_sFormat) & " lbs., $" & VB6.Format(mvarCost, p_sFormat) & ")."
				
			Case ArmorBasicFacing
				'same material and quality but different DR's and PD's
				sCompare(1) = " PD " & VB6.Format(mvarPD1) & ", DR " & VB6.Format(mvarDR1) & " "
				sCompare(2) = " PD " & VB6.Format(mvarPD2) & ", DR " & VB6.Format(mvarDR2) & " "
				sCompare(3) = " PD " & VB6.Format(mvarPD3) & ", DR " & VB6.Format(mvarDR3) & " "
				sCompare(4) = " PD " & VB6.Format(mvarPD4) & ", DR " & VB6.Format(mvarDR4) & " "
				sCompare(5) = " PD " & VB6.Format(mvarPD5) & ", DR " & VB6.Format(mvarDR5) & " "
				sCompare(6) = " PD " & VB6.Format(mvarPD6) & ", DR " & VB6.Format(mvarDR6) & " "
				'move the first side (the Right) into the sSides array
				ReDim sSides(2, 1)
				sSides(1, 1) = sCompare(1)
				sSides(2, 1) = "R"
				count = 1
				'compare each side to see if it can be grouped with another
				For j = 2 To 6
					counter = count
					For i = 1 To counter
						If sCompare(j) = sSides(1, i) Then
							sSides(1, i) = sCompare(j)
							sSides(2, i) = sSides(2, i) & "," & GetSideLetterFromNumber(j)
						ElseIf i = count Then 
							ReDim Preserve sSides(2, i + 1)
							count = count + 1
							sSides(1, i + 1) = sCompare(j)
							sSides(2, i + 1) = GetSideLetterFromNumber(j)
						End If
					Next 
				Next 
				'get final string and include the surface options to the armor
				For i = 1 To count
					mvarPrintOutput = mvarPrintOutput & " " & sSides(2, i) & ": " & sSides(1, i)
				Next 
				mvarPrintOutput = mvarQuality & " " & mvarMaterial & mvarPrintOutput & " (" & VB6.Format(mvarWeight, p_sFormat) & " lbs., $" & VB6.Format(mvarCost, p_sFormat) & ")."
				
				
			Case ArmorLocation 'this can still have different PD's on Body and Turrets do to slope differences
				sCompare(1) = " PD " & VB6.Format(mvarPD1) & ", DR " & VB6.Format(mvarDR) & " "
				sCompare(2) = " PD " & VB6.Format(mvarPD2) & ", DR " & VB6.Format(mvarDR) & " "
				sCompare(3) = " PD " & VB6.Format(mvarPD3) & ", DR " & VB6.Format(mvarDR) & " "
				sCompare(4) = " PD " & VB6.Format(mvarPD4) & ", DR " & VB6.Format(mvarDR) & " "
				sCompare(5) = " PD " & VB6.Format(mvarPD5) & ", DR " & VB6.Format(mvarDR) & " "
				sCompare(6) = " PD " & VB6.Format(mvarPD6) & ", DR " & VB6.Format(mvarDR) & " "
				'move the first side (the Right) into the sSides array
				ReDim sSides(2, 1)
				sSides(1, 1) = sCompare(1)
				sSides(2, 1) = "R"
				count = 1
				'compare each side to see if it can be grouped with another
				For j = 2 To 6
					counter = count
					For i = 1 To counter
						If sCompare(j) = sSides(1, i) Then
							sSides(1, i) = sCompare(j)
							sSides(2, i) = sSides(2, i) & "," & GetSideLetterFromNumber(j)
						ElseIf i = count Then 
							ReDim Preserve sSides(2, i + 1)
							count = count + 1
							sSides(1, i + 1) = sCompare(j)
							sSides(2, i + 1) = GetSideLetterFromNumber(j)
						End If
					Next 
				Next 
				'get final string and include the surface options to the armor
				mvarPrintOutput = "DR " & VB6.Format(mvarDR) & " "
				For i = 1 To count
					mvarPrintOutput = mvarPrintOutput & " " & sSides(2, i) & ": " & sSides(1, i)
				Next 
				mvarPrintOutput = mvarPrintOutput & mvarQuality & " " & mvarMaterial & " (" & VB6.Format(mvarWeight, p_sFormat) & " lbs., $" & VB6.Format(mvarCost, p_sFormat) & ")."
				
				
			Case ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
				mvarPrintOutput = "PD " & VB6.Format(mvarPD) & ", DR " & VB6.Format(mvarDR) & " " & mvarQuality & " " & mvarMaterial & " (" & VB6.Format(mvarWeight, p_sFormat) & " lbs., $" & VB6.Format(mvarCost, p_sFormat) & ")."
				
		End Select
		
	End Sub
	
	Private Function GetSideLetterFromNumber(ByRef i As Integer) As String
		Select Case i
			Case 1
				GetSideLetterFromNumber = "R"
			Case 2
				GetSideLetterFromNumber = "L"
			Case 3
				GetSideLetterFromNumber = "F"
			Case 4
				GetSideLetterFromNumber = "B"
			Case 5
				GetSideLetterFromNumber = "T"
			Case 6
				GetSideLetterFromNumber = "U"
		End Select
	End Function
	
	Public Function FillMaterial() As String()
		' populate the material combo
		Dim materialarray() As String
		ReDim materialarray(1)
		
		If mvarTL <= 6 Then
			materialarray = mAddKey(materialarray, "wood")
			materialarray = mAddKey(materialarray, "metal")
			materialarray = mAddKey(materialarray, "nonrigid")
		Else ' if its greater than or equal to 7
			materialarray = mAddKey(materialarray, "wood")
			materialarray = mAddKey(materialarray, "metal")
			materialarray = mAddKey(materialarray, "ablative")
			materialarray = mAddKey(materialarray, "fireproof ablative")
			materialarray = mAddKey(materialarray, "nonrigid")
			materialarray = mAddKey(materialarray, "composite")
			materialarray = mAddKey(materialarray, "laminate")
		End If
		
		FillMaterial = VB6.CopyArray(materialarray)
	End Function
	
	Public Function FillQuality(ByRef sMaterial As String) As String()
		Dim MaterialCombo As System.Windows.Forms.ComboBox
		Dim Selected As String ' holds the users selected Armor material
		Dim arrQuality() As Short 'holds the list of suitable quality
		Dim iSelected As Short 'holds converted Selected string
		Dim element As Object 'one element of the arrQuality array
		Dim i As Short 'counter
		Dim count As Short 'another counter
		Dim qualityarray() As String
		Dim TempTL As Short
		
		ReDim qualityarray(1)
		
		Const Cheap As Short = 1
		Const Standard As Short = 2
		Const Expensive As Short = 3
		Const Advanced As Short = 4
		
		'get the type of armor that the user selected
		Selected = sMaterial
		'convert the Selected into an integer
		Select Case Selected
			Case "wood"
				iSelected = 1
			Case "metal"
				iSelected = 2
			Case "ablative"
				iSelected = 3
			Case "fireproof ablative"
				iSelected = 4
			Case "nonrigid"
				iSelected = 5
			Case "composite"
				iSelected = 6
			Case "laminate"
				iSelected = 7
		End Select
		
		count = 1 ' init the counter
		'given the tech level, produce list of quality types
		'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempTL = modPerformance.Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
		'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempTL = modPerformance.Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
		For i = 1 To UBound(ArmorMatrix)
			If ArmorMatrix(i).TL = TempTL Then
				If ArmorMatrix(i).MaterialType = iSelected Then
					If ArmorMatrix(i).WeightMod <> 0 Then
						ReDim Preserve arrQuality(count)
						arrQuality(count) = ArmorMatrix(i).Quality
						count = count + 1
						If count > 5 Then Exit For Else 
					End If
				End If
			End If
		Next 
		
		'fill the Quality combo with the list of available items
		For	Each element In arrQuality
			Select Case element
				Case Cheap
					qualityarray = mAddKey(qualityarray, "cheap")
					'If TempText = "cheap" Then TextFlag = True
				Case Standard
					qualityarray = mAddKey(qualityarray, "standard")
					'If TempText = "standard" Then TextFlag = True
				Case Expensive
					qualityarray = mAddKey(qualityarray, "expensive")
					'If TempText = "expensive" Then TextFlag = True
				Case Advanced
					qualityarray = mAddKey(qualityarray, "advanced")
					'If TempText = "advanced" Then TextFlag = True
			End Select
		Next element
		
		FillQuality = VB6.CopyArray(qualityarray)
		
	End Function
	
	Sub CalcArmorWeightCost()
		Dim TempTL As Short
		' This routine calculates the Cost and Weight of the armor.
		'These contstant values must match those in the module "modArmor" since
		'the armormatrix uses integers and not string names for the material and quality
		Const Cheap As Short = 1
		Const Standard As Short = 2
		Const Expensive As Short = 3
		Const Advanced As Short = 4
		
		Const Wood As Short = 1
		Const Metal As Short = 2
		Const Ablative As Short = 3
		Const FireproofAblative As Short = 4
		Const NonRigid As Short = 5
		Const Composite As Short = 6
		Const Laminate As Short = 7
		
		Dim Area As Single 'holds the surface area
		Dim CostModifier As Single
		Dim WeightModifier As Single
		Dim SelectedMaterial As Short
		Dim SelectedQuality As Short
		Dim i As Short ' counter
		
		' Get the surface area based on the armor datatype being used
		If (mvarDatatype = ArmorWheelGuard) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Area = Veh.Components(mvarParent).SurfaceArea / 2
		ElseIf (mvarDatatype = ArmorGunShield) Or (mvarDatatype = ArmorOpenFrame) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Area = Veh.Components(mvarParent).SurfaceArea / 5
		ElseIf (mvarDatatype = ArmorLocation) Or (mvarDatatype = ArmorComponent) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Area = Veh.Components(mvarParent).SurfaceArea
		ElseIf mvarDatatype = ArmorOverall Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Area = Veh.Stats.StructuralSurfaceArea
		End If
		
		'Determine Material
		If mvarMaterial = "wood" Then
			SelectedMaterial = Wood
		ElseIf mvarMaterial = "metal" Then 
			SelectedMaterial = Metal
		ElseIf mvarMaterial = "ablative" Then 
			SelectedMaterial = Ablative
		ElseIf mvarMaterial = "fireproof ablative" Then 
			SelectedMaterial = FireproofAblative
		ElseIf mvarMaterial = "nonrigid" Then 
			SelectedMaterial = NonRigid
		ElseIf mvarMaterial = "composite" Then 
			SelectedMaterial = Composite
		ElseIf mvarMaterial = "laminate" Then 
			SelectedMaterial = Laminate
		End If
		
		'Determine Quality
		If mvarQuality = "cheap" Then
			SelectedQuality = Cheap
		ElseIf mvarQuality = "standard" Then 
			SelectedQuality = Standard
		ElseIf mvarQuality = "expensive" Then 
			SelectedQuality = Expensive
		ElseIf mvarQuality = "advanced" Then 
			SelectedQuality = Advanced
		End If
		
		' Get the Cost and Weight Modifiers
		'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempTL = modPerformance.Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
		'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempTL = modPerformance.Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
		For i = 1 To UBound(ArmorMatrix)
			If ArmorMatrix(i).TL = TempTL Then
				If ArmorMatrix(i).MaterialType = SelectedMaterial Then
					If ArmorMatrix(i).Quality = SelectedQuality Then
						CostModifier = ArmorMatrix(i).Cost
						WeightModifier = ArmorMatrix(i).WeightMod
						Exit For
					End If
				End If
			End If
		Next 
		
		
		mvarWeight = mvarDR * Area * WeightModifier
		mvarCost = mvarWeight * CostModifier
		
		'get the final weight and cost by adding the cost/weight of the surface features
		CalcSurfaceFeaturesCostandWeight(Area)
	End Sub
	
	Sub CalcByFacingArmorWeightCost()
		' This routine calculates the Cost and Weight of the armor
		Dim Area As Single 'holds the surface area
		Dim CostModifier(6) As Single
		Dim WeightModifier(6) As Single
		Dim iWeight(6) As Single
		Dim iCost(6) As Single
		Dim SelectedMaterial(6) As String
		Dim SelectedQuality(6) As String
		Dim iSelectedQuality(6) As Short
		Dim iSelectedMaterial(6) As Short
		Dim count As Short
		Dim i As Short
		Dim arrMaterial(6) As String
		Dim arrQuality(6) As String
		Dim arrDR(6) As Integer
		Dim TempCost As Single
		Dim TempWeight As Single
		Dim TempTL As Short
		'fill the arrMaterial array and arrQuality
		arrMaterial(0) = mvarMaterial1
		arrMaterial(1) = mvarMaterial2
		arrMaterial(2) = mvarMaterial3
		arrMaterial(3) = mvarMaterial4
		arrMaterial(4) = mvarMaterial5
		arrMaterial(5) = mvarMaterial6
		
		arrQuality(0) = mvarQuality1
		arrQuality(1) = mvarQuality2
		arrQuality(2) = mvarQuality3
		arrQuality(3) = mvarQuality4
		arrQuality(4) = mvarQuality5
		arrQuality(5) = mvarQuality6
		
		'fill the arrDR aray
		arrDR(0) = mvarDR1
		arrDR(1) = mvarDR2
		arrDR(2) = mvarDR3
		arrDR(3) = mvarDR4
		arrDR(4) = mvarDR5
		arrDR(5) = mvarDR6
		
		' re-init variables
		TempCost = 0
		TempWeight = 0
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Area = Veh.Components(mvarParent).SurfaceArea
		' There are just two paths, one for Complex and one for Basic
		Select Case mvarDatatype
			Case ArmorComplexFacing
				For count = 0 To 5
					' Get the quality and material of the armor
					SelectedMaterial(count) = arrMaterial(count)
					SelectedQuality(count) = arrQuality(count)
					'convert the Selected into an integer
					Select Case SelectedMaterial(count)
						Case "wood"
							iSelectedMaterial(count) = 1
						Case "metal"
							iSelectedMaterial(count) = 2
						Case "ablative"
							iSelectedMaterial(count) = 3
						Case "fireproof ablative"
							iSelectedMaterial(count) = 4
						Case "nonrigid"
							iSelectedMaterial(count) = 5
						Case "composite"
							iSelectedMaterial(count) = 6
						Case "laminate"
							iSelectedMaterial(count) = 7
					End Select
					Select Case SelectedQuality(count)
						Case "cheap"
							iSelectedQuality(count) = 1
						Case "standard"
							iSelectedQuality(count) = 2
						Case "expensive"
							iSelectedQuality(count) = 3
						Case "advanced"
							iSelectedQuality(count) = 4
					End Select
					' Get the Cost and Weight Modifiers
					'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempTL = modPerformance.Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
					'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempTL = modPerformance.Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
					For i = 1 To UBound(ArmorMatrix)
						If ArmorMatrix(i).TL = TempTL Then
							If ArmorMatrix(i).MaterialType = iSelectedMaterial(count) Then
								If ArmorMatrix(i).Quality = iSelectedQuality(count) Then
									CostModifier(count) = ArmorMatrix(i).Cost
									WeightModifier(count) = ArmorMatrix(i).WeightMod
									Exit For
								End If
							End If
						End If
					Next 
					'get the average DR
					CalcAverageDr()
					' Get the Cost and weight of each face
					If mvarParent = "1_" Then
						iWeight(count) = Val(CStr(arrDR(count))) * (Area / 6) * WeightModifier(count)
					Else
						iWeight(count) = Val(CStr(arrDR(count))) * (Area / 5) * WeightModifier(count)
					End If
					iCost(count) = iWeight(count) * CostModifier(count)
					TempWeight = TempWeight + iWeight(count)
					TempCost = TempCost + iCost(count)
				Next 
				
			Case ArmorBasicFacing
				
				'convert the Selected into an integer
				Select Case mvarMaterial
					Case "wood"
						iSelectedMaterial(0) = 1
					Case "metal"
						iSelectedMaterial(0) = 2
					Case "ablative"
						iSelectedMaterial(0) = 3
					Case "fireproof ablative"
						iSelectedMaterial(0) = 4
					Case "nonrigid"
						iSelectedMaterial(0) = 5
					Case "composite"
						iSelectedMaterial(0) = 6
					Case "laminate"
						iSelectedMaterial(0) = 7
				End Select
				Select Case mvarQuality
					Case "cheap"
						iSelectedQuality(0) = 1
					Case "standard"
						iSelectedQuality(0) = 2
					Case "expensive"
						iSelectedQuality(0) = 3
					Case "advanced"
						iSelectedQuality(0) = 4
				End Select
				' Get the Cost and Weight Modifiers
				'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TempTL = modPerformance.Maximum(4, mvarTL) 'our matrix assumes 4 for TL4-
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TempTL = modPerformance.Minimum(13, TempTL) 'our matrix only goes to TL13 since values over 13 use the same stats
				For i = 1 To UBound(ArmorMatrix)
					If ArmorMatrix(i).TL = TempTL Then
						If ArmorMatrix(i).MaterialType = iSelectedMaterial(0) Then
							If ArmorMatrix(i).Quality = iSelectedQuality(0) Then
								CostModifier(0) = ArmorMatrix(i).Cost
								WeightModifier(0) = ArmorMatrix(i).WeightMod
							End If
						End If
					End If
				Next 
				'call routine to calc averagedr
				CalcAverageDr()
				' Get the Final Cost and Final Weight
				TempWeight = AverageDR * Area * WeightModifier(0)
				TempCost = TempWeight * CostModifier(0)
		End Select
		
		
		'save these cost and weight results to the armor class
		mvarCost = TempCost
		mvarWeight = TempWeight
		
		'get the final cost which includes the cost of the surface features
		CalcSurfaceFeaturesCostandWeight(Area)
	End Sub
	
	
	Function CalcEffectiveDR(ByRef DR As Integer, ByRef Slope As String) As Object
		Dim Modifier As Single
		
		If Slope = "none" Then
			Modifier = 1
		ElseIf Slope = "30 degrees" Then 
			Modifier = 1.5
		ElseIf Slope = "60 degrees" Then 
			Modifier = 2
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcEffectiveDR. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalcEffectiveDR = DR * Modifier
	End Function
	Function CalcPD(ByRef DR As Integer, ByRef Slope As String, ByRef Material As String) As Integer
		Dim PD As Short
		
		If DR = 0 Then
			PD = 0
		ElseIf DR = 1 Then 
			PD = 1
		ElseIf DR <= 4 Then 
			PD = 2
		ElseIf DR <= 15 Then 
			PD = 3
		ElseIf DR >= 16 Then 
			PD = 4
		End If
		
		'check for max values for Wood and Nonrigid armor
		If Material = "nonrigid" Then
			If PD > 2 Then PD = 2
		ElseIf Material = "wood" Then 
			If PD > 3 Then PD = 3
		End If
		
		'add bonus for slope
		'Note: the bonus's are placed below the checks for  max values for wood and nonrigid.
		'if users request, it can be moved above it
		If Slope = "none" Then
		ElseIf Slope = "30 degrees" Then 
			PD = PD + 1
		ElseIf Slope = "60 degrees" Then 
			PD = PD + 2
		End If
		
		CalcPD = PD
	End Function
	
	
	Function GetLowestDR() As Integer
		'//this function only gets called during Aerial performance calculations.
		'//its job is to return the DR of the armor. In the case of seperate DR's
		'//for each face, then  it will return the lowest one.
		'//only DR from metal, composite or laminate armor counts.. all other types
		'//return 0.
		On Error Resume Next
		Dim lngRetval As Integer
		
		Select Case mvarDatatype
			Case ArmorComplexFacing
				Select Case mvarMaterial1
					Case "metal", "composite", "laminate"
						lngRetval = mvarDR1
						Select Case mvarMaterial2
							Case "metal", "composite", "laminate"
								'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								lngRetval = modPerformance.Minimum(mvarDR1, mvarDR2)
								Select Case mvarMaterial3
									Case "metal", "composite", "laminate"
										'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										lngRetval = modPerformance.Minimum(lngRetval, mvarDR3)
										Select Case mvarMaterial4
											Case "metal", "composite", "laminate"
												'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												lngRetval = modPerformance.Minimum(lngRetval, mvarDR4)
												Select Case mvarMaterial5
													Case "metal", "composite", "laminate"
														'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														lngRetval = modPerformance.Minimum(lngRetval, mvarDR5)
														Select Case mvarMaterial6
															Case "metal", "composite", "laminate"
																'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
																lngRetval = modPerformance.Minimum(lngRetval, mvarDR6)
															Case Else
																lngRetval = 0
														End Select
													Case Else
														lngRetval = 0
												End Select
											Case Else
												lngRetval = 0
										End Select
									Case Else
										lngRetval = 0
								End Select
							Case Else
								lngRetval = 0
						End Select
					Case Else
						lngRetval = 0
				End Select
				
			Case ArmorBasicFacing
				Select Case mvarMaterial
					Case "metal", "composite", "laminate"
						'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRetval = modPerformance.Minimum(mvarDR1, mvarDR2)
						'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRetval = modPerformance.Minimum(lngRetval, mvarDR3)
						'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRetval = modPerformance.Minimum(lngRetval, mvarDR4)
						'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRetval = modPerformance.Minimum(lngRetval, mvarDR5)
						'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRetval = modPerformance.Minimum(lngRetval, mvarDR6)
						
					Case Else
						lngRetval = 0
				End Select
				
			Case ArmorLocation, ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
				Select Case mvarMaterial
					Case "metal", "composite", "laminate"
						lngRetval = mvarDR
					Case Else
						lngRetval = 0
				End Select
				
		End Select
		
		
		GetLowestDR = lngRetval
	End Function
	
	Function GetLowestCrushDepthDR() As Integer
		
		On Error Resume Next
		Dim lngRetval As Integer
		
		Select Case mvarDatatype
			Case ArmorComplexFacing
				
				lngRetval = mvarDR1
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(mvarDR1, mvarDR2)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR3)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR4)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR5)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR6)
				
			Case ArmorBasicFacing
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(mvarDR1, mvarDR2)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR3)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR4)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR5)
				'UPGRADE_WARNING: Couldn't resolve default property of object Minimum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngRetval = modPerformance.Minimum(lngRetval, mvarDR6)
				
				
			Case ArmorLocation, ArmorOpenFrame, ArmorGunShield, ArmorComponent, ArmorOverall, ArmorWheelGuard
				lngRetval = mvarDR
				
		End Select
		
		
		GetLowestCrushDepthDR = lngRetval
	End Function
	
	Private Sub CalcAverageDr()
		Dim i As Short
		Dim TempAverage As Integer
		Dim Divisor As Short
		Dim arrDR(5) As Integer
		
		arrDR(0) = Val(CStr(mvarDR1))
		arrDR(1) = Val(CStr(mvarDR2))
		arrDR(2) = Val(CStr(mvarDR3))
		arrDR(3) = Val(CStr(mvarDR4))
		arrDR(4) = Val(CStr(mvarDR5))
		arrDR(5) = Val(CStr(mvarDR6))
		
		TempAverage = 0
		
		For i = 0 To 5
			TempAverage = TempAverage + arrDR(i)
		Next 
		
		'find the Average DR (this is done only for Basic Armor by facing
		If mvarParent = "1_" Then
			Divisor = 6
		Else
			Divisor = 5
		End If
		
		mvarAverageDR = System.Math.Round(TempAverage / Divisor, 2)
		
	End Sub
	
	
	Sub CalcSurfaceFeaturesCostandWeight(ByRef SurfaceArea As Single)
		Const RadWeight As Short = 2
		Const RadCost As Short = 20
		Const ReflectCost As Short = 30
		Const RetroCost As Short = 150
		Const ThermCost As Short = 250
		Const ThermWeight As Double = 0.25
		Const RAPCost As Short = 20
		Const RAPWeight As Short = 8
		Const ElectCost As Short = 10
		Const ElectWeight As Double = 0.2
		Dim TempWeight As Single
		Dim TempCost As Single
		
		If SurfaceArea = 0 Then Exit Sub
		
		If mvarRadiation Then
			TempWeight = RadWeight * SurfaceArea
			TempCost = RadCost * SurfaceArea
		End If
		
		If mvarCoating = "reflective" Then
			TempCost = TempCost + (ReflectCost * SurfaceArea)
		ElseIf mvarCoating = "retro-reflective" Then 
			TempCost = TempCost + (RetroCost * SurfaceArea)
		End If
		
		If mvarThermal Then
			TempWeight = TempWeight + (ThermWeight * SurfaceArea)
			TempCost = TempCost + (ThermCost * SurfaceArea)
		End If
		
		If mvarRAP Then
			TempWeight = TempWeight + (RAPWeight * SurfaceArea)
			TempCost = TempCost + (RAPCost * SurfaceArea)
		End If
		
		If mvarElectrified Then
			TempWeight = TempWeight + (ElectWeight * SurfaceArea)
			TempCost = TempCost + (ElectCost * SurfaceArea)
		End If
		
		mvarCost = System.Math.Round(mvarCost + TempCost, 2)
		mvarWeight = System.Math.Round(mvarWeight + TempWeight, 2)
		
	End Sub
	
	
	Public Sub QueryParent()
		' if the object has a parent, query it and check to see if
		' more stats/property updates are needed for other objects in the collection
		If mvarParent <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Veh.Components(Parent).StatsUpdate()
		End If
		
	End Sub
	
	Public Sub QueryChild()
		' if the object has children, query it and check to see if
		' more stats/property updates are needed for other objects in the collection
		' (see the StatusUpdate property for help on checking for childeren.  Can i use that one in place of this?)
		
	End Sub
End Class