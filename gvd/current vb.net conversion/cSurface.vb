Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cSurface_NET.cSurface")> Public Class cSurface
	Implements _cIPersist
	Implements _cINode
	Implements _cIDisplay
	
	
	' -------cINode interface variables
	Private m_lngMaxChildren As Integer
	Private m_lngChildCount As Integer
	Private m_oChildren() As cINode
	Private m_lngAttributes As Integer
	Private m_hParent As Integer
	Private m_hMe As Integer
	Private m_sName As String
	Private m_sDescription As String
	Private m_sImage As String
	
	' -------cIDisplay interface variables
	Private m_lngPropCount As Integer
	Private m_lngCurrentPropItem As Integer
	'UPGRADE_ISSUE: cPropertyItem object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Private m_oProperties() As cPropertyItem
	
	
	
	Private m_bLevitationCoating As Boolean
	Private m_bMagicLevitation As Boolean
	Private m_bAntigravityCoating As Boolean
	Private m_bSuperScienceCoating As Boolean
	
	Private m_sngMagicLevitationEnergyCost As Single
	Private m_sngAntigravityCoatingCost As Single
	Private m_sngSuperScienceCoatingCost As Single
	Private m_sngMagicLevitationEnergyCostPerPound As Single
	Private m_sngAntigravityCoatingCostPerSquareFoot As Single
	Private m_sngSuperScienceCoatingCostPerSquareFoot As Single
	Private m_byteAntigravityCoatingSurfaceAreaUseage As Byte
	Private m_byteSuperScienceCoatingSurfaceAreaUseage As Byte
	
	
	
	' totals
	Private m_dblWeight As Double
	Private m_dblCost As Double
	
	Public Property Cost() As Double
		Get
			Cost = m_dblCost
		End Get
		Set(ByVal Value As Double)
			m_dblCost = Value
		End Set
	End Property
	Public Property Weight() As Double
		Get
			Weight = m_dblWeight
		End Get
		Set(ByVal Value As Double)
			m_dblWeight = Value
		End Set
	End Property
	
	
	
	Public Property MagicLevitationEnergyCostPerPound() As Single
		Get
			Dim mvarMagicLevitationEnergyCostPerPound As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarMagicLevitationEnergyCostPerPound. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MagicLevitationEnergyCostPerPound = mvarMagicLevitationEnergyCostPerPound
		End Get
		Set(ByVal Value As Single)
			Dim mvarMagicLevitationEnergyCostPerPound As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarMagicLevitationEnergyCostPerPound. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarMagicLevitationEnergyCostPerPound = Value
		End Set
	End Property
	Public Property MagicLevitationEnergyCost() As Single
		Get
			Dim mvarMagicLevitationEnergyCost As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarMagicLevitationEnergyCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MagicLevitationEnergyCost = mvarMagicLevitationEnergyCost
		End Get
		Set(ByVal Value As Single)
			Dim mvarMagicLevitationEnergyCost As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarMagicLevitationEnergyCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarMagicLevitationEnergyCost = Value
		End Set
	End Property
	Public Property AntigravityCoatingCostPerSquareFoot() As Single
		Get
			Dim mvarAntigravityCoatingCostPerSquareFoot As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarAntigravityCoatingCostPerSquareFoot. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AntigravityCoatingCostPerSquareFoot = mvarAntigravityCoatingCostPerSquareFoot
		End Get
		Set(ByVal Value As Single)
			Dim mvarAntigravityCoatingCostPerSquareFoot As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarAntigravityCoatingCostPerSquareFoot. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarAntigravityCoatingCostPerSquareFoot = Value
		End Set
	End Property
	Public Property AntigravityCoatingCost() As Single
		Get
			Dim mvarAntigravityCoatingCost As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarAntigravityCoatingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AntigravityCoatingCost = mvarAntigravityCoatingCost
		End Get
		Set(ByVal Value As Single)
			Dim mvarAntigravityCoatingCost As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarAntigravityCoatingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarAntigravityCoatingCost = Value
		End Set
	End Property
	Public Property SuperScienceCoatingCostPerSquareFoot() As Single
		Get
			Dim mvarSuperScienceCoatingCostPerSquareFoot As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSuperScienceCoatingCostPerSquareFoot. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SuperScienceCoatingCostPerSquareFoot = mvarSuperScienceCoatingCostPerSquareFoot
		End Get
		Set(ByVal Value As Single)
			Dim mvarSuperScienceCoatingCostPerSquareFoot As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSuperScienceCoatingCostPerSquareFoot. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarSuperScienceCoatingCostPerSquareFoot = Value
		End Set
	End Property
	Public Property SuperScienceCoatingCost() As Single
		Get
			Dim mvarSuperScienceCoatingCost As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSuperScienceCoatingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SuperScienceCoatingCost = mvarSuperScienceCoatingCost
		End Get
		Set(ByVal Value As Single)
			Dim mvarSuperScienceCoatingCost As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSuperScienceCoatingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarSuperScienceCoatingCost = Value
		End Set
	End Property
	
	
	Public Property bMagicLevitation() As Boolean
		Get
			Dim mvarbMagicLevitation As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarbMagicLevitation. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			bMagicLevitation = mvarbMagicLevitation
		End Get
		Set(ByVal Value As Boolean)
			Dim mvarbMagicLevitation As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarbMagicLevitation. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarbMagicLevitation = Value
		End Set
	End Property
	Public Property bAntigravityCoating() As Boolean
		Get
			Dim mvarbAntigravityCoating As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarbAntigravityCoating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			bAntigravityCoating = mvarbAntigravityCoating
		End Get
		Set(ByVal Value As Boolean)
			Dim mvarbAntigravityCoating As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarbAntigravityCoating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarbAntigravityCoating = Value
		End Set
	End Property
	Public Property bSuperScienceCoating() As Boolean
		Get
			Dim mvarbSuperScienceCoating As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarbSuperScienceCoating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			bSuperScienceCoating = mvarbSuperScienceCoating
		End Get
		Set(ByVal Value As Boolean)
			Dim mvarbSuperScienceCoating As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarbSuperScienceCoating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarbSuperScienceCoating = Value
		End Set
	End Property
	Public Property AntigravityCoatingSurfaceAreaUseage() As String
		Get
			Dim mvarAntigravityCoatingSurfaceAreaUseage As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarAntigravityCoatingSurfaceAreaUseage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AntigravityCoatingSurfaceAreaUseage = mvarAntigravityCoatingSurfaceAreaUseage
		End Get
		Set(ByVal Value As String)
			Dim mvarAntigravityCoatingSurfaceAreaUseage As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarAntigravityCoatingSurfaceAreaUseage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarAntigravityCoatingSurfaceAreaUseage = Value
		End Set
	End Property
	Public Property SuperScienceCoatingSurfaceAreaUseage() As String
		Get
			Dim mvarSuperScienceCoatingSurfaceAreaUseage As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSuperScienceCoatingSurfaceAreaUseage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SuperScienceCoatingSurfaceAreaUseage = mvarSuperScienceCoatingSurfaceAreaUseage
		End Get
		Set(ByVal Value As String)
			Dim mvarSuperScienceCoatingSurfaceAreaUseage As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSuperScienceCoatingSurfaceAreaUseage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarSuperScienceCoatingSurfaceAreaUseage = Value
		End Set
	End Property
	Private ReadOnly Property cINode_childCount() As Integer Implements _cINode.childCount
		Get
			cINode_childCount = m_lngChildCount
		End Get
	End Property
	Private ReadOnly Property cINode_ClassName() As String Implements _cINode.ClassName
		Get
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			cINode_ClassName = TypeName(Me)
		End Get
	End Property
	Private ReadOnly Property cINode_Attributes() As Integer Implements _cINode.Attributes
		Get
			cINode_Attributes = m_lngAttributes
		End Get
	End Property
	Private Property cINode_Handle() As Integer Implements _cINode.Handle
		Get
			cINode_Handle = m_hMe
		End Get
		Set(ByVal Value As Integer)
			m_hMe = Value
		End Set
	End Property
	Private Property cINode_Parent() As Integer Implements _cINode.Parent
		Get
			cINode_Parent = m_hParent
		End Get
		Set(ByVal Value As Integer)
			m_hParent = Value
		End Set
	End Property
	Private Property cINode_Name() As String Implements _cINode.Name
		Get
			cINode_Name = m_sName
		End Get
		Set(ByVal Value As String)
			m_sName = Value
		End Set
	End Property
	Private Property cINode_Description() As String Implements _cINode.Description
		Get
			cINode_Description = m_sDescription
		End Get
		Set(ByVal Value As String)
			m_sDescription = Value
		End Set
	End Property
	Private Property cINode_Image() As String Implements _cINode.Image
		Get
			cINode_Image = m_sImage
		End Get
		Set(ByVal Value As String)
			m_sImage = Value
		End Set
	End Property
	
	'//cIPersist Interface
	Private ReadOnly Property cIPersist_Classname() As String
		Get
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			cIPersist_Classname = TypeName(Me) 'todo: did i ever test that typename me returns the base class and not the class of an interface in use?
		End Get
	End Property
	Private ReadOnly Property cIPersist_GUID() As String
		Get
		End Get
	End Property
	
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		m_bLevitationCoating = False
		m_bMagicLevitation = False
		m_bAntigravityCoating = False
		m_bSuperScienceCoating = False
		m_sngMagicLevitationEnergyCostPerPound = 700
		m_sngAntigravityCoatingCostPerSquareFoot = 10
		m_sngSuperScienceCoatingCostPerSquareFoot = 100
		m_byteAntigravityCoatingSurfaceAreaUseage = 1 '"Body"
		m_byteSuperScienceCoatingSurfaceAreaUseage = 1 '"Body"
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	Sub Update()
		Dim mvarCost As Object
		Dim mvarCamouflageCost As Object
		Dim mvarInfraredCost As Object
		Dim mvarWeight As Object
		Dim mvarInfraredWeight As Object
		Dim mvarSealed As Object
		Dim mvarSealedCost As Object
		Dim mvarWaterProof As Object
		Dim mvarWaterProofCost As Object
		Dim mvarbSuperScienceCoating As Object
		Dim mvarSuperScienceCoatingSurfaceAreaUseage As Object
		Dim mvarSuperScienceCoatingCost As Object
		Dim mvarSuperScienceCoatingCostPerSquareFoot As Object
		Dim mvarbAntigravityCoating As Object
		Dim mvarAntigravityCoatingSurfaceAreaUseage As Object
		Dim mvarAntigravityCoatingCost As Object
		Dim mvarAntigravityCoatingCostPerSquareFoot As Object
		Dim mvarbMagicLevitation As Object
		Dim mvarMagicLevitationEnergyCost As Object
		Dim mvarMagicLevitationEnergyCostPerPound As Object
		Dim mvarPsiShielding As Object
		Dim mvarPsiShieldingWeight As Object
		Dim mvarPsiShieldingCost As Object
		Dim mvarLiquidCrystal As Object
		Dim mvarLiquidCrystalWeight As Object
		Dim mvarLiquidCrystalCost As Object
		Dim mvarChameleon As Object
		Dim mvarChameleonWeight As Object
		Dim mvarChameleonCost As Object
		Dim mvarStealth As Object
		Dim mvarStealthWeight As Object
		Dim mvarStealthCost As Object
		Dim m_byteSoundBaffling As Object
		Dim mvarSoundWeight As Object
		Dim mvarSoundCost As Object
		Dim m_byteEmissionCloaking As Object
		Dim mvarEmissionWeight As Object
		Dim mvarEmissionCost As Object
		Dim m_byteInfraredCloaking As Object
		Dim m_sngInfraredWeight As Object
		Dim m_sngInfraredCost As Object
		Dim m_bCamouflage As Object
		Dim m_sngCamouflageCost As Object
		
		Const BasicInfrared As Short = 2
		Const RadicalInfrared As Short = 3
		Const BasicEmission As Short = 4
		Const RadicalEmission As Short = 5
		Const BasicSound As Short = 6
		Const RadicalSound As Short = 7
		Const BasicStealth As Short = 8
		Const RadicalStealth As Short = 9
		Const BasicChameleon As Short = 10
		Const InstantChameleon As Short = 11
		Const IntruderChameleon As Short = 12
		Const LiquidCrystal As Short = 13
		Const PsiShielding As Short = 14
		
		'----------
		Dim dblBodyArea As Double
		Dim dblTotalArea As Double
		Dim BodyHits As Single
		Dim dblTrimmedArea As Double
		Dim dblIgnoredArea As Double
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With Veh.Components(BODY_KEY)
			' Get the surface area of the body
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dblBodyArea = .SurfaceArea
			' Get the surface are of the entire vehicle
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dblTotalArea = Veh.Stats.totalSurfaceArea
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BodyHits = .HitPoints
		End With
		' Get the area minus skids, masts and gas bags (for sound baffling rules on page 92)
		' TODO: dblIgnoredArea is 0? wtf i think i forgot to get the skis,mast and bags area.
		' And isnt that the same as structural surface area anyway?  I forget, look it up.
		dblTrimmedArea = dblTotalArea - dblIgnoredArea
		' Get Cost for camouflage (note there is no weight)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_bCamouflage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If m_bCamouflage = False Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_sngCamouflageCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sngCamouflageCost = 0
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object m_sngCamouflageCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sngCamouflageCost = 0.1 * dblTotalArea
		End If
		' Get Cost and Weight for Infrared Cloaking
		If m_byteInfraredCloaking = modConstants.EMISSION_CLOAKING.None Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_sngInfraredCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sngInfraredCost = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object m_sngInfraredWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sngInfraredWeight = 0
		ElseIf m_byteInfraredCloaking = modConstants.EMISSION_CLOAKING.BASIC Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object m_sngInfraredCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sngInfraredCost = GetSurfaceCost(BasicInfrared) * dblTotalArea
			'UPGRADE_WARNING: Couldn't resolve default property of object m_sngInfraredWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sngInfraredWeight = GetSurfaceWeight(BasicInfrared) * dblTotalArea
		Else
			System.Diagnostics.Debug.Assert(m_byteInfraredCloaking = modConstants.EMISSION_CLOAKING.RADICAL, "")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_sngInfraredCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sngInfraredCost = GetSurfaceCost(RadicalInfrared) * dblTotalArea
			'UPGRADE_WARNING: Couldn't resolve default property of object m_sngInfraredWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sngInfraredWeight = GetSurfaceWeight(RadicalInfrared) * dblTotalArea
		End If
		' Get Cost and Weight for Emission Cloaking
		If m_byteEmissionCloaking = modConstants.EMISSION_CLOAKING.None Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEmissionCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEmissionCost = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEmissionWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEmissionWeight = 0
		ElseIf m_byteEmissionCloaking = modConstants.EMISSION_CLOAKING.BASIC Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEmissionCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEmissionCost = GetSurfaceCost(BasicEmission) * dblTotalArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEmissionWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEmissionWeight = GetSurfaceWeight(BasicEmission) * dblTotalArea
		Else
			System.Diagnostics.Debug.Assert(m_byteEmissionCloaking = modConstants.EMISSION_CLOAKING.RADICAL, "")
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEmissionCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEmissionCost = GetSurfaceCost(RadicalEmission) * dblTotalArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEmissionWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarEmissionWeight = GetSurfaceWeight(RadicalEmission) * dblTotalArea
		End If
		' Get cost and weight for Sound Baffling
		If m_byteSoundBaffling = modConstants.EMISSION_CLOAKING.None Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSoundCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarSoundCost = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSoundWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarSoundWeight = 0
		ElseIf m_byteSoundBaffling = modConstants.EMISSION_CLOAKING.BASIC Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSoundCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarSoundCost = GetSurfaceCost(BasicSound) * dblTrimmedArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSoundWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarSoundWeight = GetSurfaceWeight(BasicSound) * dblTrimmedArea
		Else
			System.Diagnostics.Debug.Assert(m_byteSoundBaffling = modConstants.EMISSION_CLOAKING.RADICAL, "")
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSoundCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarSoundCost = GetSurfaceCost(RadicalSound) * dblTrimmedArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSoundWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarSoundWeight = GetSurfaceWeight(RadicalSound) * dblTrimmedArea
		End If
		' Get cost and weight for Stealth
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarStealth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarStealth = "none" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarStealthCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarStealthCost = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarStealthWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarStealthWeight = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarStealth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarStealth = "basic" Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarStealthCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarStealthCost = GetSurfaceCost(BasicStealth) * dblTotalArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarStealthWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarStealthWeight = GetSurfaceWeight(BasicStealth) * dblTotalArea
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarStealthCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarStealthCost = GetSurfaceCost(RadicalStealth) * dblTotalArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarStealthWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarStealthWeight = GetSurfaceWeight(RadicalStealth) * dblTotalArea
		End If
		'Get cost and weight for Chameleon system
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarChameleon. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarChameleon = "none" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarChameleonCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarChameleonCost = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarChameleonWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarChameleonWeight = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarChameleon. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarChameleon = "basic" Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarChameleonCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarChameleonCost = GetSurfaceCost(BasicChameleon) * dblTotalArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarChameleonWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarChameleonWeight = GetSurfaceWeight(BasicChameleon) * dblTotalArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarChameleon. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf mvarChameleon = "instant" Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarChameleonCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarChameleonCost = GetSurfaceCost(InstantChameleon) * dblTotalArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarChameleonWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarChameleonWeight = GetSurfaceWeight(InstantChameleon) * dblTotalArea
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarChameleonCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarChameleonCost = GetSurfaceCost(IntruderChameleon) * dblTotalArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarChameleonWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarChameleonWeight = GetSurfaceWeight(IntruderChameleon) * dblTotalArea
		End If
		'Get cost and weight for LiquidCrystal skin
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarLiquidCrystal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarLiquidCrystal = False Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarLiquidCrystalCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarLiquidCrystalCost = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarLiquidCrystalWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarLiquidCrystalWeight = 0
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarLiquidCrystalCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarLiquidCrystalCost = GetSurfaceCost(LiquidCrystal) * dblTotalArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarLiquidCrystalWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarLiquidCrystalWeight = GetSurfaceWeight(LiquidCrystal) * dblTotalArea
		End If
		' Get cost and weight for PsiShielding
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPsiShielding. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarPsiShielding = False Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPsiShieldingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPsiShieldingCost = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPsiShieldingWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPsiShieldingWeight = 0
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPsiShieldingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPsiShieldingCost = GetSurfaceCost(PsiShielding) * dblTotalArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPsiShieldingWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPsiShieldingWeight = GetSurfaceWeight(PsiShielding) * dblTotalArea
		End If
		' Get cost for Levitation
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarbMagicLevitation. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarbMagicLevitation = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarMagicLevitationEnergyCostPerPound. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarMagicLevitationEnergyCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarMagicLevitationEnergyCost = mvarMagicLevitationEnergyCostPerPound * (dblTotalArea / 250)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarMagicLevitationEnergyCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarMagicLevitationEnergyCost = 0
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarbAntigravityCoating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarbAntigravityCoating = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarAntigravityCoatingSurfaceAreaUseage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If mvarAntigravityCoatingSurfaceAreaUseage = "Vehicle" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarAntigravityCoatingCostPerSquareFoot. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarAntigravityCoatingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarAntigravityCoatingCost = mvarAntigravityCoatingCostPerSquareFoot * dblTotalArea
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarAntigravityCoatingCostPerSquareFoot. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarAntigravityCoatingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarAntigravityCoatingCost = mvarAntigravityCoatingCostPerSquareFoot * dblBodyArea
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarAntigravityCoatingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarAntigravityCoatingCost = 0
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarbSuperScienceCoating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarbSuperScienceCoating = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSuperScienceCoatingSurfaceAreaUseage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If mvarSuperScienceCoatingSurfaceAreaUseage = "Vehicle" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSuperScienceCoatingCostPerSquareFoot. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSuperScienceCoatingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarSuperScienceCoatingCost = mvarSuperScienceCoatingCostPerSquareFoot * dblTotalArea
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSuperScienceCoatingCostPerSquareFoot. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSuperScienceCoatingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarSuperScienceCoatingCost = mvarSuperScienceCoatingCostPerSquareFoot * dblBodyArea
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSuperScienceCoatingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarSuperScienceCoatingCost = 0
		End If
		
		
		'Get Cost for waterproofing
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarWaterProof. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarWaterProof Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarWaterProofCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarWaterProofCost = 2 * Veh.Stats.StructuralSurfaceArea
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarWaterProofCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarWaterProofCost = 0
		End If
		'Get Cost for Sealed vehicle
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarSealed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarSealed Then
			
			If gVehicleTL <= 7 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSealedCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarSealedCost = 40 * Veh.Stats.StructuralSurfaceArea
			ElseIf gVehicleTL = 8 Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSealedCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarSealedCost = 20 * Veh.Stats.StructuralSurfaceArea
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSealedCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarSealedCost = 10 * Veh.Stats.StructuralSurfaceArea
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarSealedCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarSealedCost = 0
		End If
		
		' total surface feature weights
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarInfraredWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarEmissionWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarSoundWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarStealthWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarChameleonWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarLiquidCrystalWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPsiShieldingWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarWeight = mvarPsiShieldingWeight + mvarLiquidCrystalWeight + mvarChameleonWeight + mvarStealthWeight + mvarSoundWeight + mvarEmissionWeight + mvarInfraredWeight
		
		' total surface feature costs
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarCamouflageCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarEmissionCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarInfraredCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarSoundCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarStealthCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarChameleonCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarLiquidCrystalCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPsiShieldingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarWaterProofCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarSealedCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarSuperScienceCoatingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarAntigravityCoatingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarCost = mvarAntigravityCoatingCost + mvarSuperScienceCoatingCost + mvarSealedCost + mvarWaterProofCost + mvarPsiShieldingCost + mvarLiquidCrystalCost + mvarChameleonCost + mvarStealthCost + mvarSoundCost + mvarInfraredCost + mvarEmissionCost + mvarCamouflageCost
		
		
	End Sub
	
	Function GetSurfaceCost(ByRef FeatureID As Short) As Single
		
		' This routine calculates the Cost of a surface Feature
		' IMPORTANT: This routine is optimized to only check valid techlevels!!
		' If the user has somehow enabled a Surface Feature that is not allowed
		' at the vehicles tech level, this routine will return 0!!!!  I must make
		' sure to gray out features that cant be selected in the dialog
		
		Dim CostModifier As Single
		Dim i As Short ' counter
		Dim TempModifier As Single
		Dim TempTech As Short
		
		'On Error GoTo TechLevelError
		' init the two temporary variables
		TempModifier = 0
		TempTech = 0
		' Get the Cost and Weight Modifiers
		For i = 1 To UBound(SurfaceMatrix)
			If SurfaceMatrix(i).FeatureType = FeatureID Then
				If SurfaceMatrix(i).TL = gVehicleTL Then
					CostModifier = SurfaceMatrix(i).CostMod
					GetSurfaceCost = CostModifier
					Exit Function
				ElseIf SurfaceMatrix(i).TL < gVehicleTL Then 
					If SurfaceMatrix(i).TL > TempTech Then
						CostModifier = SurfaceMatrix(i).CostMod
						TempTech = SurfaceMatrix(i).TL
					End If
				End If
			End If
		Next 
		GetSurfaceCost = CostModifier
		'TechLevelError:
		'MsgBox "Error In Function GetSurfaceCost:Unsupported TechLevel with feature ID # " & FeatureID
	End Function
	
	Function GetSurfaceWeight(ByRef FeatureID As Short) As Single
		' This routine calculates the Weight of a surface Feature
		' IMPORTANT: This routine is optimized to only check valid techlevels!!
		' If the user has somehow enabled a Surface Feature that is not allowed
		' at the vehicles tech level, this routine will return 0!!!!
		
		'todo: note the inconsistancies here with these functions returning singles into doubles and what not.
		' I need to make sure they are the same.
		Dim WeightModifier As Single
		Dim i As Short ' counter
		Dim TempTech As Short
		Dim TempModifier As Single
		
		'On Error GoTo TechLevelError
		' init the two temporary variables
		TempModifier = 0
		TempTech = 0
		' Get the Weight Modifiers
		For i = 1 To UBound(SurfaceMatrix)
			If SurfaceMatrix(i).FeatureType = FeatureID Then
				If SurfaceMatrix(i).TL = gVehicleTL Then
					WeightModifier = SurfaceMatrix(i).WeightMod
					GetSurfaceWeight = WeightModifier
					Exit Function
				ElseIf SurfaceMatrix(i).TL < gVehicleTL Then 
					If SurfaceMatrix(i).TL > TempTech Then
						WeightModifier = SurfaceMatrix(i).WeightMod
						TempTech = SurfaceMatrix(i).TL
					End If
				End If
			End If
		Next 
		GetSurfaceWeight = WeightModifier
		'TechLevelError:
		'MsgBox "Error In Function GetSurfaceWeight:Unsupported TechLevel with feature ID # " & FeatureID
	End Function
	
	
	
	'///////////////////////////////////////////////////
	'//cIDisplay Implemented Properties and Functions
	Private Function cIDisplay_getFirstPropertyItem() As cPropertyItem Implements _cIDisplay.getFirstPropertyItem
		On Error GoTo err_Renamed
		If Not m_oProperties(0) Is Nothing Then
			cIDisplay_getFirstPropertyItem = m_oProperties(0)
			m_lngCurrentPropItem = 0
		End If
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print(TypeName(Me) & ":cIDisplay_getFirstPropertyItem() -- no properties in m_oProperties() array for " & TypeName(Me))
	End Function
	
	Private Function cIDisplay_getNextPropertyItem() As cPropertyItem Implements _cIDisplay.getNextPropertyItem
		m_lngCurrentPropItem = m_lngCurrentPropItem + 1
		If m_lngCurrentPropItem <= m_lngPropCount - 1 Then
			If Not m_oProperties(m_lngCurrentPropItem) Is Nothing Then
				cIDisplay_getNextPropertyItem = m_oProperties(m_lngCurrentPropItem)
			End If
		Else
			m_lngCurrentPropItem = m_lngCurrentPropItem - 1
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			Debug.Print(TypeName(Me) & ":cIDisplay:getNextPropertyItem -- nextItem exceeds Property Count.")
		End If
	End Function
	Private Function cIDisplay_getPropertyItemByIndex(ByVal iIndex As Integer) As cPropertyItem Implements _cIDisplay.getPropertyItemByIndex
		On Error Resume Next
		cIDisplay_getPropertyItemByIndex = m_oProperties(iIndex)
	End Function
	'/////////////////////////////////////////////
	'//Implemented cINode Properties and Functions
	Private Function cINode_AddChild(ByRef oBase As _cINode) As Boolean Implements _cINode.AddChild
		If m_lngMaxChildren = m_lngChildCount Then
			cINode_AddChild = False
		Else
			m_lngChildCount = m_lngChildCount + 1
			ReDim Preserve m_oChildren(m_lngChildCount - 1)
			m_oChildren(m_lngChildCount - 1) = oBase
			cINode_AddChild = True
		End If
	End Function
	Private Function cINode_getChildrenByClassName(ByRef Classname As String, ByRef hChilds() As Integer) As Boolean Implements _cINode.getChildrenByClassName
	End Function
	Private Function cINode_getChildIndexByHandle(ByVal h As Integer) As Integer Implements _cINode.getChildIndexByHandle
		Dim i As Integer
		Dim lRet As Integer
		lRet = -1
		For i = 0 To m_lngChildCount - 1
			If m_oChildren(i).Handle = h Then lRet = i : Exit For
		Next 
		cINode_getChildIndexByHandle = lRet
	End Function
	Private Function cINode_getChild(ByVal lngIndex As Integer) As _cINode Implements _cINode.getChild
		If (lngIndex >= 0) And (m_lngChildCount > 0) And (lngIndex <= m_lngChildCount - 1) Then
			cINode_getChild = m_oChildren(lngIndex)
		End If
	End Function
	Private Function cINode_removeChild(ByVal lngIndex As Integer) As Boolean Implements _cINode.removeChild
		Dim i As Integer
		If (lngIndex <= m_lngChildCount - 1) And (lngIndex >= 0) Then
			'UPGRADE_NOTE: Object m_oChildren() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			m_oChildren(lngIndex) = Nothing
			For i = lngIndex + 1 To m_lngChildCount - 1
				m_oChildren(i - 1) = m_oChildren(i)
			Next 
		End If
		m_lngChildCount = m_lngChildCount - 1
		If m_lngChildCount > 0 Then
			ReDim Preserve m_oChildren(m_lngChildCount - 1)
		Else
			Erase m_oChildren
		End If
	End Function
	Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Integer)
		Dim i As Integer
		
		If iMode = modGVDXMLSchemaConstants.GVD_XML_TYPE.cmp Then
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sName = op.Load(XML_NODE_NAME)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sDescription = op.Load(XML_NODE_DESCRIPTION)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngChildCount = op.Load(XML_NODE_CHILDCOUNT)
			If m_lngChildCount > 0 Then
				ReDim m_oChildren(m_lngChildCount - 1)
				For i = 0 To m_lngChildCount - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_oChildren(i) = op.Load(XML_NODE_CHILD & i)
					m_oChildren(i).Parent = m_hMe
					System.Diagnostics.Debug.Assert(m_oChildren(i).Parent <> 0, "")
				Next 
			End If
		Else 'DEF
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngMaxChildren = op.Load(XML_NODE_MAXCHILDREN)
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sImage = op.Load(XML_NODE_IMAGE)
			
			' load properties last, these will reference variables initialized above
			'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lngPropCount = op.Load(XML_NODE_PROPERTYCOUNT)
			If m_lngPropCount > 0 Then
				ReDim m_oProperties(m_lngPropCount - 1)
				For i = 0 To m_lngPropCount - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_oProperties(i) = op.Load(XML_NODE_PROPERTY & i)
				Next 
			End If
		End If
	End Sub
	Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
	End Sub
End Class