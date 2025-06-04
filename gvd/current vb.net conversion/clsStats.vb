Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsStats_NET.clsStats")> Public Class clsStats
	
	'//////////////////////////////////////////////////////
	' Total Vehicle Stats
	Private mvarFloatationRating As Single
	Private mvarTotalPrice As Single 'todo: rename these to just price, surface, weight, etc to be consistant. In other words, we dont need "total" in each variables name
	Private mvarTotalSurfaceArea As Single
	Private mvarStructuralSurfaceArea As Single
	Private mvarStructuralCost As Single
	Private mvarStructuralWeight As Single
	Private mvarEmptyWeight As Single
	Private mvarHLoadedWeight As Single
	Private mvarLoadedWeight As Single
	Private mvarLoadedMass As Single
	Private mvarHLoadedMass As Single
	Private mvarTotalVolume As Single
	Private mvarSubmergedWeight As Single
	Private mvarSubmergedMass As Single
	Private mvarHSubmergedWeight As Single ' with hardpoints loaded
	Private mvarHSubmergedMass As Single
	Private mvarUsualInternalPayload As Single
	Private mvarStructuralHealth As Integer 'always with hardpoints loaded
	Private mvarSizeModifier As Short
	Private mvarTotalContinuousPower As Single 'todo: these may need to be stats of the power profiles since it depends on what devices the user is running
	Private mvarTotalContinuousPowerConsumption As Single
	Private mvarTotalStoredPower As Single
	Private mvarTotalStoredPowerConsumption As Single
	Private mvarTotalHardPointConnections As Integer
	
	' option settings which should probably be moved to a dialog, but still loaded
	' via a class (e.g. cOptions duh?(
	Private mvarBattleSuit As String ' setting type "none", "form-fitting" etc
	Private mvarQuality As String
	Private mvarPerPersonWeight As Single
	Private mvarPerCargoWeight As Single
	Private mvarRecommendedPayload As Single
	Private mvarRecommendedAccessSpace As Single
	Private mvarAccessSpaceVolumeMod As Single
	Private mvarUseHardpointMountedWeights As Boolean
	Private mvarUseSurfaceAreaTable As Boolean
	
	Public Property BattleSuit() As String
		Get
			BattleSuit = mvarBattleSuit
		End Get
		Set(ByVal Value As String)
			mvarBattleSuit = Value
		End Set
	End Property
	Public Property Quality() As String
		Get
			Quality = mvarQuality
		End Get
		Set(ByVal Value As String)
			mvarQuality = Value
		End Set
	End Property
	Public Property UseSurfaceAreaTable() As Boolean
		Get
			UseSurfaceAreaTable = mvarUseSurfaceAreaTable
		End Get
		Set(ByVal Value As Boolean)
			mvarUseSurfaceAreaTable = Value
		End Set
	End Property
	Public Property UseHardpointMountedWeights() As Boolean
		Get
			UseHardpointMountedWeights = mvarUseHardpointMountedWeights
		End Get
		Set(ByVal Value As Boolean)
			mvarUseHardpointMountedWeights = Value
		End Set
	End Property
	Public Property RecommendedAccessSpace() As Boolean
		Get
			RecommendedAccessSpace = mvarRecommendedAccessSpace
		End Get
		Set(ByVal Value As Boolean)
			mvarRecommendedAccessSpace = Value
		End Set
	End Property
	Public Property RecommendedPayload() As Boolean
		Get
			RecommendedPayload = mvarRecommendedPayload
		End Get
		Set(ByVal Value As Boolean)
			mvarRecommendedPayload = Value
		End Set
	End Property
	Public Property PerPersonWeight() As Single
		Get
			PerPersonWeight = mvarPerPersonWeight
		End Get
		Set(ByVal Value As Single)
			mvarPerPersonWeight = Value
		End Set
	End Property
	Public Property PerCargoWeight() As Single
		Get
			PerCargoWeight = mvarPerCargoWeight
		End Get
		Set(ByVal Value As Single)
			mvarPerCargoWeight = Value
		End Set
	End Property
	Public Property AccessSpaceVolumeMod() As Single
		Get
			AccessSpaceVolumeMod = mvarAccessSpaceVolumeMod
		End Get
		Set(ByVal Value As Single)
			mvarAccessSpaceVolumeMod = Value
		End Set
	End Property
	
	
	Public Property TotalHardPointConnections() As Integer
		Get
			TotalHardPointConnections = mvarTotalHardPointConnections
		End Get
		Set(ByVal Value As Integer)
			mvarTotalHardPointConnections = Value
		End Set
	End Property
	
	Public Property UsualInternalPayload() As Double
		Get
			UsualInternalPayload = mvarUsualInternalPayload
		End Get
		Set(ByVal Value As Double)
			mvarUsualInternalPayload = Value
		End Set
	End Property
	Public Property StructuralSurfaceArea() As Double
		Get
			StructuralSurfaceArea = mvarStructuralSurfaceArea
		End Get
		Set(ByVal Value As Double)
			mvarStructuralSurfaceArea = Value
		End Set
	End Property
	Public Property totalSurfaceArea() As Double
		Get
			totalSurfaceArea = mvarTotalSurfaceArea
		End Get
		Set(ByVal Value As Double)
			mvarTotalSurfaceArea = Value
		End Set
	End Property
	Public Property StructuralCost() As Double
		Get
			StructuralCost = mvarStructuralCost
		End Get
		Set(ByVal Value As Double)
			mvarStructuralCost = Value
		End Set
	End Property
	Public Property StructuralWeight() As Double
		Get
			StructuralWeight = mvarStructuralWeight
		End Get
		Set(ByVal Value As Double)
			mvarStructuralWeight = Value
		End Set
	End Property
	Public Property SizeModifier() As Short
		Get
			SizeModifier = mvarSizeModifier
		End Get
		Set(ByVal Value As Short)
			mvarSizeModifier = Value
		End Set
	End Property
	Public Property TotalPrice() As Single
		Get
			TotalPrice = mvarTotalPrice
		End Get
		Set(ByVal Value As Single)
			mvarTotalPrice = Value
		End Set
	End Property
	Public Property StructuralHealth() As Integer
		Get
			StructuralHealth = mvarStructuralHealth
		End Get
		Set(ByVal Value As Integer)
			mvarStructuralHealth = Value
		End Set
	End Property
	Public Property EmptyWeight() As Double
		Get
			EmptyWeight = mvarEmptyWeight
		End Get
		Set(ByVal Value As Double)
			mvarEmptyWeight = Value
		End Set
	End Property
	Public Property HSubmergedWeight() As Double
		Get
			HSubmergedWeight = mvarHSubmergedWeight
		End Get
		Set(ByVal Value As Double)
			mvarHSubmergedWeight = Value
		End Set
	End Property
	Public Property SubmergedWeight() As Double
		Get
			SubmergedWeight = mvarSubmergedWeight
		End Get
		Set(ByVal Value As Double)
			mvarSubmergedWeight = Value
		End Set
	End Property
	Public Property HSubmergedMass() As Double
		Get
			HSubmergedMass = mvarHSubmergedMass
		End Get
		Set(ByVal Value As Double)
			mvarHSubmergedMass = Value
		End Set
	End Property
	Public Property SubmergedMass() As Double
		Get
			SubmergedMass = mvarSubmergedMass
		End Get
		Set(ByVal Value As Double)
			mvarSubmergedMass = Value
		End Set
	End Property
	Public Property FloatationRating() As Single
		Get
			FloatationRating = mvarFloatationRating
		End Get
		Set(ByVal Value As Single)
			mvarFloatationRating = Value
		End Set
	End Property
	Public Property HLoadedWeight() As Double
		Get
			HLoadedWeight = mvarHLoadedWeight
		End Get
		Set(ByVal Value As Double)
			mvarHLoadedWeight = Value
		End Set
	End Property
	Public Property HLoadedMass() As Double
		Get
			HLoadedMass = mvarHLoadedMass
		End Get
		Set(ByVal Value As Double)
			mvarHLoadedMass = Value
		End Set
	End Property
	Public Property LoadedWeight() As Double
		Get
			LoadedWeight = mvarLoadedWeight
		End Get
		Set(ByVal Value As Double)
			mvarLoadedWeight = Value
		End Set
	End Property
	Public Property TotalVolume() As Double
		Get
			TotalVolume = mvarTotalVolume
		End Get
		Set(ByVal Value As Double)
			mvarTotalVolume = Value
		End Set
	End Property
	Public Property LoadedMass() As Double
		Get
			LoadedMass = mvarLoadedMass
		End Get
		Set(ByVal Value As Double)
			mvarLoadedMass = Value
		End Set
	End Property
	Public Property TotalContinuousPower() As Single
		Get
			TotalContinuousPower = mvarTotalContinuousPower
		End Get
		Set(ByVal Value As Single)
			mvarTotalContinuousPower = Value
		End Set
	End Property
	Public Property TotalStoredPower() As Single
		Get
			TotalStoredPower = mvarTotalStoredPower
		End Get
		Set(ByVal Value As Single)
			mvarTotalStoredPower = Value
		End Set
	End Property
	Public Property TotalContinuousPowerConsumption() As Single
		Get
			TotalContinuousPowerConsumption = mvarTotalContinuousPowerConsumption
		End Get
		Set(ByVal Value As Single)
			mvarTotalContinuousPowerConsumption = Value
		End Set
	End Property
	Public Property TotalStoredPowerConsumption() As Single
		Get
			TotalStoredPowerConsumption = mvarTotalStoredPowerConsumption
		End Get
		Set(ByVal Value As Single)
			mvarTotalStoredPowerConsumption = Value
		End Set
	End Property
	
	
	'//////////////////////////////////////////////////////////////////////
	'//////////////////////////////////////////////////////////////////////
	'//////////////////////////////////////////////////////////////////////
	' Statistics Calculation Functions
	'//////////////////////////////////////////////////////////////////////
	'//////////////////////////////////////////////////////////////////////
	Sub Update()
		Dim surface As Object
		' see pages 25 - 26 of Vehicles
		Dim element As Object
		Dim lngStart As Integer
		Dim lngStop As Integer
		On Error Resume Next
		
		'//first, recalc every single component's stats
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.StatsUpdate. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			element.StatsUpdate()
		Next element
		'calc total structural stats and have each sub update
		'its own surface area, volume and weight stats
		'this routine saves the Vehicle's Structural Surface area and Weight
		Call Me.CalcStructuralStats()
		
		'calc weight and cost for all options and surface features
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Veh.crew.UseRecommendedCrew Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Veh.crew.StatsUpdate()
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Veh.Options.CalcOptionsWeightandCost()
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Veh.surface.Update()
		
		System.Windows.Forms.Application.DoEvents()
		'calc vehicles total empty weight, loadedweight, etc
		Call CalcWeight()
		
		
		'calc total volume of vehicle
		mvarTotalVolume = Me.CalcTotalVolume
		mvarTotalPrice = Me.CalcPrice ' calc final price of vehicle
		mvarSizeModifier = Me.CalcSizeModifier(mvarTotalVolume)
		mvarTotalContinuousPower = Me.CalcTotalGeneratedPower
		mvarTotalContinuousPowerConsumption = Me.CalcTotalContinuousPowerConsumption
		
		' If vehicle is submersible, calculate the submerged weight and mass
		'UPGRADE_WARNING: Couldn't resolve default property of object surface.Submersible. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If surface.Submersible Then
			mvarSubmergedWeight = Me.CalcSubmergedWeight(mvarLoadedWeight, mvarTotalVolume)
			mvarSubmergedMass = mvarSubmergedWeight / 2000
		Else
			mvarSubmergedWeight = 0
			mvarSubmergedMass = 0
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarStructuralHealth = Me.CalcHealth(Veh.Components(BODY_KEY).HitPoints, mvarHLoadedWeight)
		
		' re-calc vehicle performance figures
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.PerformanceProfiles
			'UPGRADE_WARNING: Couldn't resolve default property of object element.CalcPerformance. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			element.CalcPerformance()
		Next element
		
	End Sub
	
	Public Function CalcFloatationRating() As Single
		Dim mvarVolume As Object
		'todo: single or double return value?
		Dim Floatation As Single ' the floatation multiplier
		
		' Calculate the floatation rating (must be done after volume)
		
		' to have a floatation rating, it must have floatation enabled.
		' Note that subermisble hulls require the floatation option
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Veh.surface.Floatation) Then
			Floatation = 62.5
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not (Veh.surface.Submersible) Then
				' submersible hulls will have 62.5 rating irregardless of line type
				' so if its not a submersible hull, we check the lines to reduce the rating
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case Veh.surface.HydrodynamicLines
					Case "none"
						Floatation = 62.5
					Case "mediocre"
						Floatation = 57
					Case "average"
						Floatation = 52
					Case "submarine"
						Floatation = 62.5
					Case "fine"
						Floatation = 48
					Case "very fine"
						Floatation = 45
				End Select
			End If
		Else
			Floatation = 0
		End If
		
		'todo: the floatation rating calc below is correct for a floating ship, but not for a submersible one! since
		'      those require the volume of all supers and turrets as well as body!!!
		'      This stat should be MOVED to the stats class anyways!  Its not a body stat!
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarVolume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarFloatationRating = System.Math.Round(Floatation * mvarVolume, 2)
	End Function
	Public Sub CalcStructuralStats()
		Dim temparea As Single
		Dim TempTotalArea As Single
		Dim TempCost As Single
		Dim TempWeight As Single
		Dim element As Object
		' Calculate the Total Structural and Total Surface Area Weight of the Vehicle
		'NOTE: This is NOT supposed to use components, ONLY other SUB ASSEMBLIES
		'todo: optimize to only check on "subassemblies" or depending on changes to code
		' architecture, "container" type objects sans "group object"
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case Body, Skid, Wheel, Track, Leg, Wing, AutogyroRotor, TTRotor, CARotor, MMRotor, Hydrofoil, Hovercraft, Superstructure, Pod, Turret, Popturret, Arm
					'element.StatsUpdate 'Note this call must be made
					'(NOTE: The above is commented out since im now updating ALL STATS
					'before hand in CalcStats)
					'since Volumes for most of the above Subassemblies takes
					'the body's Volume into account.  Note that the Body must
					'statupdate FIRST as a result which is ensured here
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SurfaceArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					temparea = element.SurfaceArea + temparea
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempWeight = element.Weight + TempWeight
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Cost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempCost = element.Cost + TempCost
				Case Mast, Gasbag, OpenMount
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SurfaceArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempTotalArea = element.SurfaceArea + TempTotalArea
			End Select
		Next element
		
		
		mvarStructuralSurfaceArea = temparea
		mvarTotalSurfaceArea = temparea + TempTotalArea
		mvarStructuralCost = TempCost
		mvarStructuralWeight = TempWeight
		
	End Sub
	
	Sub CalcWeight()
		Dim FuelWeight As Single
		Dim AmmoWeight As Single
		Dim ProvisionsWeight As Single
		Dim GunCarriagesWeight As Single
		Dim AuxVehiclesWeight As Single
		Dim HardPointWeight As Single
		Dim CargoWeight As Single
		Dim lngObjectsConnectedtoHardpoints As Integer
		Dim ComponentsWeight As Single
		
		Dim element As Object ' any object in the collection
		' add up the weight of all componenets, structures, features except for Fuel, Ammunition
		' provisions, any carried vehicles, and TL5- guns that require carriages.
		' the resulting sum is the Empty Weight
		On Error GoTo errorhandler
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			
			If TypeOf element Is clsBody Then
				With element
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ComponentsWeight = ComponentsWeight + element.Weight
				End With
				
				'ElseIf TypeOf element Is clsWeaponLink Then '07/09/02 MPJ OBSOLETE. Weaponlinks no longer stored in Components collection
				
			ElseIf TypeOf element Is clsCargo Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.CargoWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CargoWeight = CargoWeight + element.CargoWeight '//we get cargo weight in the Usual internal Payload
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ComponentsWeight = ComponentsWeight + element.Weight
				
			ElseIf TypeOf element Is clsHardPoint Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Options.UseHardpointMountedWeights = False Then
					'UPGRADE_WARNING: Couldn't resolve default property of object element.LoadCapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Quantity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					HardPointWeight = HardPointWeight + (element.Quantity * element.LoadCapacity)
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ComponentsWeight = ComponentsWeight + element.Weight
				'//note equipment pods can only attach to hardpoints so thats where
				'//we get their weights
			ElseIf TypeOf element Is clsSoftware Then 
				
			ElseIf TypeOf element Is clsLiftingGas Then 
				
				'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsHardPoint Then 
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Options.UseHardpointMountedWeights Then
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					HardPointWeight = HardPointWeight + element.Weight
				End If
				'//count up how many objects are actually attached to hardpoints so
				'//we can use these in drag calculations later
				'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(element.LogicalParent).Datatype = HardPoint Then
					lngObjectsConnectedtoHardpoints = lngObjectsConnectedtoHardpoints + 1
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf TypeOf Veh.Components(element.LogicalParent) Is clsEquipmentPod Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Options.UseHardpointMountedWeights Then
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					HardPointWeight = HardPointWeight + element.Weight
				End If
				'//count up how many objects are actually attached to hardpoints so
				'//we can use these in drag calculations later
				'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Veh.Components(element.LogicalParent).Datatype = HardPoint Then
					lngObjectsConnectedtoHardpoints = lngObjectsConnectedtoHardpoints + 1
				End If
			ElseIf TypeOf element Is clsFuelTank Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.FuelWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FuelWeight = FuelWeight + element.FuelWeight
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ComponentsWeight = ComponentsWeight + element.Weight
				
			ElseIf TypeOf element Is clsProvisions Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ProvisionsWeight = ProvisionsWeight + element.Weight
				
			ElseIf TypeOf element Is clsVehicleStorage Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.CraftWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AuxVehiclesWeight = AuxVehiclesWeight + element.CraftWeight
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ComponentsWeight = ComponentsWeight + element.Weight
				
			ElseIf TypeOf element Is clsWeaponAmmunition Then  'check for ammunition
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AmmoWeight = AmmoWeight + element.Weight
				
			ElseIf TypeOf element Is clsWeaponGun Then  'check for guns with carriages
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Carriage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If element.Carriage Then
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GunCarriagesWeight = GunCarriagesWeight + element.Weight
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ComponentsWeight = ComponentsWeight + element.Weight
				End If
				
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ComponentsWeight = ComponentsWeight + element.Weight
			End If
		Next element
		
		'save these overall vehicle stats
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarUsualInternalPayload = (Veh.Options.PerPersonWeight * Veh.crew.TotalNumberCrewPassengers) + CargoWeight
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarEmptyWeight = ComponentsWeight + Veh.Options.OptionsWeight
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarLoadedWeight = mvarEmptyWeight + mvarUsualInternalPayload + Veh.Options.RollStabilizersWeight + FuelWeight + AmmoWeight + ProvisionsWeight + AuxVehiclesWeight + Veh.Options.TotalVariableSweepWeight + Veh.Options.TotalFoldingWingWeight + Veh.Options.TotalCompartmentalizationWeight
		mvarLoadedMass = mvarLoadedWeight / 2000
		mvarHLoadedWeight = mvarLoadedWeight + HardPointWeight
		mvarHLoadedMass = mvarHLoadedWeight / 2000
		mvarTotalHardPointConnections = lngObjectsConnectedtoHardpoints
		
		Exit Sub
		
errorhandler: 
		Debug.Print("CalcWeight: " & Err.Description)
	End Sub
	
	Function CalcSubmergedWeight(ByRef Lweight As Single, ByRef TVolume As Single) As Single
		Dim TempSweight As Single
		Const Multiplier As Double = 62.5
		
		TempSweight = Multiplier * TVolume
		
		If TempSweight < Lweight Then
			TempSweight = Lweight
		End If
		
		CalcSubmergedWeight = TempSweight
		'TODO: Do i need to do a submerged weight with hardpoints?
		'i've already got a HSubmergedWeight and HSubmergedMass properties in the clsBody
	End Function
	
	Function CalcTotalVolume() As Single
		
		Dim TempVolume As Single ' holds the volume in progress
		Dim element As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype 'Only subassemblies are computed in totalvolume
				Case Mast, OpenMount, Gasbag, Body, Skid, Wheel, Track, Leg, Wing, AutogyroRotor, TTRotor, CARotor, MMRotor, Hydrofoil, Hovercraft, Superstructure, Pod, Turret, Popturret, Arm
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Volume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TempVolume = TempVolume + element.Volume
					'todo: are masts to be included here? gasbags too?
			End Select
		Next element
		CalcTotalVolume = TempVolume
	End Function
	
	Function CalcPrice() As Single
		' Add up total cost of everything built into the vehicle except for ammunition
		' and fuel to find out how much it costs
		Dim TempPrice As Single ' holds the volume in progress
		Dim element As Object
		Dim sQuality As String
		Dim QMod As Single
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			If Not TypeOf element Is clsWeaponAmmunition Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Cost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TempPrice = TempPrice + element.Cost
			End If
		Next element
		TempPrice = TempPrice
		
		'Caculate the Vehicle Quality Modifiers
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sQuality = Veh.Options.Quality
		If sQuality = "standard" Then
			QMod = 1
		ElseIf sQuality = "cheap" Then 
			QMod = 0.5
		ElseIf sQuality = "fine" Then 
			QMod = 4
		ElseIf sQuality = "very fine" Then 
			QMod = 20
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalcPrice = (TempPrice + Veh.Options.OptionsCost) * QMod
	End Function
	
	Function CalcHealth(ByRef BodyHitPoints As Double, ByRef LoadedWeight As Single) As Integer
		Dim TempHealth As Double
		Dim HMod As Short
		Dim sQuality As String
		
		' structural HT = (200 * BodyHitPoints / LoadedWeight) + 5
		' If the vehicle has Hardpoints, always use the weight WITH hardpoints loaded -
		' do NOT use two different values
		' Round HT to the nearest whole number.
		If LoadedWeight = 0 Then
			TempHealth = 1
		Else
			TempHealth = System.Math.Round((200 * BodyHitPoints / LoadedWeight) + 5, 0)
		End If
		
		'Note: im assume maximum HT of 12 STILL applies. If users disagree
		'i can move the below HMods below the Max HT check below
		'Caculate the Vehicle Quality Modifiers
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sQuality = Veh.Options.Quality
		If sQuality = "standard" Then
			HMod = 0
		ElseIf sQuality = "cheap" Then 
			HMod = -1
		ElseIf sQuality = "fine" Then 
			HMod = 1
		ElseIf sQuality = "very fine" Then 
			HMod = 2
		End If
		
		TempHealth = TempHealth + HMod
		
		' The maximum allowed structural HT is 12 or the vehicle's TL, whichever is greater.
		If TempHealth > 12 Then
			If gVehicleTL > 12 Then
				TempHealth = Val(CStr(gVehicleTL))
			Else
				TempHealth = 12
			End If
		End If
		
		CalcHealth = TempHealth
	End Function
	
	Function CalcSizeModifier(ByVal Volume As Double) As Integer
		' todo: will try to optimize later when i finish. Maybe.
		Dim TempModifier As Integer
		Dim TVolume As Double
		Dim y As Single
		Dim i As Integer
		
		TempModifier = -4
		TVolume = 0.1
		y = 0.7
		i = 0
		
		If Volume <= TVolume Then
			CalcSizeModifier = -4
			Exit Function
		End If
		
		Do 
			TVolume = TVolume * 3
			TempModifier = TempModifier + 1
			If Volume <= TVolume Then
				Exit Do
			End If
			TVolume = TVolume + (y * 10 ^ i)
			TempModifier = TempModifier + 1
			If Volume <= TVolume Then
				Exit Do
			End If
			i = i + 1
		Loop 
		CalcSizeModifier = TempModifier
	End Function
	
	Function CalcTotalGeneratedPower() As Double
		Dim dlbstored As Object
		'add all the outputs of all the Power Systems from the Vehicle
		Dim TempOutput As Double
		Dim dblStored As Double
		Dim Keys() As String
		Dim i As Integer
		
		On Error GoTo errorhandler
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Keys = VB6.CopyArray(Veh.KeyManager.GetCurrentPowerSystemKeys)
		
		For i = 1 To UBound(Keys)
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If TypeOf Veh.Components(Keys(i)) Is clsEnergyBank Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object dlbstored. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dblStored = dlbstored + Veh.Components(Keys(i)).Output
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TempOutput = TempOutput + Veh.Components(Keys(i)).Output
			End If
		Next 
		
		mvarTotalStoredPower = dblStored
		CalcTotalGeneratedPower = TempOutput
		Exit Function
		
errorhandler: 
		If Err.Number = 9 Then 'array not dimensioned properly.  Return a value of 0
			CalcTotalGeneratedPower = 0
			Exit Function
		End If
	End Function
	
	Function CalcTotalContinuousPowerConsumption() As Double
		'add all the outputs of all the Power Systems from the Vehicle
		Dim TempPower As Double
		Dim Keys() As String
		Dim i As Integer
		
		'todo: we cant determine total power consumed FROM STORED POWER suppliers
		' unless we use a "Primary" power config setting so that we know which
		' configuration to base this stat on.
		On Error GoTo errorhandler
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.KeyManager. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Keys = VB6.CopyArray(Veh.KeyManager.GetCurrentPowerConsumptionKeys)
		
		For i = 1 To UBound(Keys)
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TempPower = TempPower + Veh.Components(Keys(i)).PowerReqt
		Next 
		
		CalcTotalContinuousPowerConsumption = TempPower
		Exit Function
errorhandler: 
		If Err.Number = 9 Then 'array not dimensioned properly.  Return a value of 0
			CalcTotalContinuousPowerConsumption = 0
			Exit Function
		End If
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		mvarTotalSurfaceArea = 0
		mvarStructuralSurfaceArea = 0
		
		mvarEmptyWeight = 0
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Sub CalcOptionsWeightandCost()
		Dim mvarOptionsCost As Object
		Dim PuntureResistantCost As Object
		Dim mvarTotalSnowTiresCost As Object
		Dim mvarOptionsWeight As Object
		Dim mvarRollStabilizers As Object
		Dim mvarRollStabilizersWeight As Object
		Dim mvarRollStabilizersCost As Object
		Dim mvarPin As Object
		Dim mvarPinWeight As Object
		Dim mvarPinCost As Object
		Dim mvarHitch As Object
		Dim mvarHitchWeight As Object
		Dim mvarHitchCost As Object
		Dim mvarConvertible As Object
		Dim mvarConvertibleWeight As Object
		Dim mvarConvertibleCost As Object
		Dim mvarPlow As Object
		Dim mvarPlowCost As Object
		Dim mvarPlowWeight As Object
		Dim mvarBulldozer As Object
		Dim mvarBullDozerCost As Object
		Dim mvarBullDozerWeight As Object
		Dim mvarRam As Object
		Dim mvarRamCost As Object
		Dim mvarRamWeight As Object
		
		Dim element As Object
		Dim TotalArea As Single
		Dim BodyArea As Single
		Dim BodyHits As Integer
		Dim TrimmedArea As Single
		Dim IgnoredArea As Single
		
		Dim compartmentalizationWeight As Single
		Dim compartmentalizationCost As Single
		Dim ControlledInstabilityCost As Single
		Dim VariableSweepCost As Single
		Dim VariableSweepWeight As Single
		Dim foldingwingsWeight As Single
		Dim foldingwingsCost As Single
		Dim ImprovedSuspensionCost As Single
		Dim SnowTiresCost As Single
		Dim RacingTiresCost As Single
		Dim PunctureResistantCost As Single
		Dim WheelBladesCost As Single
		Dim WheelBladesWeight As Single
		Dim OtherWheelCosts As Single
		
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In Veh.Components
			'update the weight and cost for the armor
			If TypeOf element Is clsArmor Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.StatsUpdate. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				element.StatsUpdate()
				'this is not a misprint.  See page92 under the table for rules
				'which say you do not include surfaceare for Skids, Gasbag and Masts
			ElseIf TypeOf element Is clsMast Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.SurfaceArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				IgnoredArea = IgnoredArea + element.SurfaceArea
			ElseIf TypeOf element Is clsSkid Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.SurfaceArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				IgnoredArea = IgnoredArea + element.SurfaceArea
			ElseIf TypeOf element Is clsGasbag Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.SurfaceArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				IgnoredArea = IgnoredArea + element.SurfaceArea
				
			ElseIf TypeOf element Is clsBody Then 
				With element
					'UPGRADE_WARNING: Couldn't resolve default property of object element.CalcCompartmentalizationStats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.CalcCompartmentalizationStats()
					'UPGRADE_WARNING: Couldn't resolve default property of object element.compartmentalizationWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					compartmentalizationWeight = compartmentalizationWeight + .compartmentalizationWeight
					'UPGRADE_WARNING: Couldn't resolve default property of object element.compartmentalizationCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					compartmentalizationCost = compartmentalizationCost + .compartmentalizationCost
					'UPGRADE_WARNING: Couldn't resolve default property of object element.ImprovedSuspensionCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ImprovedSuspensionCost = ImprovedSuspensionCost + .ImprovedSuspensionCost
				End With
			ElseIf TypeOf element Is clsWheel Then 
				With element
					'UPGRADE_WARNING: Couldn't resolve default property of object element.ImprovedSuspensionCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ImprovedSuspensionCost = ImprovedSuspensionCost + .ImprovedSuspensionCost
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SnowTiresCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					SnowTiresCost = SnowTiresCost + .SnowTiresCost
					'UPGRADE_WARNING: Couldn't resolve default property of object element.RacingTiresCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					RacingTiresCost = RacingTiresCost + .RacingTiresCost
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PunctureResistantCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PunctureResistantCost = PunctureResistantCost + .PunctureResistantCost
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SmartWheelsCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.AllWheelSteeringCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.ImprovedBrakesCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					OtherWheelCosts = OtherWheelCosts + .ImprovedBrakesCost + .AllWheelSteeringCost + .SmartWheelsCost
					'UPGRADE_WARNING: Couldn't resolve default property of object element.WheelBladesCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					WheelBladesCost = WheelBladesCost + .WheelBladesCost
					'UPGRADE_WARNING: Couldn't resolve default property of object element.WheelBladesWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					WheelBladesWeight = WheelBladesWeight + .WheelBladesWeight
					
				End With
			ElseIf (TypeOf element Is clsSkid) Or (TypeOf element Is clsLeg) Or (TypeOf element Is clsTrack) Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object element.ImprovedSuspensionCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ImprovedSuspensionCost = ImprovedSuspensionCost + element.ImprovedSuspensionCost
			ElseIf (TypeOf element Is clsSuperStructure) Or (TypeOf element Is clsTurret) Or (TypeOf element Is clsPopTurret) Then 
				With element
					'UPGRADE_WARNING: Couldn't resolve default property of object element.CalcCompartmentalizationStats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.CalcCompartmentalizationStats()
					'UPGRADE_WARNING: Couldn't resolve default property of object element.compartmentalizationWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					compartmentalizationWeight = compartmentalizationWeight + .compartmentalizationWeight
					'UPGRADE_WARNING: Couldn't resolve default property of object element.compartmentalizationCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					compartmentalizationCost = compartmentalizationCost + .compartmentalizationCost
				End With
			ElseIf TypeOf element Is clsWing Then 
				With element
					'UPGRADE_WARNING: Couldn't resolve default property of object element.CalcWingRotorOptionWeightsAndCosts. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.CalcWingRotorOptionWeightsAndCosts()
					'UPGRADE_WARNING: Couldn't resolve default property of object element.FoldingWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					foldingwingsWeight = foldingwingsWeight + .FoldingWeight
					'UPGRADE_WARNING: Couldn't resolve default property of object element.VariableSweepWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VariableSweepWeight = VariableSweepWeight + .VariableSweepWeight
					'UPGRADE_WARNING: Couldn't resolve default property of object element.VariableSweepCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VariableSweepCost = VariableSweepCost + .VariableSweepCost
					'UPGRADE_WARNING: Couldn't resolve default property of object element.ControlledInstabilityCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ControlledInstabilityCost = ControlledInstabilityCost + .ControlledInstabilityCost
				End With
			ElseIf TypeOf element Is clsRotor Then 
				With element
					'UPGRADE_WARNING: Couldn't resolve default property of object element.CalcWingRotorOptionWeightsAndCosts. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.CalcWingRotorOptionWeightsAndCosts()
					'UPGRADE_WARNING: Couldn't resolve default property of object element.FoldingWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					foldingwingsWeight = foldingwingsWeight + .FoldingWeight
					'UPGRADE_WARNING: Couldn't resolve default property of object element.FoldingCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					foldingwingsCost = foldingwingsCost + .FoldingCost
					'UPGRADE_WARNING: Couldn't resolve default property of object element.ControlledInstabilityCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ControlledInstabilityCost = ControlledInstabilityCost + .ControlledInstabilityCost
				End With
			End If
		Next element
		
		
		' Get weight and cost for Rams
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRam. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarRam Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarRamWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarRamWeight = 1 * BodyArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarRamCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarRamCost = 2 * BodyArea
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarRamWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarRamWeight = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarRamCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarRamCost = 0
		End If
		' Get weight and cost for Bulldozers
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarBulldozer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarBulldozer Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarBullDozerWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarBullDozerWeight = 2 * BodyArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarBullDozerCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarBullDozerCost = 4 * BodyArea
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarBullDozerWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarBullDozerWeight = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarBullDozerCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarBullDozerCost = 0
		End If
		' Get wieght and cost for Plows
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPlow. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarPlow Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPlowWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPlowWeight = 2 * BodyArea
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPlowCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPlowCost = 4 * BodyArea
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPlowWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPlowWeight = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPlowCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPlowCost = 0
		End If
		' Get weight and cost for Convertible
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarConvertible. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarConvertible = "none" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarConvertibleCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarConvertibleCost = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarConvertibleWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarConvertibleWeight = 0
		Else
		End If
		' Get weight and cost for Hitch
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarHitch. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarHitch Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarHitchCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarHitchCost = 0.1 * BodyHits
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarHitchCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarHitchWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarHitchWeight = mvarHitchCost
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarHitchCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarHitchCost = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarHitchWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarHitchWeight = 0
		End If
		' Get weight and cost for Pin
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPin. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarPin <> "none" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPinCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPinCost = 0.05 * BodyHits
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPinWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPinWeight = 0.1 * BodyHits
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPin. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If mvarPin = "explosive" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarPinCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarPinCost = mvarPinCost * 5
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPinCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPinCost = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarPinWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarPinWeight = 0
		End If
		'Get weight and cost for Roll Stabilizers
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRollStabilizers. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mvarRollStabilizers = False Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarRollStabilizersCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarRollStabilizersCost = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarRollStabilizersWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarRollStabilizersWeight = 0
		Else
			'do divide by zero check
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Veh.Stats.StructuralSurfaceArea = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarRollStabilizersCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarRollStabilizersCost = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarRollStabilizersWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarRollStabilizersWeight = 0
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarRollStabilizersCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarRollStabilizersCost = 0.1 * (BodyArea / Veh.Stats.StructuralSurfaceArea) * Veh.Stats.StructuralCost
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarRollStabilizersWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarRollStabilizersWeight = 0.05 * (BodyArea / Veh.Stats.StructuralSurfaceArea) * Veh.Stats.StructuralWeight
			End If
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRamWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarBullDozerWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPlowWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarHitchWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPinWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarConvertibleWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRollStabilizersWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarOptionsWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarOptionsWeight = mvarRollStabilizersWeight + mvarConvertibleWeight + mvarPinWeight + mvarHitchWeight + mvarPlowWeight + mvarBullDozerWeight + mvarRamWeight + WheelBladesWeight + compartmentalizationWeight + VariableSweepWeight + foldingwingsWeight
		
		'note the MagicLevitationEnergyCost is not in dollars but in units of Mana or Energy right? thats why its not added to OptionsCost
		'UPGRADE_WARNING: Couldn't resolve default property of object PuntureResistantCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarTotalSnowTiresCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRamCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarBullDozerCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPlowCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarHitchCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPinCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarConvertibleCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarRollStabilizersCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarOptionsCost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarOptionsCost = mvarRollStabilizersCost + mvarConvertibleCost + mvarPinCost + mvarHitchCost + mvarPlowCost + mvarBullDozerCost + mvarRamCost + VariableSweepCost + ControlledInstabilityCost + foldingwingsCost + ImprovedSuspensionCost + compartmentalizationCost + mvarTotalSnowTiresCost + RacingTiresCost + PuntureResistantCost + WheelBladesCost + OtherWheelCosts
	End Sub
End Class