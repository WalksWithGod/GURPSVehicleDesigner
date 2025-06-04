VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 1


#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    Private mlClassDebugID As Long
#End If
Private mvarClassDebugID As Variant 'local copy
'////////////////////////////////////////////////
'Standard SubAssembly Properties
'///////////////////////////////////////////////
Private mvarImage As Integer 'local copy
Private mvarSelectedImage As Integer 'local copy
Private mvarParent As String 'local copy
Private mvarKey As String 'local copy
Private mvarComponent As String 'local copy
Private mvarTL As Integer 'local copy
Private mvarSlopeR As String 'local copy
Private mvarSlopeL As String 'local copy
Private mvarSlopeF As String 'local copy
Private mvarSlopeB As String 'local copy
Private mvarRobotic As Boolean 'local copy
Private mvarCompartmentalization As String 'local copy
Private mvarResponsive As Boolean 'local copy
Private mvarMaterials As String 'local copy
Private mvarLivingMetal As Boolean 'local copy
Private mvarFrameStrength As String 'local copy
Private mvarBiomechanical As Boolean 'local copy
'///////////////////////////////////////////////////////
'Tracked Flags
'////////////////////////////////////////////////////////\
Private mvarRetractableWheelsSkids As String 'local copy
Private mvarWingsorRotors As Boolean 'local copy
Private mvarNumLegs As Integer 'local copy
Private mvarNumWings As Integer 'local copy
Private mvarImprovedSuspension As Boolean 'local copy
'//////////////////////////////////////////////
' Body Stats
'/////////////////////////////////////////////
Private mvarTotalRotationSpace As Double 'local copy
Private mvarMinimumVolume As Double 'local copy
Private mvarFloatationRating As Double 'local copy
Private mvarVolume As Double 'local copy
Private mvarSurfaceArea As Double 'local copy
Private mvarCost As Double 'local copy
Private mvarWeight As Double 'local copy
Private mvarHitPoints As Integer 'local copy
Private mvarDR As Variant 'local copy
Private mvarRuggedized As Boolean 'local copy
Private mvarMinTL As Integer 'local copy
Private mvarText As String 'local copy
Private mvarEmptySpace As Double
Private mvarAccessSpace As Double
Private mvarTopDeckArea As Double 'local copy
Private mvarFlightDeckArea As Double 'local copy
Private mvarCoveredDeckArea As Double 'local copy
Private mvarLandingPad As Boolean 'local copy
Private mvarAngledFlightDeck As Boolean 'local copy
'//////////////////////////////////////////////
'Overall Vehicle Options
'//////////////////////////////////////////////
Private mvarRollStabilizers As Boolean 'local copy
Private mvarFloatationHull As Boolean 'local copy
Private mvarLevitationCoating As Boolean 'local copy
Private mvarStreamLining As String 'local copy
Private mvarSubmersible As Boolean 'local copy
Private mvarSubmarineLines As Boolean 'local copy
Private mvarSealed As Boolean 'local copy
Private mvarWaterProof As Boolean 'local copy
Private mvarHydrodynamicLines As String 'local copy
Private mvarLiftingBody As Boolean 'local copy
Private mvarFlexibodyDrivetrain As Boolean 'local copy
Private mvarTrimaran As Boolean 'local copy
Private mvarCatamaran As Boolean 'local copy
Private mvarSoundBaffling As String 'local copy
Private mvarInfraredCloaking As String 'local copy
Private mvarCamouflage As Boolean 'local copy
Private mvarEmissionCloaking As String 'local copy
Private mvarStealth As String 'local copy
Private mvarChameleon As String 'local copy
Private mvarLiquidCrystal As Boolean 'local copy
Private mvarPsiShielding As Boolean 'local copy
Private mvarDeflector As Boolean 'local copy
Private mvarForceScreen As String 'local copy
Private mvarRam As Boolean 'local copy
Private mvarBulldozer As Boolean 'local copy
Private mvarPlow As Boolean 'local copy
Private mvarHitch As Boolean 'local copy
Private mvarPin As String 'local copy
Private mvarConvertible As String 'local copy
Private mvarForceDR As Long 'local copy
'/////////////////////////////////////////////
'Vehicle Performance
'/////////////////////////////////////////////
Private mvaraTopSpeed As Integer 'local copy
Private mvaraAcceleration As Integer 'local copy
Private mvaraDeceleration As Integer 'local copy
Private mvaraStability As Integer 'local copy
Private mvaraManeuverability As Integer 'local copy
Private mvaraDrag As Integer 'local copy
Private mvaraStallSpeed As Integer 'local copy
Private mvaraMotiveThrust As Integer 'local copy
Private mvargTopSpeed As Integer 'local copy
Private mvargAcceleration As Integer 'local copy
Private mvargDeceleration As Integer 'local copy
Private mvargStability As Integer 'local copy
Private mvargManeuverability As Integer 'local copy
Private mvargPressure As Integer 'local copy
Private mvargOffRoad As Integer 'local copy
Private mvarhTopSpeed As Integer 'local copy
Private mvarhAcceleration As Integer 'local copy
Private mvarhDeceleration As Integer 'local copy
Private mvarhStability As Integer 'local copy
Private mvarhManeuverability As Integer 'local copy
Private mvarmTopSpeed As Integer 'local copy
Private mvarmAcceleration As Integer 'local copy
Private mvarmDeceleration As Integer 'local copy
Private mvarmStability As Integer 'local copy
Private mvarmManeuverability As Integer 'local copy
Private mvarsTopSpeed As Integer 'local copy
Private mvarsAcceleration As Integer 'local copy
Private mvarsDeceleration As Integer 'local copy
Private mvarsStability As Integer 'local copy
Private mvarsManeuverability As Integer 'local copy
Private mvarsFTLSpeed As Integer 'local copy
Private mvarsuTopSpeed As Integer 'local copy
Private mvarsuAcceleration As Integer 'local copy
Private mvarsuDeceleration As Integer 'local copy
Private mvarsuStability As Integer 'local copy
Private mvarsuManeuverability As Integer 'local copy
Private mvarsuHydroDrag As Integer 'local copy
Private mvarsuCrushDepth As Integer 'local copy
Private mvarwTopSpeed As Integer 'local copy
Private mvarwAcceleration As Integer 'local copy
Private mvarwDeceleration As Integer 'local copy
Private mvarwStability As Integer 'local copy
Private mvarwManeuverability As Integer 'local copy
Private mvarwHydroDrag As Integer 'local copy
Private mvarwDraft As Integer 'local copy
'//////////////////////////////////////////////////////
' Total Vehicle Stats
Private mvarTotalPrice As Double 'local copy
Private mvarTotalSurfaceArea As Double 'local copy
Private mvarStructuralSurfaceArea As Double 'local copy
Private mvarStructuralCost As Double
Private mvarStructuralWeight As Double
Private mvarCombinedComponentVolume As Double 'local copy
Private mvarEmptyWeight As Double 'local copy
Private mvarHLoadedWeight As Double
Private mvarLoadedWeight As Double
Private mvarLoadedMass As Double
Private mvarHLoadedMass As Double
Private mvarTotalVolume As Double
Private mvarSubmergedWeight As Double
Private mvarSubmergedMass As Double
Private mvarHSubmergedWeight As Double ' with hardpoints loaded
Private mvarHSubmergedMass As Double
Private mvarUsualInternalPayload As Double
Private mvarStructuralHealth As Integer 'always with hardpoints loaded
Private mvarSizeModifier As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarStructuralTL As Integer 'local copy
Public Property Let StructuralTL(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StructuralTL = 5
    mvarStructuralTL = vData
End Property


Public Property Get StructuralTL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StructuralTL
    StructuralTL = mvarStructuralTL
End Property



Public Property Let SizeModifier(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SizeModifier = 5
    mvarSizeModifier = vData
End Property


Public Property Get SizeModifier() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SizeModifier
    SizeModifier = mvarSizeModifier
End Property



Public Property Let TotalPrice(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalPrice = 5
    mvarTotalPrice = vData
End Property


Public Property Get TotalPrice() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalPrice
    TotalPrice = mvarTotalPrice
End Property




Public Property Let StructuralHealth(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StructuralHealth = 5
    mvarStructuralHealth = vData
End Property

Public Property Get StructuralHealth() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StructuralHealth
    StructuralHealth = mvarStructuralHealth
End Property

Public Property Let HSubmergedWeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HSubmergedWeight = 5
    mvarHSubmergedWeight = vData
    HSubmergedMass = HSubmergedWeight / 2000 'todo round to 2decimal places
End Property

Public Property Get HSubmergedWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HSubmergedWeight
    HSubmergedWeight = mvarHSubmergedWeight
End Property
Public Property Let SubmergedWeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubmergedWeight = 5
    mvarSubmergedWeight = vData
    SubmergedMass = SubmergedWeight / 2000 ' todo round to 2decimal places
End Property

Public Property Get SubmergedWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubmergedWeight
    SubmergedWeight = mvarSubmergedWeight
End Property

Public Property Let HSubmergedMass(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HSubmergedMass = 5
    mvarHSubmergedMass = vData
End Property

Public Property Get HSubmergedMass() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HSubmergedMass
    HSubmergedMass = mvarHSubmergedMass
End Property

Public Property Let SubmergedMass(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubmergedMass = 5
    mvarSubmergedMass = vData
End Property

Public Property Get SubmergedMass() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubmergedMass
    SubmergedMass = mvarSubmergedMass
End Property

Public Property Let HLoadedWeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HLoadedWeight = 5
    mvarHLoadedWeight = vData
    HLoadedMass = HLoadedWeight / 2000 'todo round to 2decimal places
End Property

Public Property Get HLoadedWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HLoadedWeight
    HLoadedWeight = mvarHLoadedWeight
End Property

Public Property Let HLoadedMass(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HLoadedMass = 5
    mvarHLoadedMass = vData
End Property

Public Property Get HLoadedMass() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HLoadedMass
    HLoadedMass = mvarHLoadedMass
End Property

Public Property Let LoadedWeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LoadedWeight = 5
    mvarLoadedWeight = vData
    LoadedMass = LoadedWeight / 2000 'todo round to two decimal places
' Calculte the SubmergedWeight
Dim TempWeight As Double ' holds the value while calculations are made
Dim TempWeight2 As Double

TempWeight = 62.5 * mvarTotalVolume
TempWeight2 = TempWeight
' Submerged weight is the greater two values "LoadedWeight" and "TempWeight"
If mvarLoadedWeight > TempWeight Then
    TempWeight = mvarLoadedWeight
    TempWeight2 = mvarHLoadedWeight
End If
SubmergedWeight = TempWeight
HSubmergedWeight = TempWeight2
    
End Property

Public Property Get TotalVolume() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalVolume
    TotalVolume = mvarTotalVolume
End Property
Public Property Let TotalVolume(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalVolume = 5
    mvarTotalVolume = vData
End Property

Public Property Get LoadedWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LoadedWeight
    LoadedWeight = mvarLoadedWeight
End Property
Public Property Let LoadedMass(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LoadedMass = 5
    mvarLoadedMass = vData
End Property

Public Property Get LoadedMass() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LoadedMass
    LoadedMass = mvarLoadedMass
End Property
Public Property Let EmptySpace(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmptySpace = 5
    mvarEmptySpace = vData
End Property


Public Property Get EmptySpace() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmptySpace
    EmptySpace = mvarEmptySpace
End Property
Public Property Let AccessSpace(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccessSpace = 5
    mvarAccessSpace = vData
End Property


Public Property Get AccessSpace() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccessSpace
    AccessSpace = mvarAccessSpace
End Property

Public Property Let ForceDR(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForceDR = 5
    mvarForceDR = vData
End Property


Public Property Get ForceDR() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ForceDR
    ForceDR = mvarForceDR
End Property

Public Property Let AngledFlightDeck(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AngledFlightDeck = 5
    mvarAngledFlightDeck = vData
End Property


Public Property Get AngledFlightDeck() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AngledFlightDeck
    AngledFlightDeck = mvarAngledFlightDeck
End Property



Public Property Let LandingPad(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LandingPad = 5
    mvarLandingPad = vData
End Property


Public Property Get LandingPad() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LandingPad
    LandingPad = mvarLandingPad
End Property



Public Property Let CoveredDeckArea(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CoveredDeckArea = 5
    mvarCoveredDeckArea = vData
End Property


Public Property Get CoveredDeckArea() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CoveredDeckArea
    CoveredDeckArea = mvarCoveredDeckArea
End Property



Public Property Let FlightDeckArea(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FlightDeckArea = 5
    mvarFlightDeckArea = vData
End Property


Public Property Get FlightDeckArea() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FlightDeckArea
    FlightDeckArea = mvarFlightDeckArea
End Property



Public Property Let TopDeckArea(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TopDeckArea = 5
    mvarTopDeckArea = vData
End Property


Public Property Get TopDeckArea() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TopDeckArea
    TopDeckArea = mvarTopDeckArea
End Property



Public Property Let Convertible(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Convertible = 5
    mvarConvertible = vData
End Property


Public Property Get Convertible() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Convertible
    Convertible = mvarConvertible
End Property



Public Property Let Pin(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Pin = 5
    mvarPin = vData
End Property


Public Property Get Pin() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Pin
    Pin = mvarPin
End Property



Public Property Let Hitch(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Hitch = 5
    mvarHitch = vData
End Property


Public Property Get Hitch() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Hitch
    Hitch = mvarHitch
End Property



Public Property Let Plow(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Plow = 5
    mvarPlow = vData
End Property


Public Property Get Plow() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Plow
    Plow = mvarPlow
End Property



Public Property Let Bulldozer(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Bulldozer = 5
    mvarBulldozer = vData
End Property


Public Property Get Bulldozer() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Bulldozer
    Bulldozer = mvarBulldozer
End Property



Public Property Let Ram(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ram = 5
    mvarRam = vData
End Property


Public Property Get Ram() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ram
    Ram = mvarRam
End Property



Public Property Let ForceScreen(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForceScreen = 5
    mvarForceScreen = vData
End Property


Public Property Get ForceScreen() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ForceScreen
    ForceScreen = mvarForceScreen
End Property



Public Property Let Deflector(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Deflector = 5
    mvarDeflector = vData
End Property


Public Property Get Deflector() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Deflector
    Deflector = mvarDeflector
End Property



Public Property Let PsiShielding(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PsiShielding = 5
    mvarPsiShielding = vData
End Property


Public Property Get PsiShielding() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PsiShielding
    PsiShielding = mvarPsiShielding
End Property



Public Property Let LiquidCrystal(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LiquidCrystal = 5
    mvarLiquidCrystal = vData
End Property


Public Property Get LiquidCrystal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LiquidCrystal
    LiquidCrystal = mvarLiquidCrystal
End Property



Public Property Let Chameleon(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Chameleon = 5
    mvarChameleon = vData
End Property


Public Property Get Chameleon() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Chameleon
    Chameleon = mvarChameleon
End Property



Public Property Let Stealth(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Stealth = 5
    mvarStealth = vData
End Property


Public Property Get Stealth() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Stealth
    Stealth = mvarStealth
End Property



Public Property Let EmissionCloaking(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmissionCloaking = 5
    mvarEmissionCloaking = vData
End Property


Public Property Get EmissionCloaking() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmissionCloaking
    EmissionCloaking = mvarEmissionCloaking
End Property



Public Property Let Camouflage(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Camouflage = 5
    mvarCamouflage = vData
End Property


Public Property Get Camouflage() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Camouflage
    Camouflage = mvarCamouflage
End Property



Public Property Let InfraredCloaking(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InfraredCloaking = 5
    mvarInfraredCloaking = vData
End Property


Public Property Get InfraredCloaking() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InfraredCloaking
    InfraredCloaking = mvarInfraredCloaking
End Property



Public Property Let SoundBaffling(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SoundBaffling = 5
    mvarSoundBaffling = vData
End Property


Public Property Get SoundBaffling() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SoundBaffling
    SoundBaffling = mvarSoundBaffling
End Property





Public Property Let Text(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Text = 5
    mvarText = vData
End Property


Public Property Get Text() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Text
    Text = mvarText
End Property



Public Property Let MinTL(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MinTL = 5
    mvarMinTL = vData
End Property


Public Property Get MinTL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinTL
    MinTL = mvarMinTL
End Property



Public Property Let Ruggedized(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ruggedized = 5
    mvarRuggedized = vData
End Property


Public Property Get Ruggedized() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ruggedized
    Ruggedized = mvarRuggedized
End Property



Public Property Let DR(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR = 5
    mvarDR = vData
End Property


Public Property Set DR(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DR = Form1
    Set mvarDR = vData
End Property


Public Property Get DR() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DR
    If IsObject(mvarDR) Then
        Set DR = mvarDR
    Else
        DR = mvarDR
    End If
End Property



Public Property Let HitPoints(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HitPoints = 5
    mvarHitPoints = vData
End Property


Public Property Get HitPoints() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HitPoints
    HitPoints = mvarHitPoints
End Property

Public Property Let EmptyWeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmptyWeight = 5
    mvarEmptyWeight = vData
End Property


Public Property Get EmptyWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmptyWeight
    EmptyWeight = mvarEmptyWeight
End Property

Public Property Let UsualInternalPayload(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UsualInternalPayload = 5
    mvarUsualInternalPayload = vData
End Property


Public Property Get UsualInternalPayload() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UsualInternalPayload
    UsualInternalPayload = mvarUsualInternalPayload
End Property

Public Property Let StructuralSurfaceArea(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StructuralSurfaceArea = 5
    mvarStructuralSurfaceArea = vData
End Property


Public Property Get StructuralSurfaceArea() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StructuralSurfaceArea
    StructuralSurfaceArea = mvarStructuralSurfaceArea
End Property



Public Property Let TotalSurfaceArea(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalSurfaceArea = 5
    mvarTotalSurfaceArea = vData
End Property


Public Property Get TotalSurfaceArea() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalSurfaceArea
    TotalSurfaceArea = mvarTotalSurfaceArea
End Property

Public Property Let StructuralCost(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StructuralCost = 5
    mvarStructuralCost = vData
End Property


Public Property Get StructuralCost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StructuralCost
    StructuralCost = mvarStructuralCost
End Property
Public Property Let StructuralWeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StructuralWeight = 5
    mvarStructuralWeight = vData
End Property


Public Property Get StructuralWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StructuralWeight
    StructuralWeight = mvarStructuralWeight
End Property

Public Property Let Biomechanical(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Biomechanical = 5
    mvarBiomechanical = vData
End Property


Public Property Get Biomechanical() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Biomechanical
    Biomechanical = mvarBiomechanical
End Property



Public Property Let FrameStrength(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FrameStrength = 5
    mvarFrameStrength = vData
End Property


Public Property Get FrameStrength() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FrameStrength
    FrameStrength = mvarFrameStrength
End Property



Public Property Let LivingMetal(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LivingMetal = 5
    mvarLivingMetal = vData
End Property


Public Property Get LivingMetal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LivingMetal
    LivingMetal = mvarLivingMetal
End Property



Public Property Let Materials(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Materials = 5
    mvarMaterials = vData
End Property


Public Property Get Materials() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Materials
    Materials = mvarMaterials
End Property



Public Property Let Responsive(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Responsive = 5
    mvarResponsive = vData
End Property


Public Property Get Responsive() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Responsive
    Responsive = mvarResponsive
End Property



Public Property Let Robotic(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Robotic = 5
    mvarRobotic = vData
End Property


Public Property Get Robotic() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Robotic
    Robotic = mvarRobotic
End Property




Public Property Let component(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Component = 5
    mvarComponent = vData
End Property


Public Property Get component() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Component
    component = mvarComponent
End Property



Public Property Let Volume(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Volume = 5
    mvarVolume = vData
End Property


Public Property Get Volume() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Volume
    Volume = mvarVolume
End Property



Public Property Let CombinedComponentVolume(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CombinedComponentVolume = 5
    mvarCombinedComponentVolume = vData
End Property


Public Property Get CombinedComponentVolume() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CombinedComponentVolume
    CombinedComponentVolume = mvarCombinedComponentVolume
End Property
Public Property Let Weight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Weight = 5
    mvarWeight = vData
End Property


Public Property Get Weight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Weight
    Weight = mvarWeight
End Property



Public Property Let Cost(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cost = 5
    mvarCost = vData
End Property


Public Property Get Cost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cost
    Cost = mvarCost
End Property



Public Property Let RollStabilizers(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RollStabilizers = 5
    mvarRollStabilizers = vData
End Property


Public Property Get RollStabilizers() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RollStabilizers
    RollStabilizers = mvarRollStabilizers
End Property



Public Property Let ImprovedSuspension(ByVal vData As Boolean)
Attribute ImprovedSuspension.VB_Description = "Does the wheel sub assembly have improved suspension or not?"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ImprovedSuspension = 5
    mvarImprovedSuspension = vData
End Property


Public Property Get ImprovedSuspension() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ImprovedSuspension
    ImprovedSuspension = mvarImprovedSuspension
End Property



Public Property Let Compartmentalization(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Compartmentalization = 5
    mvarCompartmentalization = vData
End Property


Public Property Get Compartmentalization() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Compartmentalization
    Compartmentalization = mvarCompartmentalization
End Property



Public Property Let FloatationRating(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FloatationRating = 5
    mvarFloatationRating = vData
End Property


Public Property Get FloatationRating() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FloatationRating
    FloatationRating = mvarFloatationRating
End Property



Public Property Let FloatationHull(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FloatationHull = 5
    mvarFloatationHull = vData
End Property


Public Property Get FloatationHull() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FloatationHull
    FloatationHull = mvarFloatationHull
End Property

Public Property Let Sealed(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sealed = 5
    mvarSealed = vData
End Property


Public Property Get Sealed() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sealed
    Sealed = mvarSealed
End Property
Public Property Let WaterProof(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WaterProof = 5
    mvarWaterProof = vData
End Property


Public Property Get WaterProof() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WaterProof
    WaterProof = mvarWaterProof
End Property

Public Property Let MinimumVolume(ByVal vData As Double)
Attribute MinimumVolume.VB_Description = "Masts impose a certain minimum size on the vehicle.  Add cargo space or empty space to increase its size"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MinimumVolume = 5
    mvarMinimumVolume = vData
End Property


Public Property Get MinimumVolume() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinimumVolume
    MinimumVolume = mvarMinimumVolume
End Property

Public Property Let TotalRotationSpace(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalRotationSpace = 5
    mvarTotalRotationSpace = vData
End Property


Public Property Get TotalRotationSpace() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalRotationSpace
    TotalRotationSpace = mvarTotalRotationSpace
End Property

Public Property Let NumLegs(ByVal vData As Integer)
Attribute NumLegs.VB_Description = "Flag for how many legs are attached to the body.  This flag must be counted when legs are added and / or removed."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumLegs = 5
    mvarNumLegs = vData
End Property


Public Property Get NumLegs() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumLegs
    NumLegs = mvarNumLegs
End Property

Public Property Let NumWings(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumWings = 5
    mvarNumWings = vData
End Property


Public Property Get NumWings() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumWings
    NumWings = mvarNumWings
End Property

Public Property Let SubmarineLines(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubmarineLines = 5
    mvarSubmarineLines = vData
End Property


Public Property Get SubmarineLines() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubmarineLines
    SubmarineLines = mvarSubmarineLines
End Property

Public Property Let HydrodynamicLines(ByVal vData As String)
Attribute HydrodynamicLines.VB_Description = "A hydrodynamic hull is rated for its sleekness, classed as either Mediocre, Average, Fine or Very Fine lines.   Unless I add ""submerine lines"" i will need to add this as a new property."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HydrodynamicLines = 5
    mvarHydrodynamicLines = vData
End Property


Public Property Get HydrodynamicLines() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HydrodynamicLines
    HydrodynamicLines = mvarHydrodynamicLines
End Property



Public Property Let SlopeB(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlopeB = 5
    mvarSlopeB = vData
End Property


Public Property Get SlopeB() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlopeB
    SlopeB = mvarSlopeB
End Property



Public Property Let SlopeF(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlopeF = 5
    mvarSlopeF = vData
End Property


Public Property Get SlopeF() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlopeF
    SlopeF = mvarSlopeF
End Property



Public Property Let SlopeL(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlopeL = 5
    mvarSlopeL = vData
End Property


Public Property Get SlopeL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlopeL
    SlopeL = mvarSlopeL
End Property



Public Property Let SlopeR(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlopeR = 5
    mvarSlopeR = vData
End Property


Public Property Get SlopeR() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlopeR
    SlopeR = mvarSlopeR
End Property



Public Property Let RetractableWheelsSkids(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RetractableWheelsSkids = 5
    mvarRetractableWheelsSkids = vData
End Property


Public Property Get RetractableWheelsSkids() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RetractableWheelsSkids
    RetractableWheelsSkids = mvarRetractableWheelsSkids
End Property



Public Property Let Catamaran(ByVal vData As Boolean)
Attribute Catamaran.VB_Description = "Standard, Catamaran, or Trimaran (this effects volume for starters.. see pag 16 second column)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Catamaran = 5
    mvarCatamaran = vData
End Property


Public Property Get Catamaran() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Catamaran
    Catamaran = mvarCatamaran
End Property

Public Property Let Trimaran(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Trimaran = 5
    mvarTrimaran = vData
End Property


Public Property Get Trimaran() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Trimaran
    Trimaran = mvarTrimaran
End Property

Public Property Let SurfaceArea(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SurfaceArea = 5
    mvarSurfaceArea = vData
End Property


Public Property Get SurfaceArea() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SurfaceArea
    SurfaceArea = mvarSurfaceArea
End Property



Public Property Let FlexibodyDrivetrain(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FlexibodyDrivetrain = 5
    mvarFlexibodyDrivetrain = vData
End Property


Public Property Get FlexibodyDrivetrain() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FlexibodyDrivetrain
    FlexibodyDrivetrain = mvarFlexibodyDrivetrain
End Property



Public Property Let LiftingBody(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LiftingBody = 5
    mvarLiftingBody = vData
End Property


Public Property Get LiftingBody() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LiftingBody
    LiftingBody = mvarLiftingBody
End Property



Public Property Let StreamLining(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Streamlining = 5
    mvarStreamLining = vData
End Property


Public Property Get StreamLining() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Streamlining
    StreamLining = mvarStreamLining
End Property



Public Property Let WingsorRotors(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WingsorRotors = 5
    mvarWingsorRotors = vData
End Property


Public Property Get WingsorRotors() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WingsorRotors
    WingsorRotors = mvarWingsorRotors
End Property



Public Property Let Submersible(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Submersible = 5
    mvarSubmersible = vData
End Property


Public Property Get Submersible() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Submersible
    Submersible = mvarSubmersible
End Property



Public Property Let TL(ByVal vData As Integer)
Attribute TL.VB_Description = "Is the age that the vehicle was built.  See page 6 Vehicles bottom first column"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TL = 5
    mvarTL = vData
    gVehicleTL = mvarTL ' update the Global variable
    
End Property


Public Property Get TL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    TL = mvarTL
End Property


Public Property Let wDraft(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wDraft = 5
    mvarwDraft = vData
End Property


Public Property Get wDraft() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wDraft
    wDraft = mvarwDraft
End Property



Public Property Let wHydroDrag(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wHydroDrag = 5
    mvarwHydroDrag = vData
End Property


Public Property Get wHydroDrag() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wHydroDrag
    wHydroDrag = mvarwHydroDrag
End Property



Public Property Let wManeuverability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wManeuverability = 5
    mvarwManeuverability = vData
End Property


Public Property Get wManeuverability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wManeuverability
    wManeuverability = mvarwManeuverability
End Property



Public Property Let wStability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wStability = 5
    mvarwStability = vData
End Property


Public Property Get wStability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wStability
    wStability = mvarwStability
End Property



Public Property Let wDeceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wDeceleration = 5
    mvarwDeceleration = vData
End Property


Public Property Get wDeceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wDeceleration
    wDeceleration = mvarwDeceleration
End Property



Public Property Let wAcceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wAcceleration = 5
    mvarwAcceleration = vData
End Property


Public Property Get wAcceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wAcceleration
    wAcceleration = mvarwAcceleration
End Property



Public Property Let wTopSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.wTopSpeed = 5
    mvarwTopSpeed = vData
End Property


Public Property Get wTopSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.wTopSpeed
    wTopSpeed = mvarwTopSpeed
End Property



Public Property Let suCrushDepth(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suCrushDepth = 5
    mvarsuCrushDepth = vData
End Property


Public Property Get suCrushDepth() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suCrushDepth
    suCrushDepth = mvarsuCrushDepth
End Property



Public Property Let suHydroDrag(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suHydroDrag = 5
    mvarsuHydroDrag = vData
End Property


Public Property Get suHydroDrag() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suHydroDrag
    suHydroDrag = mvarsuHydroDrag
End Property



Public Property Let suManeuverability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suManeuverability = 5
    mvarsuManeuverability = vData
End Property


Public Property Get suManeuverability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suManeuverability
    suManeuverability = mvarsuManeuverability
End Property



Public Property Let suStability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suStability = 5
    mvarsuStability = vData
End Property


Public Property Get suStability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suStability
    suStability = mvarsuStability
End Property



Public Property Let suDeceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suDeceleration = 5
    mvarsuDeceleration = vData
End Property


Public Property Get suDeceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suDeceleration
    suDeceleration = mvarsuDeceleration
End Property



Public Property Let suAcceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suAcceleration = 5
    mvarsuAcceleration = vData
End Property


Public Property Get suAcceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suAcceleration
    suAcceleration = mvarsuAcceleration
End Property



Public Property Let suTopSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.suTopSpeed = 5
    mvarsuTopSpeed = vData
End Property


Public Property Get suTopSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.suTopSpeed
    suTopSpeed = mvarsuTopSpeed
End Property



Public Property Let sFTLSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFTLSpeed = 5
    mvarsFTLSpeed = vData
End Property


Public Property Get sFTLSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFTLSpeed
    sFTLSpeed = mvarsFTLSpeed
End Property



Public Property Let sManeuverability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sManeuverability = 5
    mvarsManeuverability = vData
End Property


Public Property Get sManeuverability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sManeuverability
    sManeuverability = mvarsManeuverability
End Property



Public Property Let sStability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sStability = 5
    mvarsStability = vData
End Property


Public Property Get sStability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sStability
    sStability = mvarsStability
End Property



Public Property Let sDeceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDeceleration = 5
    mvarsDeceleration = vData
End Property


Public Property Get sDeceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDeceleration
    sDeceleration = mvarsDeceleration
End Property



Public Property Let sAcceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sAcceleration = 5
    mvarsAcceleration = vData
End Property


Public Property Get sAcceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sAcceleration
    sAcceleration = mvarsAcceleration
End Property



Public Property Let sTopSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTopSpeed = 5
    mvarsTopSpeed = vData
End Property


Public Property Get sTopSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTopSpeed
    sTopSpeed = mvarsTopSpeed
End Property



Public Property Let mManeuverability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mManeuverability = 5
    mvarmManeuverability = vData
End Property


Public Property Get mManeuverability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mManeuverability
    mManeuverability = mvarmManeuverability
End Property



Public Property Let mStability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mStability = 5
    mvarmStability = vData
End Property


Public Property Get mStability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mStability
    mStability = mvarmStability
End Property



Public Property Let mDeceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mDeceleration = 5
    mvarmDeceleration = vData
End Property


Public Property Get mDeceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mDeceleration
    mDeceleration = mvarmDeceleration
End Property



Public Property Let mAcceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mAcceleration = 5
    mvarmAcceleration = vData
End Property


Public Property Get mAcceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mAcceleration
    mAcceleration = mvarmAcceleration
End Property



Public Property Let mTopSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mTopSpeed = 5
    mvarmTopSpeed = vData
End Property


Public Property Get mTopSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mTopSpeed
    mTopSpeed = mvarmTopSpeed
End Property



Public Property Let hManeuverability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hManeuverability = 5
    mvarhManeuverability = vData
End Property


Public Property Get hManeuverability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hManeuverability
    hManeuverability = mvarhManeuverability
End Property



Public Property Let hStability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hStability = 5
    mvarhStability = vData
End Property


Public Property Get hStability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hStability
    hStability = mvarhStability
End Property



Public Property Let hDeceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hDeceleration = 5
    mvarhDeceleration = vData
End Property


Public Property Get hDeceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hDeceleration
    hDeceleration = mvarhDeceleration
End Property



Public Property Let hAcceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hAcceleration = 5
    mvarhAcceleration = vData
End Property


Public Property Get hAcceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hAcceleration
    hAcceleration = mvarhAcceleration
End Property



Public Property Let hTopSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hTopSpeed = 5
    mvarhTopSpeed = vData
End Property


Public Property Get hTopSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hTopSpeed
    hTopSpeed = mvarhTopSpeed
End Property

Public Property Let gOffRoad(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gOffRoad = 5
    mvargOffRoad = vData
End Property


Public Property Get gOffRoad() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gOffRoad
    gOffRoad = mvargOffRoad
End Property



Public Property Let gPressure(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gPressure = 5
    mvargPressure = vData
End Property


Public Property Get gPressure() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gPressure
    gPressure = mvargPressure
End Property



Public Property Let gManeuverability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gManeuverability = 5
    mvargManeuverability = vData
End Property


Public Property Get gManeuverability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gManeuverability
    gManeuverability = mvargManeuverability
End Property



Public Property Let gStability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gStability = 5
    mvargStability = vData
End Property


Public Property Get gStability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gStability
    gStability = mvargStability
End Property



Public Property Let gDeceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gDeceleration = 5
    mvargDeceleration = vData
End Property


Public Property Get gDeceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gDeceleration
    gDeceleration = mvargDeceleration
End Property



Public Property Let gAcceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gAcceleration = 5
    mvargAcceleration = vData
End Property


Public Property Get gAcceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gAcceleration
    gAcceleration = mvargAcceleration
End Property



Public Property Let gTopSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gTopSpeed = 5
    mvargTopSpeed = vData
End Property


Public Property Get gTopSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gTopSpeed
    gTopSpeed = mvargTopSpeed
End Property



Public Property Let aMotiveThrust(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aMotiveThrust = 5
    mvaraMotiveThrust = vData
End Property


Public Property Get aMotiveThrust() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aMotiveThrust
    aMotiveThrust = mvaraMotiveThrust
End Property



Public Property Let aStallSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aStallSpeed = 5
    mvaraStallSpeed = vData
End Property


Public Property Get aStallSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aStallSpeed
    aStallSpeed = mvaraStallSpeed
End Property



Public Property Let aDrag(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aDrag = 5
    mvaraDrag = vData
End Property


Public Property Get aDrag() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aDrag
    aDrag = mvaraDrag
End Property



Public Property Let aManeuverability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aManeuverability = 5
    mvaraManeuverability = vData
End Property


Public Property Get aManeuverability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aManeuverability
    aManeuverability = mvaraManeuverability
End Property



Public Property Let aStability(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aStability = 5
    mvaraStability = vData
End Property


Public Property Get aStability() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aStability
    aStability = mvaraStability
End Property



Public Property Let aDeceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aDeceleration = 5
    mvaraDeceleration = vData
End Property


Public Property Get aDeceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aDeceleration
    aDeceleration = mvaraDeceleration
End Property



Public Property Let aAcceleration(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aAcceleration = 5
    mvaraAcceleration = vData
End Property


Public Property Get aAcceleration() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aAcceleration
    aAcceleration = mvaraAcceleration
End Property



Public Property Let aTopSpeed(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aTopSpeed = 5
    mvaraTopSpeed = vData
End Property


Public Property Get aTopSpeed() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aTopSpeed
    aTopSpeed = mvaraTopSpeed
End Property



Public Property Let LevitationCoating(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LevitationCoating = 5
    mvarLevitationCoating = vData
End Property


Public Property Get LevitationCoating() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LevitationCoating
    LevitationCoating = mvarLevitationCoating
End Property



Public Property Let Key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property



Public Property Let Parent(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Parent = 5
    mvarParent = vData
End Property


Public Property Get Parent() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parent
    Parent = mvarParent
End Property



Public Property Let SelectedImage(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SelectedImage = 5
    mvarSelectedImage = vData
End Property


Public Property Get SelectedImage() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedImage
    SelectedImage = mvarSelectedImage
End Property



Public Property Let Image(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Image = 5
    mvarImage = vData
End Property


Public Property Get Image() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Image
    Image = mvarImage
End Property

Public Property Get ClassDebugID() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClassDebugID
    If IsObject(mvarClassDebugID) Then
        Set ClassDebugID = mvarClassDebugID
    Else
        ClassDebugID = mvarClassDebugID
    End If
End Property

Private Sub Class_Initialize()
' set the default values
mvarLevitationCoating = False
mvaraTopSpeed = 0
mvaraAcceleration = 0
mvaraDeceleration = 0
mvaraStability = 0
mvaraManeuverability = 0
mvaraDrag = 0
mvaraStallSpeed = 0
mvaraMotiveThrust = 0
mvargTopSpeed = 0
mvargAcceleration = 0
mvargDeceleration = 0
mvargStability = 0
mvargManeuverability = 0
mvargPressure = 0
mvargOffRoad = 0
mvarhTopSpeed = 0
mvarhAcceleration = 0
mvarhDeceleration = 0
mvarhStability = 0
mvarhManeuverability = 0
mvarmTopSpeed = 0
mvarmAcceleration = 0
mvarmDeceleration = 0
mvarmStability = 0
mvarmManeuverability = 0
mvarsTopSpeed = 0
mvarsAcceleration = 0
mvarsDeceleration = 0
mvarsStability = 0
mvarsManeuverability = 0
mvarsFTLSpeed = 0
mvarsuTopSpeed = 0
mvarsuAcceleration = 0
mvarsuDeceleration = 0
mvarsuStability = 0
mvarsuManeuverability = 0
mvarsuHydroDrag = 0
mvarsuCrushDepth = 0
mvarwTopSpeed = 0
mvarwAcceleration = 0
mvarwDeceleration = 0
mvarwStability = 0
mvarwManeuverability = 0
mvarwHydroDrag = 0
mvarwDraft = 0
mvarTL = 0
mvarSubmersible = False
mvarSubmarineLines = False
mvarHydrodynamicLines = "None"
mvarCatamaran = False
mvarTrimaran = False
mvarStreamLining = "None"
mvarFloatationHull = False
mvarFloatationRating = 62.5
mvarCompartmentalization = "None"
mvarRollStabilizers = False
mvarWingsorRotors = False
mvarLiftingBody = False

mvarSoundBaffling = "None"
mvarInfraredCloaking = "None"
mvarCamouflage = False
mvarEmissionCloaking = "None"
mvarStealth = "None"
mvarChameleon = "None"
mvarLiquidCrystal = False
mvarPsiShielding = False
mvarDeflector = False
mvarForceScreen = "None"
mvarRam = False
mvarBulldozer = False
mvarPlow = False
mvarHitch = False
mvarPin = "None"
mvarConvertible = "None"
mvarTopDeckArea = 0
mvarFlightDeckArea = 0
mvarCoveredDeckArea = 0
mvarLandingPad = False
mvarAngledFlightDeck = False

mvarFlexibodyDrivetrain = True
mvarImprovedSuspension = False
mvarRetractableWheelsSkids = "None"
mvarSlopeR = "None"
mvarSlopeL = "None"
mvarSlopeF = "None"
mvarSlopeB = "None"
mvarNumLegs = 0
mvarSurfaceArea = 0
mvarTotalRotationSpace = 0
mvarMinimumVolume = 0
mvarVolume = 0
mvarCost = 0
mvarWeight = 0
mvarHitPoints = 0
mvarRobotic = False
mvarResponsive = False
mvarMaterials = "Standard"
mvarLivingMetal = False
mvarFrameStrength = "Medium"
mvarBiomechanical = False
mvarTotalSurfaceArea = 0
mvarStructuralSurfaceArea = 0
mvarCombinedComponentVolume = 0
mvarAccessSpace = 0
mvarEmptySpace = 0
mvarEmptyWeight = 0
mvarUsualInternalPayload = 0

    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
End Sub


Private Sub Class_Terminate()
    'the class is being destroyed
    #If DebugMode Then
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

Public Function PropQuery(ByVal Row As Integer) As Integer
Const Combobox = 1
Const Textbox = 2
Const CommandButton = 3
' clear the Properties form's combobox
frmProperties.Combo1.Clear
    Select Case Row
    Case 8 ' Tech Level
        frmProperties.Combo1.AddItem 0
        frmProperties.Combo1.AddItem 1
        frmProperties.Combo1.AddItem 2
        frmProperties.Combo1.AddItem 3
        frmProperties.Combo1.AddItem 4
        frmProperties.Combo1.AddItem 5
        frmProperties.Combo1.AddItem 6
        frmProperties.Combo1.AddItem 7
        frmProperties.Combo1.AddItem 8
        frmProperties.Combo1.AddItem 9
        frmProperties.Combo1.AddItem 10
        frmProperties.Combo1.AddItem 11
        frmProperties.Combo1.AddItem 12
        frmProperties.Combo1.AddItem 13
        frmProperties.Combo1.AddItem 14
        frmProperties.Combo1.AddItem 15
        frmProperties.Combo1.AddItem 16
        PropQuery = Combobox
    Case 5 To 7, 9 To 18, 34
        PropQuery = Textbox
    Case 26
        frmProperties.Combo1.AddItem "None"
        frmProperties.Combo1.AddItem "Mediocre"
        frmProperties.Combo1.AddItem "Average"
        frmProperties.Combo1.AddItem "Fine"
        frmProperties.Combo1.AddItem "Very Fine"
        frmProperties.Combo1.AddItem "Submarine"
        PropQuery = Combobox
    Case 27
        frmProperties.Combo1.AddItem "None"
        frmProperties.Combo1.AddItem "Catamaran"
        frmProperties.Combo1.AddItem "Trimaran"
        PropQuery = Combobox
    Case 28
        frmProperties.Combo1.AddItem "None"
        frmProperties.Combo1.AddItem "Fair"
        frmProperties.Combo1.AddItem "Good"
        frmProperties.Combo1.AddItem "Very Good"
        frmProperties.Combo1.AddItem "Superior"
        frmProperties.Combo1.AddItem "Excellent"
        frmProperties.Combo1.AddItem "Radical"
        PropQuery = Combobox
    Case 30
        frmProperties.Combo1.AddItem "None"
        frmProperties.Combo1.AddItem "Heavy"
        frmProperties.Combo1.AddItem "Total"
        PropQuery = Combobox
    Case 0 To 4
        frmProperties.Combo1.AddItem "None"
        frmProperties.Combo1.AddItem "30 degrees"
        frmProperties.Combo1.AddItem "60 degrees"
        PropQuery = Combobox
    Case 21 To 25, 29, 31, 32, 33
        frmProperties.Combo1.AddItem "True"
        frmProperties.Combo1.AddItem "False"
        PropQuery = Combobox
    Case 20
        frmProperties.Combo1.AddItem "Very Cheap"
        frmProperties.Combo1.AddItem "Cheap"
        frmProperties.Combo1.AddItem "Standard"
        frmProperties.Combo1.AddItem "Expensive"
        frmProperties.Combo1.AddItem "Very Expensive"
        frmProperties.Combo1.AddItem "Advanced"
        PropQuery = Combobox
    Case 19
        frmProperties.Combo1.AddItem "Super-Light"
        frmProperties.Combo1.AddItem "Extra-Light"
        frmProperties.Combo1.AddItem "Light"
        frmProperties.Combo1.AddItem "Medium"
        frmProperties.Combo1.AddItem "Heavy"
        frmProperties.Combo1.AddItem "Extra-Heavy"
        PropQuery = Combobox
    End Select
End Function

Public Sub PropUpdate(iRow As Integer, vNewValue As Variant)
Select Case iRow
    Case 0
        SlopeR = vNewValue
    Case 1
        SlopeL = vNewValue
    Case 2
        SlopeF = vNewValue
    Case 3
        SlopeB = vNewValue
    Case 4
        ' just a textbox for numlegs
    Case 5
        ' just a textbox for wings or rotars
    Case 6
        ' just a textbox for improved suspension
    Case 7
        ' just a textbox for retractable wheelsskids
    Case 8
        TL = vNewValue
    Case 9
        ' textbox for Cost
    Case 10
        ' textbox for weight
    Case 11
        'textbox for Volume
    Case 12
        ' just a textbox fo Surface Area
    Case 13
        ' textbox for Total rotation space
    Case 14
        ' textbox for minimumvolume
    Case 15
        ' textbox for Hitpoints
    Case 16
        ' just a textbox for floatation rating
    Case 17
        'textbox for TotalSurfaceArea
    Case 18
        ' textbox for StructuralSurfaceArea
    Case 19
        FrameStrength = vNewValue
    Case 20
        Materials = vNewValue
    Case 21
        Responsive = vNewValue
    Case 22
        Robotic = vNewValue
    Case 23
        Biomechanical = vNewValue
    Case 24
        LivingMetal = vNewValue
    Case 25
        Submersible = vNewValue
    Case 26
        HydrodynamicLines = vNewValue
    Case 27
        Catamaran = vNewValue
    Case 28
        StreamLining = vNewValue
    Case 29
        FloatationHull = vNewValue
    Case 30
        Compartmentalization = vNewValue
    Case 31
        RollStabilizers = vNewValue
    Case 32
        LiftingBody = vNewValue
    Case 33
        FlexibodyDrivetrain = vNewValue
    Case 34
        CombinedComponentVolume = vNewValue
End Select
StatsUpdate
End Sub


Public Sub StatsUpdate()
Dim Floatation As Double ' the floatation multiplier
' calculate the real body volume
mvarVolume = CalcBodyVolume
    
' Calculate the floatation rating (must be done after volume)
If mvarSubmersible = True Then
    Floatation = 62.5
Else
    Select Case Me.HydrodynamicLines
        Case "None"
            Floatation = 62.5
        Case "Mediocre"
            Floatation = 57
        Case "Average"
            Floatation = 52
        Case "Submarine"
            Floatation = 62.5
        Case "Fine"
            Floatation = 48
        Case "Very Fine"
            Floatation = 45
    End Select
End If
mvarFloatationRating = Floatation * mvarVolume
' calculate the surface area
mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
' Calculate the structural cost
mvarCost = BasicDesignCost(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials, mvarResponsive, mvarRobotic, mvarBiomechanical, mvarLivingMetal)
' Calculate the structural weight
mvarWeight = BasicDesignWeight(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials)
' Calculate the Hit Points
mvarHitPoints = CalcHitPoints(TypeName(Me), mvarFrameStrength, mvarSurfaceArea)
' Calculate Special Structural Option modifiers
   ' Compartmentalization
        ' Improved Suspension
        ' Wing and Rotar Options
        ' Wheel Options
        ' Roll Stabilizers

End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    colVehicle.Item(Parent).StatsUpdate
End If

End Sub

Public Sub QueryChild()
' if the object has children, query it and check to see if
' more stats/property updates are needed for other objects in the collection
' (see the StatusUpdate property for help on checking for childeren.  Can i use that one in place of this?)

End Sub
