VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWeaponTorpMissile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarHitPoints As Double
Private mvarDR As Long
Private mvarCustom As Boolean
Private mvarCost As Double
Private mvarDatatype As Integer
Private mvarDescription As String
Private mvarImage As Integer
Private mvarSelectedImage As Integer
Private mvarKey As String
Private mvarParent As String
Private mvarQuantity As Integer
Private mvarSurfaceArea As Double
Private mvarTL As Integer
Private mvarVolume As Double
Private mvarWeight As Double
Private mvarRuggedized As Boolean
Private mvarCustomDescription As String
Private mvarGuidanceSystem As String
Private mvarBrilliantGuidanceSystem As String
Private mvarCheapGuidance As Boolean
Private mvarDiameter As Single
Private mvarWarHead As String
Private mvarBrilliant As Boolean
Private mvarMidCourseUpdate As Boolean
Private mvarPopUp As Boolean
Private mvarSkill As String
Private mvarSkillBonus As Integer
Private mvarWarheadWeight As Single
Private mvarWarheadCost As Single
Private mvarGuidanceWeight As Single
Private mvarGuidanceCost As Single
Private mvarPayloadWeight As Single
Private mvarPayloadCost As Single
Private mvarSpeed As Single
Private mvarMotorWeight As Single
Private mvarMotorCost As Single
Private mvarStealth As Boolean
Private mvarMalfunction As String
Private mvarTypeDamage1 As String
Private mvarDamage1 As String
Private mvarDamage2 As String
Private mvarTypeDamage2 As String
Private mvarBurstRadius As Long
Private mvarEndurance As Single
Private mvarAccelG As Single
Private mvarAccelMPH As Single
Private mvarMaxRange As Double
Private mvarhalfDamage As Double
Private mvarMinRange As Single
Private mvarAccuracy As Long
Private mvarWarheadSize As String
Private mvarBusMissiles As Long
Private mvarSpaceMissile As Boolean
Private mvarKEDamage As Double
Private mvarCompact As Boolean
Private mvarDetonationWeight As Single
Private mvarParachute As Boolean


Private mvarLocation As String
Private mvarComment As String
Private mvarCName As String
Private mvarMatrixPos As Long
Private mvarMatrixPos2 As Long
Private mvarPrintOutput As String
Private mvarZZInit As Byte
Private mvarLogicalParent As String

Public Property Let LogicalParent(ByVal vdata As String)
    mvarLogicalParent = vdata
End Property
Public Property Get LogicalParent() As String
    LogicalParent = mvarLogicalParent
End Property
Public Property Let PrintOutput(ByVal vdata As String)
    mvarPrintOutput = vdata
End Property
Public Property Get PrintOutput() As String
    PrintOutput = mvarPrintOutput
End Property
Public Property Let MatrixPos(ByVal vdata As Long)
    mvarMatrixPos = vdata
End Property
Public Property Get MatrixPos() As Long
    MatrixPos = mvarMatrixPos
End Property
Public Property Let CName(ByVal vdata As String)
    mvarCName = vdata
End Property
Public Property Get CName() As String
    CName = mvarCName
End Property
Public Property Let Comment(ByVal vdata As String)
    mvarComment = vdata
End Property
Public Property Get Comment() As String
    Comment = mvarComment
End Property
Public Property Let Location(ByVal vdata As String)
    mvarLocation = vdata
End Property
Public Property Get Location() As String
    Location = mvarLocation
End Property

Public Function LocationCheck() As Boolean
Dim TempCheck As Boolean
Dim InstallPoint As Integer
InstallPoint = Veh.Components(mvarParent).Datatype


Select Case mvarDatatype
    Case UnGuidedMissile, UnGuidedTorpedo, GuidedMissile, GuidedTorpedo
        If (InstallPoint = HardPoint) Or (InstallPoint = WeaponBay) Or (InstallPoint = DisposableLauncher) Or (InstallPoint = MuzzleloadingLauncher) Or (InstallPoint = BreechloadingLauncher) Or (InstallPoint = ManualRepeaterLauncher) Or (InstallPoint = SlowAutoLoaderLauncher) Or (InstallPoint = FastAutoLoaderLauncher) Or (InstallPoint = RevolverLauncher) Or (InstallPoint = lightAutomaticLauncher) Or (InstallPoint = HeavyAutomaticLauncher) Then
            TempCheck = True
        Else
            InfoPrint 1, "Missiles and Torpedos must be placed on a Launcher, Hardpoint or Weapon Bay"
            TempCheck = False
        End If
    Case IronBomb, RetardedBomb, SmartBomb
        If (InstallPoint = HardPoint) Or (InstallPoint = WeaponBay) Then
            TempCheck = True
        Else
            InfoPrint 1, "Bombs must be placed on a Hardpoint or Weapon Bay."
            TempCheck = False
        End If
    Case ContactMine, ProximityMine, PressureTriggerMine, CommandTriggerMine, SmartTriggerMine

        If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = equipmentPod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = DisposableLauncher) Or (InstallPoint = MuzzleloadingLauncher) Or (InstallPoint = BreechloadingLauncher) Or (InstallPoint = ManualRepeaterLauncher) Or (InstallPoint = SlowAutoLoaderLauncher) Or (InstallPoint = FastAutoLoaderLauncher) Or (InstallPoint = RevolverLauncher) Or (InstallPoint = lightAutomaticLauncher) Or (InstallPoint = HeavyAutomaticLauncher) Then
            TempCheck = True
        Else
            InfoPrint 1, "Mines must be placed in Body, Superstructure, Pod, equipment Pod,Turret, Popturret, Arm, Wing, Open Mount, Leg or Launcher."
            TempCheck = False
        End If

    Case SelfDestructSystem
        If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = equipmentPod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Then
            TempCheck = True
        Else
            InfoPrint 1, "Self Destruct system must be placed in Body, Superstructure, Pod, equipment Pod,Turret, Popturret, Arm, Wing, Open Mount or Leg."
            TempCheck = False
        End If
        
End Select

If TempCheck Then SetLogicalParent
LocationCheck = TempCheck
End Function

Private Function GetLocation() As String
    On Error Resume Next
    If mvarLogicalParent = "" Then SetLogicalParent
    GetLocation = Veh.Components(mvarLogicalParent).Abbrev
End Function

Public Sub SetLogicalParent()
    mvarLogicalParent = GetLogicalParent(mvarParent)
End Sub

Public Property Let KEDamage(ByVal vdata As Double)
    mvarKEDamage = vdata
End Property
Public Property Get KEDamage() As Double
    KEDamage = mvarKEDamage
End Property
Public Property Let DetonationWeight(ByVal vdata As Double)
    mvarDetonationWeight = vdata
End Property
Public Property Get DetonationWeight() As Double
    DetonationWeight = mvarDetonationWeight
End Property
Public Property Let GuidanceSystem(ByVal vdata As String)
    mvarGuidanceSystem = vdata
        Dim i As Integer
    If mvarZZInit = 0 Then Exit Property
    GetMatrixIndex
End Property
Public Property Get GuidanceSystem() As String
    GuidanceSystem = mvarGuidanceSystem
End Property
Public Property Let BrilliantGuidanceSystem(ByVal vdata As String)
    mvarBrilliantGuidanceSystem = vdata
        Dim i As Integer
    
    If mvarZZInit = 0 Then Exit Property
    GetMatrixIndex
End Property
Public Property Get BrilliantGuidanceSystem() As String
    BrilliantGuidanceSystem = mvarBrilliantGuidanceSystem
End Property
Public Property Let BusMissiles(ByVal vdata As Long)
    mvarBusMissiles = vdata
End Property
Public Property Get BusMissiles() As Long
    BusMissiles = mvarBusMissiles
End Property
Public Property Let Parachute(ByVal vdata As Boolean)
    mvarParachute = vdata
End Property
Public Property Get Parachute() As Boolean
    Parachute = mvarParachute
End Property
Public Property Let SpaceMissile(ByVal vdata As Boolean)
    mvarSpaceMissile = vdata
End Property
Public Property Get SpaceMissile() As Boolean
    SpaceMissile = mvarSpaceMissile
End Property
Public Property Let WarheadSize(ByVal vdata As String)
    mvarWarheadSize = vdata
End Property
Public Property Get WarheadSize() As String
    WarheadSize = mvarWarheadSize
End Property
Public Property Let Accuracy(ByVal vdata As Long)
    mvarAccuracy = vdata
End Property
Public Property Get Accuracy() As Long
    Accuracy = mvarAccuracy
End Property
Public Property Let MinRange(ByVal vdata As Single)
    mvarMinRange = vdata
End Property
Public Property Get MinRange() As Single
    MinRange = mvarMinRange
End Property
Public Property Let halfDamage(ByVal vdata As Double)
    mvarhalfDamage = vdata
End Property
Public Property Get halfDamage() As Double
    halfDamage = mvarhalfDamage
End Property
Public Property Let MaxRange(ByVal vdata As Double)
    mvarMaxRange = vdata
End Property
Public Property Get MaxRange() As Double
    MaxRange = mvarMaxRange
End Property
Public Property Let Endurance(ByVal vdata As Single)
    mvarEndurance = vdata
End Property
Public Property Get Endurance() As Single
    Endurance = mvarEndurance
End Property
Public Property Let Damage1(ByVal vdata As String)
    mvarDamage1 = vdata
End Property
Public Property Get Damage1() As String
    Damage1 = mvarDamage1
End Property
Public Property Let Damage2(ByVal vdata As String)
    mvarDamage2 = vdata
End Property
Public Property Get Damage2() As String
    Damage2 = mvarDamage2
End Property
Public Property Let TypeDamage1(ByVal vdata As String)
    mvarTypeDamage1 = vdata
End Property
Public Property Get TypeDamage1() As String
    TypeDamage1 = mvarTypeDamage1
End Property
Public Property Let TypeDamage2(ByVal vdata As String)
    mvarTypeDamage2 = vdata
End Property
Public Property Get TypeDamage2() As String
    TypeDamage2 = mvarTypeDamage2
End Property
Public Property Let BurstRadius(ByVal vdata As Long)
    mvarBurstRadius = vdata
End Property
Public Property Get BurstRadius() As Long
    BurstRadius = mvarBurstRadius
End Property
Public Property Let Malfunction(ByVal vdata As String)
    mvarMalfunction = vdata
End Property
Public Property Get Malfunction() As String
    Malfunction = mvarMalfunction
End Property

Public Property Let Stealth(ByVal vdata As Boolean)
    mvarStealth = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If mvarTL < 7 Then
        mvarStealth = False
        InfoPrint 1, "Stealth technology is not available until Tech Level 7"
    End If
End Property

Public Property Get Stealth() As Boolean
    Stealth = mvarStealth
End Property
Public Property Let MotorCost(ByVal vdata As Single)
    mvarMotorCost = vdata
End Property
Public Property Get MotorCost() As Single
    MotorCost = mvarMotorCost
End Property

Public Property Let MotorWeight(ByVal vdata As Double)
    'todo:  THIS IS A BUG.  vdata is a Double, yet mvarMotorWeight is a single.  THis
    '       causes precision issues when the value is converted into a single.
    '       NEED TO CHECK ALL OF GVD FOR "single" values and change them to "DOUBLES" where appropriate.
    
    
    mvarMotorWeight = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If mvarMotorWeight < 0.05 Then
        mvarMotorWeight = 0.05
        InfoPrint 1, "Minimum motor weight is .05 lbs"
    End If
    
    If mvarSpaceMissile = False Then
        If mvarMotorWeight > (mvarPayloadWeight * 11) Then
            mvarMotorWeight = mvarPayloadWeight * 11
            InfoPrint 1, "Max motor weight for NON space missiles is Payload Weight x 11"
        End If
    Else
        'spacemissiles have no upper limit
    End If
End Property

Public Property Get MotorWeight() As Double
    MotorWeight = mvarMotorWeight
End Property
Public Property Let Speed(ByVal vdata As Single)
    mvarSpeed = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If mvarSpeed < 5 Then
        mvarSpeed = 5
        InfoPrint 1, "Minimum speed for missiles and torpedos is 5 yards per second"
    End If
    
    'now check for max values
    If mvarSpaceMissile Then
        'space missiles have no max speed
    Else
        Select Case mvarDatatype
        Case UnGuidedTorpedo
            If mvarSpeed > mvarTL * 10 Then
                mvarSpeed = mvarTL * 10
                InfoPrint 1, "Max speed for NON Wire or Sonar Guided Torpedos is TL x 10"
            End If
        Case GuidedTorpedo
            Select Case mvarGuidanceSystem
                Case "WG", "PSH", "ASH", "WGSH"
                    If (mvarSpeed > mvarTL * 10) Or (mvarSpeed > 300) Then
                        mvarSpeed = mvarTL * 10
                        If mvarSpeed > 300 Then mvarSpeed = 300
                        InfoPrint 1, "Max speed for Wire or Sonar Guided Torpedos is TL x 10 or 300 yards per second. Whichever is LESS."
                    End If
            Case "IG", "CG", "TVG", "OH", "MH", "NH"
                If mvarSpeed > mvarTL * 10 Then
                    mvarSpeed = mvarTL * 10
                    InfoPrint 1, "Max speed for NON Wire or Sonar Guided Torpedos is TL x 10"
                End If
            End Select
        Case UnGuidedMissile, GuidedMissile
            If mvarSpeed > mvarTL * 1200 Then
                mvarSpeed = mvarTL * 1200
                InfoPrint 1, "Max speed for non-space Missiles is TL x 1200"
            End If
        End Select
    End If
End Property

Public Property Get Speed() As Single
    Speed = mvarSpeed
End Property
Public Property Get AccelG() As Single
    AccelG = mvarAccelG
End Property
Public Property Get AccelMPH() As Single
    AccelMPH = mvarAccelMPH
End Property
Public Property Let PayloadCost(ByVal vdata As Double)
    mvarPayloadCost = vdata
End Property
Public Property Get PayloadCost() As Double
    PayloadCost = mvarPayloadCost
End Property
Public Property Let PayloadWeight(ByVal vdata As Double)
    mvarPayloadWeight = vdata
End Property
Public Property Get PayloadWeight() As Double
    PayloadWeight = mvarPayloadWeight
End Property
Public Property Let WarheadCost(ByVal vdata As Double)
    mvarWarheadCost = vdata
End Property
Public Property Get WarheadCost() As Double
    WarheadCost = mvarWarheadCost
End Property
Public Property Let WarheadWeight(ByVal vdata As Double)
    mvarWarheadWeight = vdata
End Property
Public Property Get WarheadWeight() As Double
    WarheadWeight = mvarWarheadWeight
End Property
Public Property Let GuidanceCost(ByVal vdata As Double)
    mvarGuidanceCost = vdata
End Property
Public Property Get GuidanceCost() As Double
    GuidanceCost = mvarGuidanceCost
End Property
Public Property Let GuidanceWeight(ByVal vdata As Double)
    mvarGuidanceWeight = vdata
End Property
Public Property Get GuidanceWeight() As Double
    GuidanceWeight = mvarGuidanceWeight
End Property

Public Property Let SkillBonus(ByVal vdata As Integer)
    mvarSkillBonus = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If (mvarCompact) And (vdata <> 0) Then
        mvarSkillBonus = 0
        InfoPrint 1, "Skill Bonus cannot be given to system with Compact Guidance option."
        Exit Property
    End If
End Property
Public Property Get SkillBonus() As Integer
    SkillBonus = mvarSkillBonus
End Property
Public Property Let Skill(ByVal vdata As String)
    mvarSkill = vdata
    If mvarZZInit = 0 Then Exit Property
    
    Select Case mvarGuidanceSystem
    Case "IG", "WG", "WGSH", "RCG", "LCG", "NCG", "RTVG", "LTVG", "NTVG"
        mvarSkill = vdata
    End Select
End Property
Public Property Get Skill() As String
    Skill = mvarSkill
End Property
Public Property Let PopUp(ByVal vdata As Boolean)
    mvarPopUp = vdata
End Property
Public Property Get PopUp() As Boolean
    PopUp = mvarPopUp
End Property
Public Property Let MidCourseUpdate(ByVal vdata As Boolean)
    mvarMidCourseUpdate = vdata
End Property
Public Property Get MidCourseUpdate() As Boolean
    MidCourseUpdate = mvarMidCourseUpdate
End Property
Public Property Let Brilliant(ByVal vdata As Boolean)
    mvarBrilliant = vdata
End Property
Public Property Get Brilliant() As Boolean
    Brilliant = mvarBrilliant
End Property
Public Property Let WarHead(ByVal vdata As String)
    mvarWarHead = vdata
End Property
Public Property Get WarHead() As String
    WarHead = mvarWarHead
End Property
Public Property Let Diameter(ByVal vdata As Single)
    mvarDiameter = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If vdata > 3000 Then
        InfoPrint 1, "FYI, maximum suggested diameter for missiles and torpedos is 3,000mm"
    ElseIf vdata >= 10 Then
       
    Else
        InfoPrint 1, "Minimum diameter for missiles and torpedos is 10mm"
        mvarDiameter = 10
        Exit Property
    End If
End Property


Public Property Get Diameter() As Single
    Diameter = mvarDiameter
End Property
Public Property Let CheapGuidance(ByVal vdata As Boolean)
    mvarCheapGuidance = vdata
End Property
Public Property Get CheapGuidance() As Boolean
    CheapGuidance = mvarCheapGuidance
End Property
Public Property Let Compact(ByVal vdata As Boolean)
    mvarCompact = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If mvarSkillBonus > 0 Then
        mvarCompact = False
        InfoPrint 1, "A compact system must have 0 Skill Bonus."
        Exit Property
    End If
End Property
Public Property Get Compact() As Boolean
    Compact = mvarCompact
End Property
Public Property Let CustomDescription(ByVal vdata As String)
    mvarCustomDescription = vdata
End Property
Public Property Get CustomDescription() As String
    CustomDescription = mvarCustomDescription
End Property
Public Property Let Weight(ByVal vdata As Double)
    mvarWeight = vdata
End Property
Public Property Get Weight() As Double
    Weight = mvarWeight
End Property
Public Property Let Volume(ByVal vdata As Double)
    mvarVolume = vdata
End Property
Public Property Get Volume() As Double
    Volume = mvarVolume
End Property
Public Property Let TL(ByVal vdata As Integer)
Dim i As Integer

    If vdata = 0 Then Exit Property
    mvarTL = vdata
    If mvarZZInit = 0 Then Exit Property

    If mvarDatatype = UnGuidedMissile Then
        If vdata < 3 Then
            vdata = 3
            InfoPrint 1, "Minimum TL for Un-guided Missiles is 3"
            Exit Property
        End If
    ElseIf mvarDatatype = UnGuidedTorpedo Then
        If vdata < 5 Then
            vdata = 5
            InfoPrint 1, "Minimum TL for Un-guided Torpedo is 5"
            Exit Property
        End If
    End If
    
    '//if we have guided munitions, then the user must reset them after switching tech levels
    If mvarDatatype = ProximityMine Then
        mvarGuidanceSystem = "PSH"
    ElseIf mvarDatatype = SmartBomb Then
        mvarGuidanceSystem = "RTVG"
    Else
        mvarGuidanceSystem = "IG"
    End If
    mvarBrilliantGuidanceSystem = "none"
    InfoPrint 1, "Guidance systems must be reset after changing tech level."
    
    mvarTL = vdata
    GetMatrixIndex
End Property

Public Property Get TL() As Integer
    TL = mvarTL
End Property

Public Property Let SurfaceArea(ByVal vdata As Double)
    mvarSurfaceArea = vdata
End Property
Public Property Get SurfaceArea() As Double
    SurfaceArea = mvarSurfaceArea
End Property
Public Property Let Quantity(ByVal vdata As Integer)
    mvarQuantity = vdata
End Property
Public Property Get Quantity() As Integer
    Quantity = mvarQuantity
End Property
Public Property Let Parent(ByVal vdata As String)
    mvarParent = vdata
End Property
Public Property Get Parent() As String
    Parent = mvarParent
End Property
Public Property Let Key(ByVal vdata As String)
    mvarKey = vdata
End Property
Public Property Get Key() As String
    Key = mvarKey
End Property
Public Property Let SelectedImage(ByVal vdata As Integer)
    mvarSelectedImage = vdata
End Property
Public Property Get SelectedImage() As Integer
    SelectedImage = mvarSelectedImage
End Property
Public Property Let Image(ByVal vdata As Integer)
    mvarImage = vdata
End Property
Public Property Get Image() As Integer
    Image = mvarImage
End Property
Public Property Let Description(ByVal vdata As String)
    mvarDescription = vdata
End Property
Public Property Get Description() As String
    Description = mvarDescription
End Property
Public Property Let Datatype(ByVal vdata As Integer)
    mvarDatatype = vdata
End Property
Public Property Get Datatype() As Integer
    Datatype = mvarDatatype
End Property
Public Property Let Cost(ByVal vdata As Double)
    mvarCost = vdata
End Property
Public Property Get Cost() As Double
    Cost = mvarCost
End Property
Public Property Let Custom(ByVal vdata As Boolean)
    mvarCustom = vdata
End Property
Public Property Get Custom() As Boolean
    Custom = mvarCustom
End Property
Public Property Let DR(ByVal vdata As Long)
    mvarDR = vdata
End Property
Public Property Get DR() As Long
    DR = mvarDR
End Property
Public Property Let HitPoints(ByVal vdata As Double)
    mvarHitPoints = vdata
End Property
Public Property Get HitPoints() As Double
    HitPoints = mvarHitPoints
End Property
Public Property Let Ruggedized(ByVal vdata As Boolean)
    mvarRuggedized = vdata
End Property
Public Property Get Ruggedized() As Boolean
    Ruggedized = mvarRuggedized
End Property

Private Sub Class_Initialize()
' set the default properties
mvarCustom = False
TL = gVehicleTL
mvarQuantity = 1
mvarCheapGuidance = False
mvarCompact = False
mvarWarHead = "Solid"
mvarWarheadSize = "normal"
mvarDiameter = 100
mvarBusMissiles = "1"
'this option only available for relevant missile /torp types
mvarBrilliant = False
mvarMidCourseUpdate = False
mvarPopUp = False

mvarStealth = False
'this option only available for relevant missile types
mvarSpaceMissile = False


End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case UnGuidedMissile, UnGuidedTorpedo
        mvarGuidanceSystem = "none"
        mvarBrilliantGuidanceSystem = "none"
        mvarSpeed = 200
        mvarMotorWeight = 100
        mvarWarHead = "Solid"
        
    Case IronBomb, RetardedBomb
        mvarGuidanceSystem = "none"
        mvarBrilliantGuidanceSystem = "none"
        mvarSpeed = 80
        mvarWarHead = "LE"
        
    Case CommandTriggerMine, SmartTriggerMine, SelfDestructSystem, ContactMine
        mvarGuidanceSystem = "none"
        mvarBrilliantGuidanceSystem = "none"
        mvarSpeed = 80
        mvarWarHead = "LE"
        
    Case SmartBomb
        mvarGuidanceSystem = "RTVG"
        mvarBrilliantGuidanceSystem = "none"
        mvarSpeed = 80
        mvarWarHead = "LE"
        
    Case ProximityMine
        mvarGuidanceSystem = "PSH"
        mvarBrilliantGuidanceSystem = "none"
        mvarWarHead = "LE"
        
    Case PressureTriggerMine
        mvarGuidanceSystem = "none"
        mvarDetonationWeight = 100
        mvarBrilliantGuidanceSystem = "none"
        mvarWarHead = "LE"
        
    Case GuidedMissile
        mvarGuidanceSystem = "IG"
        mvarBrilliantGuidanceSystem = "none"
        mvarSpeed = 200
        mvarMotorWeight = 100
    Case GuidedTorpedo
        mvarGuidanceSystem = "IG"
        mvarBrilliantGuidanceSystem = "none"
        mvarSpeed = 200
        mvarMotorWeight = 100
End Select

End Sub


Public Sub StatsUpdate()
Dim sPrint1 As String
Dim sPrint2 As String
Dim sPrintPlural As String
Dim sPrintPlural2 As String
Dim sPrintPlural3 As String
Dim sPrintDirection As String
Dim QRugMod As Single
Dim RugHitMod As Long

    mvarZZInit = 1
    On Error Resume Next
    
    mvarLocation = GetLocation
    
    'set the ruggedized and quantity multipliers
    If mvarQuantity < 1 Then mvarQuantity = 1
    If mvarRuggedized Then
        QRugMod = 1.5 * mvarQuantity
        RugHitMod = 2
    Else
        QRugMod = 1 * mvarQuantity
        RugHitMod = 1
    End If
    
    'calculate statistics
    mvarWarheadWeight = GetWarheadWeight
    mvarWarheadCost = GetWarheadCost
    mvarSkill = GetSkill 'skill must be calc'd before guidance weight
    mvarGuidanceWeight = GetGuidanceWeight
    mvarGuidanceCost = GetGuidanceCost
    mvarPayloadWeight = GetPayloadWeight
    mvarPayloadCost = GetPayloadCost
    mvarMotorCost = GetMotorCost
    'todo: make sure stealth is calculated into everything (pg 117 bottomish)
    mvarWeight = GetWeight
    mvarVolume = GetVolume
    mvarCost = GetCost
    mvarMalfunction = GetMalfunction
    mvarKEDamage = GetDamage
    GetTypeDamages 'call sub to update Damages based on Ammunition Type
    'todo: determine which of the below is not needed by NON missiles or torpedoes
    mvarEndurance = GetEndurance 'todo make sure space missiles have diff. calc here
    mvarMaxRange = GetMaxRange 'todo max range in inapplicable for space missiles
    mvarhalfDamage = GetHalfDamage
    mvarMinRange = GetMinRange
    mvarAccuracy = GetAccuracy
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea) * RugHitMod

    
    'cost, malf, and accuracy modifiers for Cheap, Fine and Very Fine quality are calced in the functions below
    
    '//update the cost,weight,volume, surface area and volume based on quantity and ruggedized options
    mvarCost = mvarCost * QRugMod
    mvarWeight = mvarWeight * QRugMod
    mvarVolume = mvarVolume * QRugMod
    
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)

        
    'produce the print output
    'produce the print output
    If mvarRuggedized Then
        sPrint1 = "ruggedized "
    Else
        sPrint1 = ""
    End If

    
    'produce the print output
    If mvarGuidanceSystem <> "none" Then
        sPrint1 = sPrint1 & mvarGuidanceSystem + " "
    End If
    If mvarCheapGuidance Then
        sPrint2 = sPrint2 + ", cheap guidance"
    End If
    If mvarCompact Then
        sPrint2 = sPrint2 + ", compact guidance"
    End If
    If mvarBrilliantGuidanceSystem <> "none" Then
        sPrint2 = sPrint2 + ", " & mvarBrilliantGuidanceSystem & " terminal guidance sytem"
    End If
    If mvarMidCourseUpdate Then
        sPrint2 = sPrint2 + ", mid-course update"
    End If
    If mvarPopUp Then
        sPrint2 = sPrint2 + ", pop up"
    End If
    If mvarSkillBonus <> 0 Then
        sPrint2 = sPrint2 + ", +" + Format(mvarSkillBonus) + " skill bonus"
    End If
    If mvarParachute Then
        sPrint2 = sPrint2 + ", parachute mine"
    End If
    If mvarSpaceMissile Then
        sPrint2 = sPrint2 + ", space missile"
    End If
    If mvarStealth Then
        sPrint2 = sPrint2 + ", stealth option"
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
    
    mvarPrintOutput = NumericToString(mvarQuantity) + " TL" & mvarTL & " " & mvarDiameter & "mm " & sPrint1 + " " + mvarCustomDescription + sPrintPlural + sPrint2 _
        + " (" + mvarLocation + sPrintDirection + ", HP " & mvarHitPoints & sPrintPlural2 + ", " + sPrintPlural3 + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ")." + mvarComment

End Sub



Private Function GetWarheadWeight() As Single
Dim W As Single
Dim DCubed As Single

If mvarWarheadSize = "small" Then
    W = 0.25
ElseIf mvarWarheadSize = "modest" Then
    W = 0.5
ElseIf mvarWarheadSize = "normal" Then
    W = 1
ElseIf mvarWarheadSize = "big" Then
    W = 1.5
ElseIf mvarWarheadSize = "huge" Then
    W = 2
Else
    W = 1 'this should never occur
End If

DCubed = mvarDiameter ^ 3
GetWarheadWeight = Round(DCubed / 125000 * W, 2)

End Function

Private Function GetWarheadCost() As Single
Dim Divisor As Integer
Dim Modifier As Integer
Dim TempCost As Single
Dim NukeExtraCost As Single

 Select Case mvarWarHead
 
 Case "Solid"
        Modifier = 10
 Case "LE", "HE", "HEC", "HEDC", "HP", "SAPHE"
       Modifier = 20
 Case "AP", "APDU", "HEAT", "HEDP"
        Modifier = 30
 Case "APHD", "API", "Beehive", "HESH"
    Modifier = 40
 Case "APEX", "HEPF"
    Modifier = 50
 Case "FASCAM", "ICM"
    Modifier = 100
 Case "SICM"
    Modifier = 1000
 Case ".01 kiloton Nuke", ".001 kiloton Nuke", ".0001 kiloton Nuke"
    Modifier = 1500
 Case "SATNUC"
    Modifier = 2000
 Case "1 megaton Nuke", "100 kiloton Nuke", "10 kiloton Nuke", "1 kiloton Nuke", ".1 kiloton Nuke"
    Modifier = 30
 Case "CHEM"
    Modifier = 2
 End Select
        
If mvarTL <= 5 Then
    Divisor = 10
ElseIf mvarTL = 6 Then
    Divisor = 4
ElseIf mvarTL >= 7 Then
    Divisor = 1
End If


TempCost = Modifier * mvarWarheadWeight / Divisor

' todo: 07/29/2000 Ultimately should change this. This was a last minute fix
' and the values are hard coded.

Select Case mvarWarHead
    Case "1 megaton Nuke"
          NukeExtraCost = 64000
    Case "100 kiloton Nuke"
        NukeExtraCost = 48000
    Case "10 kiloton Nuke"
        NukeExtraCost = 42000
    Case "1 kiloton Nuke"
        NukeExtraCost = 36000
    Case ".1 kiloton Nuke"
        NukeExtraCost = 15000
    Case ".01 kiloton Nuke"
        NukeExtraCost = 12000
    Case ".001 kiloton Nuke"
        NukeExtraCost = 9000
    Case ".0001 kiloton Nuke"
        NukeExtraCost = 6000
        
End Select
    
    If mvarTL >= 9 Then
        If mvarTL >= 15 Then
            NukeExtraCost = NukeExtraCost / 4
        Else
            NukeExtraCost = NukeExtraCost / 2
        End If
    End If
    
 TempCost = TempCost + NukeExtraCost
GetWarheadCost = TempCost

End Function

Public Sub GetMatrixIndex()
Dim i As Integer

mvarMatrixPos = 0 'init the counter
For i = 1 To UBound(GuidanceMatrix())
    If GuidanceMatrix(i).Name = mvarGuidanceSystem Then
        If GuidanceMatrix(i).TL >= mvarTL Then
            mvarMatrixPos = i
            Exit For
        Else
            mvarMatrixPos = i
        End If
    End If
Next

mvarMatrixPos2 = 0
If mvarBrilliantGuidanceSystem <> "none" Then
    For i = 1 To UBound(GuidanceMatrix())
        If GuidanceMatrix(i).Name = mvarBrilliantGuidanceSystem Then
            If GuidanceMatrix(i).TL >= mvarTL Then
                mvarMatrixPos2 = i
                Exit For
            Else
                mvarMatrixPos2 = i
            End If
        End If
    Next
End If
End Sub
Private Function GetGuidanceWeight() As Single
Dim GWeight As Single
Dim b As Single
Dim s As Integer


Select Case mvarDatatype
    Case UnGuidedMissile, UnGuidedTorpedo, IronBomb, RetardedBomb, CommandTriggerMine, SmartTriggerMine, SelfDestructSystem, ContactMine, ProximityMine, PressureTriggerMine
        GetGuidanceWeight = 0
    Exit Function
End Select

'//total the weights for both our guidance systems
GWeight = GuidanceMatrix(mvarMatrixPos).WeightMod

If mvarMatrixPos2 <> 0 Then '//if we have a second guidance system add it
    '//find our brilliance modifier
    Select Case mvarTL
        Case 7
            b = 5
        Case 8
            b = 2
        Case 9
            b = 1
        Case 10
            b = 0.25
        Case Is >= 11
            b = 0.1
        Case Else
            b = 0
    End Select
    GWeight = GWeight + GuidanceMatrix(mvarMatrixPos2).WeightMod + b
End If


If mvarCompact Then
    'compact will only have half the weight
    GWeight = GWeight / 2
Else
    'add weight modifier for skill bonus
    s = Int(mvarSkillBonus / 2)
    If s = 0 Then
    Else
        GWeight = GWeight * (2 ^ s) 'double weight for each +2 of bonus skill
    End If
End If

GetGuidanceWeight = GWeight
End Function


Private Function GetGuidanceCost() As Single
Dim tcost As Single
Dim MuPu As Single
Const BrilliantCost = 10000

Select Case mvarDatatype
    Case UnGuidedMissile, UnGuidedTorpedo, IronBomb, RetardedBomb, CommandTriggerMine, SmartTriggerMine, SelfDestructSystem, ContactMine, ProximityMine, PressureTriggerMine
        GetGuidanceCost = 0
        Exit Function
    
    Case Else
        '//find our total guidance system weight
        If mvarBrilliantGuidanceSystem <> "none" Then
            tcost = GuidanceMatrix(mvarMatrixPos).CostMod + GuidanceMatrix(mvarMatrixPos2).CostMod + BrilliantCost
            tcost = tcost * mvarGuidanceWeight
        Else
            tcost = GuidanceMatrix(mvarMatrixPos).CostMod * mvarGuidanceWeight
        End If
End Select
    
'cheap guidance system is uses half weight value
If mvarCheapGuidance Then
    tcost = tcost / 2
End If

'modifiers for mid course update and pop up
If (mvarMidCourseUpdate) And (mvarPopUp) Then
    MuPu = 1.5
ElseIf mvarMidCourseUpdate Then
    MuPu = 1.2
ElseIf mvarPopUp Then
    MuPu = 1.25
Else
    MuPu = 1
End If

'get final
tcost = tcost * MuPu
GetGuidanceCost = tcost
End Function

Private Function GetPayloadWeight() As Single

If (mvarDatatype = GuidedMissile) Or (mvarDatatype = UnGuidedMissile) Then
    GetPayloadWeight = (mvarGuidanceWeight + mvarWarheadWeight) * mvarBusMissiles
Else
    GetPayloadWeight = mvarGuidanceWeight + mvarWarheadWeight
End If

End Function

Private Function GetPayloadCost() As Single
If (mvarDatatype = GuidedMissile) Or (mvarDatatype = UnGuidedMissile) Then
    GetPayloadCost = (mvarGuidanceCost + mvarWarheadCost) * mvarBusMissiles
Else
    GetPayloadCost = mvarGuidanceCost + mvarWarheadCost
End If

End Function


Private Function GetMotorCost() As Single
Dim M As Single
Select Case mvarDatatype

Case GuidedMissile, GuidedTorpedo
    'for guided weapons
    If mvarTL <= 5 Then
        M = 10
    ElseIf mvarTL = 6 Then
        M = 20
    ElseIf mvarTL <= 8 Then
        M = 100
    ElseIf mvarTL >= 9 Then
        M = 50
    End If
Case UnGuidedMissile, UnGuidedTorpedo
    'for unguided weapons
    If mvarTL <= 5 Then
        M = 1
    ElseIf mvarTL = 6 Then
        M = 2
    ElseIf mvarTL <= 8 Then
        M = 20
    ElseIf mvarTL >= 9 Then
        M = 10
    End If
End Select

GetMotorCost = mvarMotorWeight * M
End Function


Private Function GetWeight() As Double
Dim TempWeight As Double

TempWeight = mvarMotorWeight + mvarPayloadWeight

If mvarDatatype = SelfDestructSystem Then
    TempWeight = TempWeight / 2
End If
GetWeight = Round(TempWeight, 2)
End Function


Private Function GetVolume() As Double

GetVolume = mvarWeight / 50

End Function


Private Function GetCost() As Double
Dim StealthMod As Single
Dim TempCost As Double

If mvarStealth = False Then
Else
    If mvarTL = 7 Then
        StealthMod = mvarWeight * 2000
    ElseIf mvarTL = 8 Then
        StealthMod = mvarWeight * 1000
    ElseIf mvarTL >= 9 Then
        StealthMod = mvarWeight * 500
    Else
        StealthMod = 0
    End If
End If

TempCost = mvarPayloadCost + mvarMotorCost + StealthMod

If mvarDatatype = SelfDestructSystem Then
    TempCost = TempCost / 2
End If

GetCost = Round(TempCost, 2)
End Function

Private Function GetMalfunction() As String
Dim TempMalf As String
Dim count As Integer 'holds the number of reliable traits.  >=2 increase Crit to Ver.

If (mvarDatatype = UnGuidedMissile) Or (mvarDatatype = UnGuidedTorpedo) Then
    If mvarTL <= 3 Then
        TempMalf = 13
    ElseIf mvarTL = 4 Then
        TempMalf = 14
    ElseIf mvarTL = 5 Then
        TempMalf = 16
    ElseIf mvarTL = 6 Then
        TempMalf = "Crit."
    ElseIf mvarTL = 7 Then
        TempMalf = "Crit."
    ElseIf mvarTL >= 8 Then
        TempMalf = "Crit."
    End If
Else
    If mvarTL = 6 Then
        TempMalf = 15
    ElseIf mvarTL = 7 Then
        TempMalf = "Crit."
    ElseIf mvarTL >= 8 Then
        TempMalf = "Crit."
    End If
End If


GetMalfunction = TempMalf
End Function



Sub GetTypeDamages()
Dim i As Integer
Dim Suffix1 As String
Dim Suffix2 As String
Dim x As Single
Dim D As Single
Dim h As Single
Dim s As Single
Dim C As Single
Dim SizeMod As Single
Dim Divisor As Single
On Error Resume Next
mvarBurstRadius = -1  'reset to -1 which tells GVD that this option is not applicable for given warhead type
D = mvarDiameter
C = D ^ 3
If mvarTL <= 5 Then
    x = 0.375
    h = 0.25
ElseIf mvarTL = 6 Then
    x = 0.75
    h = 0.25
ElseIf mvarTL = 7 Then
    x = 1
    h = 0.375
ElseIf mvarTL = 8 Then
    x = 3
    h = 1
ElseIf mvarTL >= 9 Then
    x = 4.5
    h = 1.5
End If

If mvarDiameter <= 44 Then
    s = 0.2
ElseIf mvarDiameter <= 49 Then
    s = 0.3
ElseIf mvarDiameter <= 54 Then
    s = 0.5
ElseIf mvarDiameter <= 59 Then
    s = 0.7
Else
    s = 1
End If

mvarDamage2 = Empty

'ammunition modifier updates to 1/2D and MaxRange
For i = 1 To UBound(AmmoMatrix())
    If AmmoMatrix(i).Name = mvarWarHead Then
            Exit For
    End If
Next

'get suffix for fragmentation damage
If AmmoMatrix(i).Fragmentation Then
    If mvarDiameter < 20 Then
        Suffix2 = ""
    ElseIf mvarDiameter <= 34 Then
        Suffix2 = "[2d]"
    ElseIf mvarDiameter <= 59 Then
        Suffix2 = "[4d]"
    ElseIf mvarDiameter <= 94 Then
        Suffix2 = "[6d]"
    ElseIf mvarDiameter <= 160 Then
        Suffix2 = "[10d]"
    ElseIf mvarDiameter > 160 Then
        Suffix2 = "[12d]"
    End If
End If
'get suffix for armor divisor
If AmmoMatrix(i).Divisor <> 0 Then Suffix1 = "(" & AmmoMatrix(i).Divisor & ")"

TypeDamage1 = AmmoMatrix(i).Damage1
TypeDamage2 = AmmoMatrix(i).Damage2

Select Case AmmoMatrix(i).Name

Case "Solid", "Chainshot", "AP", "APC", "API", "APCR", "HP", "Plastic", "Baton", _
     "Needle", "APDS", "APS", "APDU", "APDSDU", "APFSDS", "APFSDSDU", "Superwire", _
     "APSHD", "APHD", "APDSHD", "APFSDSHD", "Shotshell", "Canister", "Shrapnel", "Beehive"
    mvarDamage1 = ConvertDamage(mvarKEDamage) & Suffix1
          
Case "LE", "HE", "HEC", "HEDC", "HEPF", "HESH"
    If mvarWarheadSize = "small" Then
        SizeMod = 0.5
    ElseIf mvarWarheadSize = "modest" Then
        SizeMod = 1
    ElseIf mvarWarheadSize = "normal" Then
        SizeMod = 2
    ElseIf mvarWarheadSize = "big" Then
        SizeMod = 3
    Else
        SizeMod = 4
    End If
    mvarDamage1 = ConvertDamage((C * x / AmmoMatrix(i).Multiplier) * SizeMod) & Suffix2

Case "SAPLE", "SAPHE", "APEX"
    If mvarWarheadSize = "small" Then
        SizeMod = 0.5
    ElseIf mvarWarheadSize = "modest" Then
        SizeMod = 1
    ElseIf mvarWarheadSize = "normal" Then
        SizeMod = 2
    ElseIf mvarWarheadSize = "big" Then
        SizeMod = 3
    Else
        SizeMod = 4
    End If
    mvarDamage1 = ConvertDamage(mvarKEDamage) & Suffix1
    mvarDamage2 = ConvertDamage((C * x / AmmoMatrix(i).Multiplier) * SizeMod) & Suffix2

Case "HEAT"
    If mvarWarheadSize = "small" Then
        SizeMod = 1
    ElseIf mvarWarheadSize = "modest" Then
        SizeMod = 1.25
    ElseIf mvarWarheadSize = "normal" Then
        SizeMod = 1.6
    ElseIf mvarWarheadSize = "big" Then
        SizeMod = 1.8
    Else
        SizeMod = 2
    End If
    If mvarPopUp Then
        Divisor = 1.5
    Else
        Divisor = 1
    End If
    mvarDamage1 = ConvertDamage(D * s * h * SizeMod / Divisor) & Suffix1

Case "HEDP"
    If mvarWarheadSize = "small" Then
        SizeMod = 1
    ElseIf mvarWarheadSize = "modest" Then
        SizeMod = 1.25
    ElseIf mvarWarheadSize = "normal" Then
        SizeMod = 1.6
    ElseIf mvarWarheadSize = "big" Then
        SizeMod = 1.8
    Else
        SizeMod = 2
    End If
    If mvarPopUp Then
        Divisor = 1.5
    Else
        Divisor = 1
    End If
    mvarDamage1 = ConvertDamage(D * s * h * SizeMod / Divisor) & Suffix1 & Suffix2
    
Case "CHEM", "FASCAM", "ICM", "SICM", "SATNUC"
    ' NOTE: These actually produce the BURST RADIUS in YARDS and NOT Damage (read "Damage" description at top-ish of page 112)
    If mvarWarheadSize = "small" Then
        SizeMod = 1
    ElseIf mvarWarheadSize = "modest" Then
        SizeMod = 1.4
    ElseIf mvarWarheadSize = "normal" Then
        SizeMod = 2
    ElseIf mvarWarheadSize = "big" Then
        SizeMod = 2.4
    Else
        SizeMod = 2.8
    End If
    
    mvarBurstRadius = AmmoMatrix(i).Multiplier * D * D * SizeMod
    ' these also have special damages which dont plug into the formula
    Select Case AmmoMatrix(i).Name
        Case "CHEM"
            mvarTypeDamage1 = "Spec"
            mvarDamage1 = "variable"
        Case "FASCAM"
            ' sidebar 195
            mvarTypeDamage1 = "Explosive"
            mvarDamage1 = "6d x" & Format(mvarTL - 6) & "(10)"
        Case "SATNUC"
            ' sidebar 193
            mvarTypeDamage1 = "Explosive"
            mvarTypeDamage2 = "Heat + Concussion"
            mvarDamage1 = "6d x 2000"
            mvarDamage2 = "6d x 200"
        Case "SICM"
            mvarTypeDamage1 = "Fragmentation"
            If mvarTL <= 8 Then
                mvarDamage1 = "6d x 2(5)"
            Else
                mvarDamage1 = "6d x 3(5)"
            End If
        Case "ICM"
            mvarTypeDamage1 = "Concussion"
            mvarTypeDamage2 = "Fragmentation"
            mvarDamage1 = "2d"
            mvarDamage2 = "4d"
                If mvarTL = 8 Then
                    mvarDamage1 = "4d"
                ElseIf mvarTL >= 9 Then
                    mvarDamage1 = "6d"
                End If
    End Select
    

Case "1 megaton Nuke", "100 kiloton Nuke", "10 kiloton Nuke", "1 kiloton Nuke", _
     ".1 kiloton Nuke", ".01 kiloton Nuke", ".001 kiloton Nuke", ".0001 kiloton Nuke"
    mvarDamage1 = AmmoMatrix(i).Formula
End Select

End Sub


Private Function GetDamage() As Single
Dim fKEDamage As Single 'holds numeric value of Kinetic Energy damage before its converted to a GURPS format string
Dim D As Single ' length modifier
Dim V As Single ' technology modifier
Dim A As Single 'tech level modifier
Dim W As Single

'get D
D = mvarDiameter

'get V
V = mvarSpeed / 800

'get A
If mvarWarHead = "Beehive" Then
    A = 0.25
Else
    A = 1
End If

'//modifier for warhead size
'x0.5 for small, x0.625 for modest, x0.8 for normal, x0.95 for big, x1 for huge.
Select Case mvarWarheadSize
    Case "small"
        W = 0.5
    Case "modest"
        W = 0.625
    Case "normal"
        W = 0.8
    Case "big"
        W = 0.95
    Case "huge"
        W = 1
End Select

fKEDamage = W * D * V * A

GetDamage = fKEDamage

End Function

Private Function GetSkill() As String
'cheap option reduces skill by 1
'compact option reduces skill by 2
Dim TSkill As Single
Dim TSkill2 As Single

'//tskill1 is for our primary guidance system.  If we have a second
'//guidance system, then we include its skill also seperated by a dash
TSkill = Val(GuidanceMatrix(mvarMatrixPos).Skill)
If TSkill <> 0 Then TSkill = TSkill + mvarTL
TSkill = TSkill + mvarSkillBonus
TSkill2 = Val(GuidanceMatrix(mvarMatrixPos2).Skill)
If TSkill2 <> 0 Then
    TSkill2 = TSkill2 + mvarTL
End If

If mvarCompact <> False Then
    TSkill = TSkill - 2
End If
If mvarCheapGuidance <> False Then
    TSkill = TSkill - 1
End If

If TSkill2 > 0 Then
    GetSkill = Format(TSkill) & "-" & Format(TSkill2)
Else
    GetSkill = Format(TSkill)
End If
End Function


Private Sub GetAcceleration()
    
    
    'get the acceleration in G's
    'mvarAccelG = mvarmvarweight
    
    'get the acceleration in MPH
    
    
    
End Sub
Private Function GetEndurance()
Dim E As Single
Dim T As Integer
Dim V As Single
Dim P As Integer
Dim sngTemp As Single

On Error Resume Next
'todo: bombs, mines and self destruct are getting divide by zero because they
'actually shouldnt even have an endurance!  Check rules before casing them out.
If mvarSpaceMissile = False Then
    'get E
    E = mvarMotorWeight / 10
    If E < 0.1 Then E = 0.1
    
    'get T
    If mvarTL <= 5 Then
        T = 10
    ElseIf mvarTL = 6 Then
        T = 30
    ElseIf mvarTL = 7 Then
        T = 60
    ElseIf mvarTL = 8 Then
        T = 120
    ElseIf mvarTL = 9 Then
        T = 240
    ElseIf mvarTL = 10 Then
        T = 360
    ElseIf mvarTL >= 11 Then
        T = 480
    End If
    
    'get V
    Select Case mvarDatatype
    Case UnGuidedTorpedo, GuidedTorpedo
          
         V = Round((mvarSpeed / 10) ^ 2, 2) 'fixed MPJ Oct.6.2002 Wasnt squaring this for torps
    
    Case UnGuidedMissile, GuidedMissile
        
        If mvarSpeed <= 400 Then
            V = Round((mvarSpeed / 100) ^ 2, 2)
        Else
            V = Round(((mvarSpeed / 200) + 2) ^ 2, 2)
        End If
    End Select
    'final calculation for non space missiles
    sngTemp = (mvarMotorWeight / mvarWeight) * E * (T / V)
Else
    'get P
    If mvarTL <= 4 Then
        P = 280
    ElseIf mvarTL = 5 Then
        P = 455
    ElseIf mvarTL = 6 Then
        P = 665
    ElseIf mvarTL = 7 Then
        P = 980
    ElseIf mvarTL <= 10 Then
        P = 1120
    ElseIf mvarTL >= 11 Then
        P = 1750
    End If
    'final calculation for non space missiles
    sngTemp = (mvarMotorWeight / mvarWeight) * (P / mvarSpeed)
End If

GetEndurance = Round(sngTemp, 2)
End Function


Private Function GetHalfDamage() As Single 'in yards
Dim TempHalfDamage As Single
Dim T As Single

If mvarDatatype = UnGuidedMissile Then
   TempHalfDamage = mvarEndurance
   If mvarEndurance > 1.66 Then TempHalfDamage = 1.66
   TempHalfDamage = TempHalfDamage * mvarSpeed
   'unguided space missiles have addition modifications to halfd
   If mvarSpaceMissile <> False Then
        T = mvarEndurance ^ 2
        If T > 2.75 Then T = 2.75
        TempHalfDamage = TempHalfDamage * T * mvarSpeed
   End If
Else 'all torps and guided missiles have no halfd range
    TempHalfDamage = 0
End If

GetHalfDamage = TempHalfDamage
End Function

Private Function GetMaxRange() As Double 'in yards

Dim TempMax As Double

'space missiles have no Max Range
If (mvarGuidanceSystem <> "WG") And (mvarSpaceMissile <> False) Then
    GetMaxRange = 0
    Exit Function
Else
    TempMax = mvarSpeed * mvarEndurance
End If

If mvarGuidanceSystem = "WG" Then
    If mvarDatatype = GuidedMissile Then
        If TempMax > 8800 Then TempMax = 8800
    ElseIf mvarDatatype = GuidedTorpedo Then
        If TempMax > 35200 Then TempMax = 35200
    End If
End If

GetMaxRange = Round(TempMax, 2)
End Function

Private Function GetMinRange() As Single
Dim TempMin As Single
Dim Half As Single

If (mvarGuidanceSystem = "WG") Or (mvarGuidanceSystem = "RTV") Or (mvarGuidanceSystem = "LTV") Or (mvarGuidanceSystem = "NTV") Or (mvarGuidanceSystem = "RCG") Or (mvarGuidanceSystem = "LCG") Or (mvarGuidanceSystem = "NCG") Or (mvarWarHead = "HEPF") Then
    If mvarTL <= 6 Then
        TempMin = 200
    ElseIf mvarTL = 7 Then
        TempMin = 100
    ElseIf mvarTL >= 8 Then
        TempMin = 50
    End If
    Half = mvarMaxRange / 2
    If TempMin > Half Then TempMin = Half
Else
    TempMin = 0
End If

GetMinRange = TempMin
End Function


Private Function GetAccuracy() As Long

Dim i As Integer 'ammomatrix array position
Dim R As Single
Dim Acc As Integer

R = mvarSpeed * 2
If R > mvarMaxRange Then R = mvarMaxRange
 
If R < 70 Then
    Acc = 6
ElseIf R <= 99 Then
    Acc = 7
ElseIf R <= 149 Then
    Acc = 8
ElseIf R <= 199 Then
    Acc = 9
ElseIf R <= 299 Then
    Acc = 10
ElseIf R <= 449 Then
    Acc = 11
ElseIf R <= 699 Then
    Acc = 12
ElseIf R <= 999 Then
    Acc = 13
ElseIf R <= 1499 Then
    Acc = 14
ElseIf R <= 1999 Then
    Acc = 15
ElseIf R <= 2999 Then
    Acc = 16
ElseIf R <= 4449 Then
    Acc = 17
ElseIf R <= 6999 Then
    Acc = 18
ElseIf R > 6999 Then
    Acc = 19
End If

'todo: note, accuracy is ONLY done for UNGUIDED missiles.  All others
'including UNguuided TORPEDOS are not applicable and should be a "-"
'instead of Acc, other missiles and torps use a SKILL rating

'NOTE: there is not modifiers for "cheap" UNguided missiles
GetAccuracy = Acc
End Function

Public Function FillTerminalGuidanceList() As String()
Dim guidancearray() As String
ReDim guidancearray(1)

guidancearray = mAddKey(guidancearray, "none")

'this just has to produce the correct guidance system list based on
'tech level and whether its a un/guided torpedo or un/guided missile
'The Let TL above takes care of whether the Guidance system is supported when the user
'changes the Tech Level AFTER having already selected a guidance system based on
'the former Tech Level

'NOTE2: This routine will NOT be used for bombs and mines.  Those have very simple
'rules for which systems are and are not allowed.
 
    If (mvarDatatype = GuidedMissile) Then
        If mvarTL >= 7 Then
            guidancearray = mAddKey(guidancearray, "IRH")  'missiles only
            guidancearray = mAddKey(guidancearray, "IIRH")  'missiles only
            guidancearray = mAddKey(guidancearray, "ARM")  'missiles only
            guidancearray = mAddKey(guidancearray, "SALH")  'missiles only
            guidancearray = mAddKey(guidancearray, "ARH")  'missiles only
        End If
        If mvarTL >= 8 Then
            guidancearray = mAddKey(guidancearray, "ALH")  'missiles only starting at Tech 8
            guidancearray = mAddKey(guidancearray, "OH")  'all
            guidancearray = mAddKey(guidancearray, "PEH")  'missiles only
            guidancearray = mAddKey(guidancearray, "PRH")  'missiles only
        End If
        If mvarTL >= 10 Then
            guidancearray = mAddKey(guidancearray, "MH")  'all starting at Tech 10
            guidancearray = mAddKey(guidancearray, "NH")  'all
        End If
    ElseIf (mvarDatatype = GuidedTorpedo) Then
        If mvarTL >= 6 Then
            guidancearray = mAddKey(guidancearray, "PSH")  'torps only
            guidancearray = mAddKey(guidancearray, "ASH")  'torps only
        End If
        If mvarTL >= 7 Then
            guidancearray = mAddKey(guidancearray, "SALH")
        End If
        If mvarTL >= 8 Then
            guidancearray = mAddKey(guidancearray, "ALH")
            guidancearray = mAddKey(guidancearray, "OH")  'all
        End If
        If mvarTL >= 10 Then
            guidancearray = mAddKey(guidancearray, "NH")  'all
        End If
    End If


FillTerminalGuidanceList = guidancearray
'call the Let Brilliant sub to check that the system is compatible
Brilliant = mvarBrilliant 'safe way to call the sub without changing the value

End Function
Public Function FillGuidanceList() As String()
Dim guidancearray() As String
ReDim guidancearray(1)

'this just has to produce the correct guidance system list based on
'tech level and whether its a un/guided torpedo or un/guided missile
'The Let TL above takes care of whether the Guidance system is supported when the user
'changes the Tech Level AFTER having already selected a guidance system based on
'the former Tech Level


'NOTE: compatible check with a brilliant system is done in Let Brilliant system

'NOTE2: This routine will NOT be used for bombs and mines.  Those have very simple
'rules for which systems are and are not allowed.
 
    If (mvarDatatype = UnGuidedMissile) Or (mvarDatatype = GuidedMissile) Or (mvarDatatype = SmartBomb) Then
        If mvarTL >= 6 Then
            If mvarDatatype <> SmartBomb Then
                guidancearray = mAddKey(guidancearray, "IG")  'all
                guidancearray = mAddKey(guidancearray, "WG")  'all
            End If
            guidancearray = mAddKey(guidancearray, "RCG")  'tvguided missiles not useable on torps or bus missiles
            guidancearray = mAddKey(guidancearray, "LCG")  'tvguided missiles not useable on torps or bus missiles
            guidancearray = mAddKey(guidancearray, "NCG")  'tvguided missiles not useable on torps or bus missiles
        End If
        If mvarTL >= 7 Then
            guidancearray = mAddKey(guidancearray, "RTVG")  'all starting at Tech 7
            guidancearray = mAddKey(guidancearray, "LTVG")  'all
            guidancearray = mAddKey(guidancearray, "NTVG")  'all
            guidancearray = mAddKey(guidancearray, "IRH")  'missiles only
            guidancearray = mAddKey(guidancearray, "IIRH")  'missiles only
            guidancearray = mAddKey(guidancearray, "ARM")  'missiles only
            guidancearray = mAddKey(guidancearray, "SARH")  'missiles only
            guidancearray = mAddKey(guidancearray, "SALH")  'missiles only
            guidancearray = mAddKey(guidancearray, "ARH")  'missiles only
        End If
        If mvarTL >= 8 Then
            guidancearray = mAddKey(guidancearray, "ALH")  'missiles only starting at Tech 8
            guidancearray = mAddKey(guidancearray, "OH")  'all
            guidancearray = mAddKey(guidancearray, "PEH")  'missiles only
            guidancearray = mAddKey(guidancearray, "PRH")  'missiles only
        End If
        If mvarTL >= 10 Then
            guidancearray = mAddKey(guidancearray, "MH")  'all starting at Tech 10
            guidancearray = mAddKey(guidancearray, "NH")  'all
        End If
    ElseIf (mvarDatatype = UnGuidedTorpedo) Or (mvarDatatype = GuidedTorpedo) Or (ProximityMine) Then
        If mvarTL >= 6 Then
            If mvarDatatype <> ProximityMine Then
                guidancearray = mAddKey(guidancearray, "IG")  'all
                guidancearray = mAddKey(guidancearray, "WG")  'all
            End If
            guidancearray = mAddKey(guidancearray, "PSH")  'torps only
            guidancearray = mAddKey(guidancearray, "ASH")  'torps only
        End If
        If mvarTL >= 7 Then
            If mvarDatatype <> ProximityMine Then
               guidancearray = mAddKey(guidancearray, "RTVG")  'all
                guidancearray = mAddKey(guidancearray, "LTVG")  'all
                guidancearray = mAddKey(guidancearray, "NTVG")  'all
                guidancearray = mAddKey(guidancearray, "WGSH")  'torps only
            End If
        End If
        If mvarTL >= 8 Then
            guidancearray = mAddKey(guidancearray, "OH")  'all
        End If
        If mvarTL >= 10 Then
            guidancearray = mAddKey(guidancearray, "MH")  'all
            guidancearray = mAddKey(guidancearray, "NH")  'all
        End If
    End If


FillGuidanceList = guidancearray
'call the Let Brilliant sub to check that the system is compatible
Brilliant = mvarBrilliant 'safe way to call the sub without changing the value

End Function

Public Function FillAmmunitionList()
Dim ammoarray() As String
ReDim ammoarray(1)

    If mvarTL >= 3 Then
        Select Case mvarDatatype
        Case CommandTriggerMine, SmartTriggerMine, SelfDestructSystem, ContactMine, PressureTriggerMine, ProximityMine
            ammoarray = mAddKey(ammoarray, "LE") 'low explosive, a fused round filled with black powder
        Case Else
        
            ammoarray = mAddKey(ammoarray, "Solid") 'all weapons can use this
            ammoarray = mAddKey(ammoarray, "LE") 'low explosive, a fused round filled with black powder
        End Select
        
    End If
    If mvarTL >= 6 Then
        ammoarray = mAddKey(ammoarray, "AP") 'armor piercing bullet
        ammoarray = mAddKey(ammoarray, "API") 'an AP cannon round with added incendiary material
        ammoarray = mAddKey(ammoarray, "HE") 'high explosive, a modern explosive, fragmenting round
        ammoarray = mAddKey(ammoarray, "HEAT") 'a shaped-charged, armor piercing "anti-tank" round
        ammoarray = mAddKey(ammoarray, "HEC") 'HE Concussion, with bit blast and little fragmentation
        ammoarray = mAddKey(ammoarray, "HEDC") 'HE Depth Charge round for attacking submarines
        ammoarray = mAddKey(ammoarray, "HEPF") 'HE proximity fused round for anti aircraft fire
        ammoarray = mAddKey(ammoarray, "SAPHE") 'HE fused to go off after piercing armor
        ammoarray = mAddKey(ammoarray, "CHEM")
    End If
    If mvarTL >= 7 Then
        ammoarray = mAddKey(ammoarray, "HEDP") ' HE Dual Purpose, a fragmentation and HEAT round
        ammoarray = mAddKey(ammoarray, "HESH") ' HE Squash-Head, plastic explosive built to defeat armor
        ammoarray = mAddKey(ammoarray, "HP") 'hollow point, a bullet designed to expand in flesh
        ammoarray = mAddKey(ammoarray, "APDU") 'improved APCR with depleted uranium core
        ammoarray = mAddKey(ammoarray, "APEX") 'armor piercing explosive cannon round
        ammoarray = mAddKey(ammoarray, "1 megaton Nuke")
        ammoarray = mAddKey(ammoarray, "100 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, "10 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, "1 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".1 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, "Beehive") 'high-tech shrapnel using anti-personnel darts
        ammoarray = mAddKey(ammoarray, "FASCAM") 'drops a field of scattered anti-armor mines
        ammoarray = mAddKey(ammoarray, "ICM") 'cluster munitions that scatter grenades over an area
    End If
    If mvarTL >= 8 Then
        ammoarray = mAddKey(ammoarray, "SICM") 'smart ICM that home in on targets
    End If
    If mvarTL >= 9 Then
        'microknukes
        ammoarray = mAddKey(ammoarray, ".01 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".001 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".0001 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, "SATNUC") 'smart ICM with tiny, shaped nuclear warhead
    End If
    If mvarTL >= 11 Then
        ammoarray = mAddKey(ammoarray, "APHD") 'an advanced, hyperdense, saboted bullet
    End If

FillAmmunitionList = ammoarray
End Function

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub
