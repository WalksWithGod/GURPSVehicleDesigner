VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTorpMissile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarHitPoints As Double 'local copy
Private mvarDR As Long 'local copy
Private mvarCustom As Boolean 'local copy
Private mvarCost As Double 'local copy
Private mvarDatatype As Integer 'local copy
Private mvarDescription As String 'local copy
Private mvarImage As Integer 'local copy
Private mvarSelectedImage As Integer 'local copy
Private mvarKey As String 'local copy
Private mvarParent As String 'local copy
Private mvarQuantity As Integer 'local copy
Private mvarSurfaceArea As Double 'local copy
Private mvarTL As Integer 'local copy
Private mvarVolume As Double 'local copy
Private mvarWeight As Double 'local copy
Private mvarRuggedized As Boolean
Private mvarCustomDescription As String 'local copy
Private mvarGuidanceSystem As String
Private mvarBrilliantGuidanceSystem As String
Private mvarCheapGuidance As Boolean 'local copy
Private mvarDiameter As Single 'local copy
Private mvarWarHead As String 'local copy
Private mvarBrilliant As Boolean 'local copy
Private mvarMidCourseUpdate As Boolean 'local copy
Private mvarPopUp As Boolean 'local copy
Private mvarSkill As String 'local copy
Private mvarSkillBonus As Integer
Private mvarWarheadWeight As Single
Private mvarWarheadCost As Single
Private mvarGuidanceWeight As Single 'local copy
Private mvarGuidanceCost As Single 'local copy
Private mvarPayloadWeight As Single 'local copy
Private mvarPayloadCost As Single 'local copy
Private mvarSpeed As Single 'local copy
Private mvarMotorWeight As Single 'local copy
Private mvarMotorCost As Single 'local copy
Private mvarStealth As Boolean 'local copy
Private mvarMalfunction As String 'local copy
Private mvarTypeDamage1 As String 'local copy
Private mvarDamage1 As String 'local copy
Private mvarDamage2 As String
Private mvarTypeDamage2 As String
Private mvarEndurance As Single 'local copy
Private mvarAccelG As Single
Private mvarAccelMPH As Single
Private mvarMaxRange As Double 'local copy
Private mvarhalfDamage As Double 'local copy
Private mvarMinRange As Single 'local copy
Private mvarAccuracy As Long 'local copy
Private mvarWarheadSize As String
Private mvarBusMissiles As Long
Private mvarSpaceMissile As Boolean
Private mvarKEDamage As Double
Private mvarCompact As Boolean
Private mvarDetonationWeight As Single
Private mvarParachute As Boolean

Private mvarComment As String 'local copy
Private mvarCName As String 'local copy
Private mvarMatrixPos As Long 'local copy
Private mvarMatrixPos2 As Long
Private mvarPrintOutput As String 'local copy
Private mvarZZInit As Byte


Public Property Let PrintOutput(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrintOutput = 5
    mvarPrintOutput = vdata
End Property


Public Property Get PrintOutput() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrintOutput
    PrintOutput = mvarPrintOutput
End Property



Public Property Let MatrixPos(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MatrixPos = 5
    mvarMatrixPos = vdata
End Property


Public Property Get MatrixPos() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MatrixPos
    MatrixPos = mvarMatrixPos
End Property



Public Property Let CName(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CName = 5
    mvarCName = vdata
End Property


Public Property Get CName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CName
    CName = mvarCName
End Property



Public Property Let Comment(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Comment = 5
    mvarComment = vdata
End Property


Public Property Get Comment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Comment
    Comment = mvarComment
End Property


Public Property Let KEDamage(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.KEDamage = 5
    mvarKEDamage = vdata
End Property


Public Property Get KEDamage() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.KEDamage
    KEDamage = mvarKEDamage
End Property


Public Property Let DetonationWeight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DetonationWeight = 5
    mvarDetonationWeight = vdata
End Property


Public Property Get DetonationWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DetonationWeight
    DetonationWeight = mvarDetonationWeight
End Property

Public Property Let GuidanceSystem(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GuidanceSystem = 5

    mvarGuidanceSystem = vdata
    
        Dim i As Integer
    
    If mvarZZInit = 0 Then Exit Property
    GetMatrixIndex
    
End Property


Public Property Get GuidanceSystem() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GuidanceSystem
    GuidanceSystem = mvarGuidanceSystem
End Property

Public Property Let BrilliantGuidanceSystem(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Brilliantguidancesystem = 5

    mvarBrilliantGuidanceSystem = vdata
    
        Dim i As Integer
    
    If mvarZZInit = 0 Then Exit Property
    GetMatrixIndex
    
End Property


Public Property Get BrilliantGuidanceSystem() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Brilliantguidancesystem
    BrilliantGuidanceSystem = mvarBrilliantGuidanceSystem
End Property

Public Property Let BusMissiles(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BusMissiles = 5
    mvarBusMissiles = vdata
End Property

Public Property Get BusMissiles() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BusMissiles
    BusMissiles = mvarBusMissiles
End Property

Public Property Let Parachute(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Parachute = 5
    mvarParachute = vdata
End Property

Public Property Get Parachute() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parachute
    Parachute = mvarParachute
End Property


Public Property Let SpaceMissile(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SpaceMissile = 5
    mvarSpaceMissile = vdata
End Property

Public Property Get SpaceMissile() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SpaceMissile
    SpaceMissile = mvarSpaceMissile
End Property

Public Property Let WarheadSize(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WarheadSize = 5
    mvarWarheadSize = vdata
End Property


Public Property Get WarheadSize() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WarheadSize
    WarheadSize = mvarWarheadSize
End Property

Public Property Let Accuracy(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Accuracy = 5
    mvarAccuracy = vdata
End Property


Public Property Get Accuracy() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Accuracy
    Accuracy = mvarAccuracy
End Property



Public Property Let MinRange(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MinRange = 5
    mvarMinRange = vdata
End Property


Public Property Get MinRange() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinRange
    MinRange = mvarMinRange
End Property



Public Property Let halfDamage(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.halfDamage = 5
    mvarhalfDamage = vdata
End Property


Public Property Get halfDamage() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.halfDamage
    halfDamage = mvarhalfDamage
End Property



Public Property Let MaxRange(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxRange = 5
    mvarMaxRange = vdata
End Property


Public Property Get MaxRange() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxRange
    MaxRange = mvarMaxRange
End Property



Public Property Let Endurance(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Endurance = 5
    mvarEndurance = vdata
End Property


Public Property Get Endurance() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Endurance
    Endurance = mvarEndurance
End Property



Public Property Let Damage1(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Damage1 = 5
    mvarDamage1 = vdata
End Property


Public Property Get Damage1() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Damage1
    Damage1 = mvarDamage1
End Property

Public Property Let Damage2(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Damage2 = 5
    mvarDamage2 = vdata
End Property


Public Property Get Damage2() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Damage2
    Damage2 = mvarDamage2
End Property


Public Property Let TypeDamage1(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TypeDamage1 = 5
    mvarTypeDamage1 = vdata
End Property


Public Property Get TypeDamage1() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TypeDamage1
    TypeDamage1 = mvarTypeDamage1
End Property

Public Property Let TypeDamage2(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TypeDamage2 = 5
    mvarTypeDamage2 = vdata
End Property


Public Property Get TypeDamage2() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TypeDamage2
    TypeDamage2 = mvarTypeDamage2
End Property


Public Property Let Malfunction(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Malfunction = 5
    mvarMalfunction = vdata
End Property


Public Property Get Malfunction() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Malfunction
    Malfunction = mvarMalfunction
End Property




Public Property Let Stealth(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Stealth = 5
    mvarStealth = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If mvarTL < 7 Then
        mvarStealth = False
        InfoPrint 1, "Stealth technology is not available until Tech Level 7"
    End If

End Property


Public Property Get Stealth() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Stealth
    Stealth = mvarStealth
End Property



Public Property Let MotorCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MotorCost = 5
    mvarMotorCost = vdata
End Property


Public Property Get MotorCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MotorCost
    MotorCost = mvarMotorCost
End Property



Public Property Let MotorWeight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MotorWeight = 5
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
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MotorWeight
    MotorWeight = mvarMotorWeight
End Property

Public Property Let Speed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Speed = 5
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
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Speed

    Speed = mvarSpeed
End Property

Public Property Get AccelG() As Single
    AccelG = mvarAccelG
End Property

Public Property Get AccelMPH() As Single
    AccelMPH = mvarAccelMPH
End Property

Public Property Let PayloadCost(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PayloadCost = 5
    mvarPayloadCost = vdata
End Property


Public Property Get PayloadCost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PayloadCost
    PayloadCost = mvarPayloadCost
End Property



Public Property Let PayloadWeight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PayloadWeight = 5
    mvarPayloadWeight = vdata
End Property


Public Property Get PayloadWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PayloadWeight
    PayloadWeight = mvarPayloadWeight
End Property

Public Property Let WarheadCost(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WarheadCost = 5
    mvarWarheadCost = vdata
End Property


Public Property Get WarheadCost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WarheadCost
    WarheadCost = mvarWarheadCost
End Property



Public Property Let WarheadWeight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WarheadWeight = 5
    mvarWarheadWeight = vdata
End Property


Public Property Get WarheadWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WarheadWeight
    WarheadWeight = mvarWarheadWeight
End Property


Public Property Let GuidanceCost(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GuidanceCost = 5
    mvarGuidanceCost = vdata
End Property


Public Property Get GuidanceCost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GuidanceCost
    GuidanceCost = mvarGuidanceCost
End Property



Public Property Let GuidanceWeight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GuidanceWeight = 5
    mvarGuidanceWeight = vdata
End Property


Public Property Get GuidanceWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GuidanceWeight
    GuidanceWeight = mvarGuidanceWeight
End Property



Public Property Let SkillBonus(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SkillBonus = 5
    mvarSkillBonus = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If (mvarCompact) And (vdata <> 0) Then
        mvarSkillBonus = 0
        InfoPrint 1, "Skill Bonus cannot be given to system with Compact Guidance option."
        Exit Property
    End If

End Property


Public Property Get SkillBonus() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SkillBonus
    SkillBonus = mvarSkillBonus
End Property



Public Property Let Skill(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Skill = 5
    mvarSkill = vdata
    If mvarZZInit = 0 Then Exit Property
    
    Select Case mvarGuidanceSystem
    Case "IG", "WG", "WGSH", "RCG", "LCG", "NCG", "RTVG", "LTVG", "NTVG"
        mvarSkill = vdata
    End Select
End Property


Public Property Get Skill() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Skill
    Skill = mvarSkill
End Property



Public Property Let PopUp(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PopUp = 5
    mvarPopUp = vdata
End Property


Public Property Get PopUp() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PopUp
    PopUp = mvarPopUp
End Property



Public Property Let MidCourseUpdate(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MidCourseUpdate = 5
    mvarMidCourseUpdate = vdata
End Property


Public Property Get MidCourseUpdate() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MidCourseUpdate
    MidCourseUpdate = mvarMidCourseUpdate
End Property



Public Property Let Brilliant(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Brilliant = 5
    mvarBrilliant = vdata
End Property


Public Property Get Brilliant() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Brilliant
    Brilliant = mvarBrilliant
End Property

Public Property Let WarHead(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WarHead = 5
    mvarWarHead = vdata
End Property


Public Property Get WarHead() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WarHead
    WarHead = mvarWarHead
End Property



Public Property Let Diameter(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Diameter = 5
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
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Diameter
    Diameter = mvarDiameter
End Property



Public Property Let CheapGuidance(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CheapGuidance = 5
    mvarCheapGuidance = vdata
End Property


Public Property Get CheapGuidance() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CheapGuidance
    CheapGuidance = mvarCheapGuidance
End Property


Public Property Let Compact(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Compact = 5
    mvarCompact = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If mvarSkillBonus > 0 Then
        mvarCompact = False
        InfoPrint 1, "A compact system must have 0 Skill Bonus."
        Exit Property
    End If
    
End Property


Public Property Get Compact() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Compact
    Compact = mvarCompact
End Property

Public Property Let CustomDescription(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CustomDescription = 5
    mvarCustomDescription = vdata
End Property


Public Property Get CustomDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CustomDescription
    CustomDescription = mvarCustomDescription
End Property



Public Property Let Weight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Weight = 5
    mvarWeight = vdata
End Property


Public Property Get Weight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Weight
    Weight = mvarWeight
End Property

Public Property Let Volume(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Volume = 5
    mvarVolume = vdata
End Property


Public Property Get Volume() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Volume
    Volume = mvarVolume
End Property

Public Property Let TL(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TL = 5
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
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    TL = mvarTL
End Property



Public Property Let SurfaceArea(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SurfaceArea = 5
    mvarSurfaceArea = vdata
End Property


Public Property Get SurfaceArea() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SurfaceArea
    SurfaceArea = mvarSurfaceArea
End Property



Public Property Let Quantity(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quantity = 5
    mvarQuantity = vdata
End Property


Public Property Get Quantity() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quantity
    Quantity = mvarQuantity
End Property



Public Property Let Parent(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Parent = 5
    mvarParent = vdata
End Property


Public Property Get Parent() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parent
    Parent = mvarParent
End Property



Public Property Let Key(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vdata
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property



Public Property Let SelectedImage(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SelectedImage = 5
    mvarSelectedImage = vdata
End Property


Public Property Get SelectedImage() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedImage
    SelectedImage = mvarSelectedImage
End Property



Public Property Let Image(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Image = 5
    mvarImage = vdata
End Property


Public Property Get Image() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Image
    Image = mvarImage
End Property



Public Property Let Description(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vdata
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let Datatype(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Datatype = 5
    mvarDatatype = vdata
End Property


Public Property Get Datatype() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Datatype
    Datatype = mvarDatatype
End Property



Public Property Let Cost(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cost = 5
    mvarCost = vdata
End Property


Public Property Get Cost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cost
    Cost = mvarCost
End Property



Public Property Let Custom(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Custom = 5
    mvarCustom = vdata
End Property


Public Property Get Custom() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Custom
    Custom = mvarCustom
End Property



Public Property Let DR(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR = 5
    mvarDR = vdata
End Property


Public Property Get DR() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DR
    DR = mvarDR
End Property



Public Property Let HitPoints(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HitPoints = 5
    mvarHitPoints = vdata
End Property


Public Property Get HitPoints() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HitPoints
    HitPoints = mvarHitPoints
End Property

Public Property Let Ruggedized(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ruggedized = 5
    mvarRuggedized = vdata
End Property


Public Property Get Ruggedized() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ruggedized
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

    If mvarRuggedized Then
        sPrint1 = "ruggedized "
    Else
        sPrint1 = ""
    End If

    
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
Dim B As Single
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
            B = 5
        Case 8
            B = 2
        Case 9
            B = 1
        Case 10
            B = 0.25
        Case Is >= 11
            B = 0.1
        Case Else
            B = 0
    End Select
    GWeight = GWeight + GuidanceMatrix(mvarMatrixPos2).WeightMod + B
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
Dim H As Single
Dim s As Single
Dim C As Single
Dim SizeMod As Single
Dim Divisor As Single
On Error Resume Next
D = mvarDiameter
C = D ^ 3
If mvarTL <= 5 Then
    x = 0.375
    H = 0.25
ElseIf mvarTL = 6 Then
    x = 0.75
    H = 0.25
ElseIf mvarTL = 7 Then
    x = 1
    H = 0.375
ElseIf mvarTL = 8 Then
    x = 3
    H = 1
ElseIf mvarTL >= 9 Then
    x = 4.5
    H = 1.5
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
    mvarDamage1 = ConvertDamage(D * s * H * SizeMod / Divisor) & Suffix1

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
    mvarDamage1 = ConvertDamage(D * s * H * SizeMod / Divisor) & Suffix1 & Suffix2
    
Case "CHEM", "FASCAM", "ICM", "SICM", "SATNUC"
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
    mvarDamage1 = ConvertDamage(AmmoMatrix(i).Multiplier * D * D * SizeMod)

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
          
         V = Round(mvarSpeed / 10, 2)
    
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
    Veh(Parent).StatsUpdate
End If
End Sub
