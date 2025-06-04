VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWeaponStoneBoltThrower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
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
Private mvarRuggedized As Boolean
Private mvarTL As Integer
Private mvarVolume As Double
Private mvarWeight As Double
Private mvarStrength As Long
Private mvarMagazineCapacity As Long
Private mvarAmmunitionType As String
Private mvarMalfunction As String
Private mvarTypeDamage As String
Private mvarDamage As String
Private mvarhalfDamage As Double
Private mvarMaxRange As Double
Private mvarMinRange As Single
Private mvarAccuracy As Long
Private mvarSnapShot As Long
Private mvarRoF As String
Private mvarWPS As Single
Private mvarVPS As Single
Private mvarCPS As Single
Private mvarShots As Long
Private mvarLoaders As Long
Private mvarQuality As String
Private mvarMechanism As String
Private mvarCustomDescription As String
Private mvarMount As String
Private mvarDirection As String
Private mvarLocation As String
Private mvarComment As String
Private mvarCName As String
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
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrintOutput = 5
    mvarPrintOutput = vdata
End Property


Public Property Get PrintOutput() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrintOutput
    PrintOutput = mvarPrintOutput
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



Public Property Let Location(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Location = 5
    mvarLocation = vdata
End Property


Public Property Get Location() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Location
    Location = mvarLocation
End Property





Public Function LocationCheck() As Boolean
Dim TempCheck As Boolean
Dim InstallPoint As Integer
InstallPoint = Veh.Components(mvarParent).Datatype


        If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = equipmentPod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = Module) Then
            TempCheck = True
        Else
            InfoPrint 1, "Weapons must be placed in Body, Superstructure, Pod, equipment Pod,Turret, Popturret, Arm, Wing, Open Mount, Leg or Module."
            TempCheck = False
        End If

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

Public Property Let Direction(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Direction = 5
    mvarDirection = vdata
End Property


Public Property Get Direction() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Direction
    Direction = mvarDirection
End Property

Public Property Let Mount(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Mount = 5
    mvarMount = vdata
End Property


Public Property Get Mount() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Mount
    Mount = mvarMount
End Property

Public Property Let Mechanism(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Mechanism = 5
    mvarMechanism = vdata
End Property


Public Property Get Mechanism() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Mechanism
    Mechanism = mvarMechanism
End Property



Public Property Let Quality(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quality = 5
    mvarQuality = vdata
End Property


Public Property Get Quality() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quality
    Quality = mvarQuality
End Property



Public Property Let Loaders(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Loaders = 5
    mvarLoaders = vdata
End Property


Public Property Get Loaders() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Loaders
    Loaders = mvarLoaders
End Property



Public Property Let Shots(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Shots = 5
    mvarShots = vdata
End Property


Public Property Get Shots() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Shots
    Shots = mvarShots
End Property



Public Property Let CPS(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPS = 5
    mvarCPS = vdata
End Property


Public Property Get CPS() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CPS
    CPS = mvarCPS
End Property



Public Property Let VPS(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VPS = 5
    mvarVPS = vdata
End Property


Public Property Get VPS() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VPS
    VPS = mvarVPS
End Property



Public Property Let WPS(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WPS = 5
    mvarWPS = vdata
End Property


Public Property Get WPS() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WPS
    WPS = mvarWPS
End Property



Public Property Let rof(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RoF = 5
    mvarRoF = vdata
End Property


Public Property Get rof() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RoF
    rof = mvarRoF
End Property



Public Property Let SnapShot(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SnapShot = 5
    mvarSnapShot = vdata
End Property


Public Property Get SnapShot() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SnapShot
    SnapShot = mvarSnapShot
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



Public Property Let Damage(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Damage = 5
    mvarDamage = vdata
End Property


Public Property Get Damage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Damage
    Damage = mvarDamage
End Property



Public Property Let TypeDamage(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TypeDamage = 5
    mvarTypeDamage = vdata
End Property


Public Property Get TypeDamage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TypeDamage
    TypeDamage = mvarTypeDamage
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



Public Property Let AmmunitionType(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AmmunitionType = 5
    mvarAmmunitionType = vdata
End Property


Public Property Get AmmunitionType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AmmunitionType
    AmmunitionType = mvarAmmunitionType
End Property



Public Property Let MagazineCapacity(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MagazineCapacity = 5
    mvarMagazineCapacity = vdata
End Property


Public Property Get MagazineCapacity() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MagazineCapacity
    MagazineCapacity = mvarMagazineCapacity
End Property



Public Property Let Strength(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Strength = 5
    mvarStrength = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If mvarDatatype = BoltThrower Then
        If mvarMechanism = "spring-powered" Then
            If mvarStrength > 75 Then
                mvarStrength = 75
                InfoPrint 1, "Max Practical Strength for Spring-Powered Bolt Thrower is 75"
            End If
        Else
            If mvarStrength > 150 Then
                mvarStrength = 150
                InfoPrint 1, "Max Practical Strength for Torsion-Powered Bolt Thrower is 150"
            End If
        End If
    ElseIf mvarDatatype = RepeatingBoltThrower Then
        If mvarStrength > Val(mvarTL) * 5 Then
            mvarStrength = Val(mvarTL) * 5
            InfoPrint 1, "Repeating Bolt Throwers Strength cannot exceed TL * 5"
        End If
    Else
        If mvarMechanism = "spring-powered" Then
            If mvarStrength > 50 Then
                mvarStrength = 50
                InfoPrint 1, "Maximum Strength for Spring-Powered Stone Thrower is 50"
            End If
        ElseIf mvarMechanism = "torsion-powered" Then
            If mvarStrength > 500 Then
                mvarStrength = 500
                InfoPrint 1, "Maximum Strength for Torsion-Powered Stone Throwers = 500"
            End If
        End If
    End If
    
End Property


Public Property Get Strength() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Strength
    Strength = mvarStrength
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
    mvarTL = vdata
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
mvarStrength = 50
mvarMechanism = "spring-powered" 'others are Torsion-Powered and Counterweight"
mvarQuality = "normal"
mvarMount = "normal"
mvarDirection = "front"

End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case StoneThrower
        
        mvarAmmunitionType = "stones"
    Case BoltThrower
       
        mvarAmmunitionType = "bolts"
    Case RepeatingBoltThrower
        
        mvarMagazineCapacity = 5
        mvarAmmunitionType = "bolts"
End Select

End Sub



Public Sub StatsUpdate()
    mvarZZInit = 1
    Dim sPrint1 As String
    Dim sPrint2 As String
    Dim sPrintPlural As String
    Dim sPrintPlural2 As String
    Dim sPrintPlural3 As String
    Dim sPrintDirection As String
    Dim QRugMod As Single
    Dim RugHitMod As Long

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

    mvarMalfunction = GetMalfunction
    mvarTypeDamage = GetTypeDamage
    mvarDamage = GetDamage
    mvarhalfDamage = GetHalfDamage
    mvarMaxRange = GetMaxRange
    mvarMinRange = GetMinRange
    mvarAccuracy = GetAccuracy
    mvarWeight = GetWeight
    mvarVolume = GetVolume
    mvarSnapShot = GetSnapShot
    mvarRoF = GetRoF
    mvarCost = GetCost
    mvarWPS = GetWeightPerShot
    mvarVPS = GetVolumePerShot
    mvarCPS = GetCostPerShot
    mvarShots = GetShots
    mvarLoaders = GetLoaders
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea) * RugHitMod
    
    
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
    sPrint1 = sPrint1 & "ST " & mvarStrength & " "
    
    If mvarMount <> "normal" Then
        sPrint2 = sPrint2 + ", " + mvarMount
    End If
    If mvarQuality <> "normal" Then sPrint2 = sPrint2 + ", " + mvarQuality + " construction"
    
    sPrint2 = sPrint2 + ", " + mvarMechanism
    
    If mvarMagazineCapacity <> 0 Then
        sPrint2 = sPrint2 + ", " + Format(mvarMagazineCapacity) + " round magazine"
    End If
    sPrintDirection = StrConv(Left(mvarDirection, 1), vbUpperCase)
    
    If mvarQuantity > 1 Then
        sPrintPlural = "s"
        sPrintPlural2 = " each"
        sPrintPlural3 = " total of "
    Else
        sPrintPlural = ""
        sPrintPlural2 = ""
        sPrintPlural3 = ""
    End If
    
    mvarPrintOutput = NumericToString(mvarQuantity) + " TL" & mvarTL & " " + sPrint1 + mvarCustomDescription + sPrintPlural + sPrint2 _
        + " (" + mvarLocation + sPrintDirection + ", HP " & mvarHitPoints & sPrintPlural2 + ", " + sPrintPlural3 + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ")." + mvarComment
    
    'note: cost, malf, and accuracy modifiers for Cheap, Fine and Very Fine quality are calced in the functions below
End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub

Private Function GetMalfunction() As String
Dim TempMalf As String

If mvarDatatype = RepeatingBoltThrower Then
    TempMalf = "16"
Else
    TempMalf = "Crit."
End If

'get modifier for Cheap, Fine and Very Fine quality
If mvarQuality = "cheap" Then
    TempMalf = DecreaseMalf(TempMalf)
ElseIf mvarQuality = "fine (reliable)" Then
    TempMalf = IncreaseMalf(TempMalf)
End If

GetMalfunction = TempMalf
End Function

Private Function GetTypeDamage() As String

If mvarDatatype = StoneThrower Then
    GetTypeDamage = "Cr."
Else
    GetTypeDamage = "Imp."
End If

End Function

Private Function GetDamage() As String
Dim dam1 As Long
Dim dam2 As Long
Dim i As Long

Dim WeaponST As String

WeaponST = Str(mvarStrength)

'find base damage using the lookup table
If mvarDatatype = StoneThrower Then
    If WeaponST <= 100 Then
        dam1 = StoneBoltDamageMatrix(WeaponST).Swing1
        dam2 = StoneBoltDamageMatrix(WeaponST).Swing2
    Else
        i = Int(WeaponST / 10)
        dam1 = 3 + i
    End If
Else
    If WeaponST <= 100 Then
        dam1 = StoneBoltDamageMatrix(WeaponST).Thrust1
        dam2 = StoneBoltDamageMatrix(WeaponST).Thrust2
    Else
        i = Int(WeaponST / 10)
        dam1 = 1 + i
    End If
End If

'apply modifiers depending on Stone or Bolt thrower
If mvarDatatype = StoneThrower Then
    If mvarMechanism = "counterweight" Then
        dam2 = dam2 + (1 * dam1)
    End If
Else
    If mvarStrength <= 20 Then
        dam2 = dam2 + 4
    Else
        dam1 = dam1 + 1
    End If
End If

'create merged damage string
If dam2 = 0 Then
    GetDamage = dam1 & "d"
ElseIf dam2 < 0 Then
    GetDamage = dam1 & "d" & " -" & Abs(dam2)
Else
    GetDamage = dam1 & "d" & " +" & dam2
End If

End Function

Private Function GetHalfDamage() As Single 'in yards
Dim TempHalfDamage As Single

If mvarDatatype = StoneThrower Then
    If mvarStrength <= 20 Then
        TempHalfDamage = mvarStrength * 10
    Else
        TempHalfDamage = 190 + (mvarStrength / 2)
    End If
    'counterweight stone throwers must be modified to have shorter range
    If mvarMechanism = "counterweight" Then TempHalfDamage = TempHalfDamage * 0.75
Else
    If mvarStrength <= 20 Then
        TempHalfDamage = mvarStrength * 20
    Else
        TempHalfDamage = 380 + mvarStrength
    End If
End If

GetHalfDamage = TempHalfDamage
End Function

Private Function GetMaxRange() As Double 'in yards
GetMaxRange = 1.25 * mvarhalfDamage
End Function

Private Function GetMinRange() As Single 'in yards
If mvarMechanism = "counterweight" Then
    GetMinRange = mvarMaxRange / 4
Else
    GetMinRange = 0
End If
End Function

Private Function GetAccuracy() As Long
Dim TempAccuracy As Long

If mvarDatatype = RepeatingBoltThrower Then
    TempAccuracy = 5
ElseIf mvarDatatype = BoltThrower Then
    TempAccuracy = 6
Else
    If mvarMechanism = "counterweight" Then
        TempAccuracy = 1
    Else
        TempAccuracy = 2
    End If
End If

'get modifier for Cheap, Fine and Very Fine quality
If mvarQuality = "cheap" Then
    TempAccuracy = TempAccuracy - 1
ElseIf mvarQuality = "fine (accurate)" Then
    TempAccuracy = TempAccuracy + 1
ElseIf mvarQuality = "very fine (accurate)" Then
    TempAccuracy = TempAccuracy + 2
End If

GetAccuracy = TempAccuracy
End Function

Private Function GetWeight() As Double
Dim ST As Long ' strength
Dim P As Single 'power mechanism modifier
Dim T As Single 'tech level modifier
Dim M As Single 'magazine capacity modifier
Dim D As Single 'datatype modifier

ST = mvarStrength
M = 1 + (0.05 * mvarMagazineCapacity)

If mvarMechanism = "spring-powered" Then
    P = 1
ElseIf mvarMechanism = "torsion-powered" Then
    P = 0.8
Else
    P = 0.5
End If

If mvarTL <= 4 Then
    T = 1
ElseIf mvarTL = 5 Then
    T = 0.75
ElseIf mvarTL = 6 Then
    T = 0.6
Else
    T = 0.5
End If

If mvarDatatype = StoneThrower Then
    D = 0.25
Else
    D = 0.1
End If

GetWeight = Round(ST * ST * P * T * M * D, 5)
End Function

Private Function GetVolume() As Double
If mvarMount = "normal" Then
    GetVolume = mvarWeight / 50
Else
    GetVolume = mvarWeight / 20 'concealed weapons take up more space
End If

End Function

Private Function GetSnapShot() As Long
Dim TSS As Long

Const Min = 25 'minimum value for counterweight stone throwers

If mvarWeight < 15 Then
    TSS = 12
ElseIf mvarWeight <= 25 Then
    TSS = 15
ElseIf mvarWeight <= 400 Then
    TSS = 20
Else
    TSS = 25
End If


If mvarDatatype = StoneThrower Then
    TSS = TSS + 5 'add five to all stonethrowers
    If mvarMechanism = "counterweight" Then
        If TSS < Min Then TSS = Min
    End If
End If

GetSnapShot = TSS
End Function

Private Function GetRoF() As String
Dim TempRoF As Double

If mvarDatatype = RepeatingBoltThrower Then
    TempRoF = Round(2 * Sqr(mvarStrength), 0)
ElseIf mvarMechanism = "counterweight" Then
    TempRoF = Round(10 * Sqr(mvarStrength), 0)
Else
    TempRoF = Round(5 * Sqr(mvarStrength), 0)
End If

GetRoF = "1/" & TempRoF

End Function

Private Function GetCost() As Double
Dim M As Integer 'weapon type modifier
Dim P As Integer 'power mechanism modifier
Dim TempCost As Double

If mvarMechanism = "spring-powered" Then
    M = 1
Else
    M = 2
End If

If mvarDatatype = RepeatingBoltThrower Then
    P = 2
Else
    P = 1
End If

If mvarWeight < 100 Then
    TempCost = 25 * mvarWeight * M * P
Else
    TempCost = (2400 + mvarWeight) * M * P
End If

'get modifier for Cheap, Fine and Very Fine quality
If mvarQuality = "cheap" Then
    TempCost = TempCost / 2
ElseIf mvarQuality = "fine (accurate)" Then
    TempCost = TempCost * 5
ElseIf mvarQuality = "very fine (accurate)" Then
    TempCost = TempCost * 30
ElseIf mvarQuality = "fine (reliable)" Then
    TempCost = TempCost * 5
End If

GetCost = Round(TempCost, 2)
End Function

Private Function GetWeightPerShot() As Single
    Dim sngTemp As Single
    
    If mvarDatatype = StoneThrower Then
        sngTemp = mvarStrength / 10
    Else
        sngTemp = mvarStrength * mvarStrength / 1600
    End If
    
    GetWeightPerShot = Round(sngTemp, 5)
End Function

Private Function GetVolumePerShot() As Single

GetVolumePerShot = mvarWPS / 50
End Function

Private Function GetCostPerShot() As Single
Const Min = 2
Dim TempCost As Single

If mvarDatatype = StoneThrower Then
    GetCostPerShot = mvarWPS * 0.5
Else
    TempCost = mvarWPS * 2
    If TempCost < Min Then TempCost = Min
    GetCostPerShot = TempCost
End If
End Function

Private Function GetShots() As Long
'number of shots the weapon has ready to fire.
'stone throwers and bolt throwers have 1.  Repeating bolt throwers use their
'magaizine capacity
If mvarDatatype = RepeatingBoltThrower Then
    GetShots = mvarMagazineCapacity
Else
    GetShots = 1
End If

End Function

Private Function GetLoaders() As Long
'TODO inquire to Pulver.  There is a discrepancy in Crew versus Loaders
' of mech weapons and guns.  One asks for totalcrew while other asks for loaders. Why?
Dim Divisor As Integer
Dim TempCrew As Single
Const GunnerST = 12 ' the arbitary value for the gunner's strength

If mvarMechanism = "counterweight" Then
    Divisor = 20
Else
    Divisor = 40
End If

If GunnerST + 4 <= mvarStrength Then
    TempCrew = (mvarStrength / Divisor) - 1
    GetLoaders = RoundUP(TempCrew)
Else 'no loaders required
    GetLoaders = 0
End If

End Function
