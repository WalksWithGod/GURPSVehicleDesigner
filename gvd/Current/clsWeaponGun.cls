VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWeaponGun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' optional flags
Private mvarCustom As Boolean

' stats
Private mvarHitPoints As Double
Private mvarDR As Long
Private mvarCost As Double
Private mvarDatatype As Integer
Private mvarDescription As String
Private mvarSurfaceArea As Double
Private mvarRuggedized As Boolean
Private mvarTL As Integer
Private mvarVolume As Double
Private mvarWeight As Double
Private mvarCustomDescription As String ' this is not needed in 2.0.

' treeview crap
Private mvarImage As Integer
Private mvarSelectedImage As Integer
Private mvarKey As String

' node specific
Private mvarParent As String
Private mvarQuantity As Integer

' this may wind up just being a subtype of class
Private mvarTechnology As String
    ' i.e. smoothbore, rifled, electromag, gravitic
    
' gun specific options
Private mvarMount As String     ' this should be a seperate component
Private mvarBoreSize As Single
Private mvarRecoiless As Boolean  ' these modifiers effect cost/weight/ and such but arent needed at runtime.
                                  ' however, they SHOULD have descriptions that say "+3 to blah" or whatever where appropriate
Private mvarQuality As String
Private mvarPowerOption As String
Private mvarBarrel As String         '
Private mvarAdvancedOption As String
Private mvarAmmunitionType As String
Private mvarCylinders As Long
Private mvarElectric As Boolean
Private mvarCarriage As Boolean
Private mvarBoxMagazine As Boolean
Private mvarDirection As String
Private mvarLocation As String
Private mvarComment As String
Private mvarCName As String
Private mvarTempWPS As Double

' gun specific stats used at runtime
Private mvarMalfunction As String
Private mvarReliable As Boolean
Private mvarTypeDamage1 As String
Private mvarTypeDamage2 As String
Private mvarDamage1 As String
Private mvarDamage2 As String
Private mvarBurstRadius As Long
Private mvarKEDamage As Double
Private mvarhalfDamage As Double
Private mvarMaxRange As Double
Private mvarAccuracy As Long
Private mvarSnapShot As Long
Private mvarDXPlusSkill As Long
Private mvarsRoF As String 'one holds the string conversion of the Rof
Private mvardRof As Single 'one holds the Numeric value for rate of fire
Private mvarWPS As Double
Private mvarVPS As Double
Private mvarCPS As Double
Private mvarPowerReqt As Double
Private mvarLoaders As Long
Private mvarShots As String

Private mvarPrintOutput As String
Private ammoarray() As String
Private rofarray() As String
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

'Syntax: Debug.Print X.PrintOutput
    PrintOutput = mvarPrintOutput
End Property


Public Property Let CName(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CName = 5
    mvarCName = vdata
End Property


Public Property Get CName() As String

'Syntax: Debug.Print X.CName
    CName = mvarCName
End Property



Public Property Let Comment(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Comment = 5
    mvarComment = vdata
End Property


Public Property Get Comment() As String

'Syntax: Debug.Print X.Comment
    Comment = mvarComment
End Property



Public Property Let Location(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Location = 5
    mvarLocation = vdata
End Property


Public Property Get Location() As String

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


Public Property Let Carriage(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Carriage = 5
    mvarCarriage = vdata
End Property


Public Property Get Carriage() As Boolean

'Syntax: Debug.Print X.Carriage
    Carriage = mvarCarriage
End Property

Public Property Let BoxMagazine(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BoxMagazine = 5
    mvarBoxMagazine = vdata
End Property


Public Property Get BoxMagazine() As Boolean

'Syntax: Debug.Print X.BoxMagazine
    BoxMagazine = mvarBoxMagazine
End Property



Public Property Let DXPlusSkill(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DXPlusSkill = 5
    mvarDXPlusSkill = vdata
End Property


Public Property Get DXPlusSkill() As Long

'Syntax: Debug.Print X.DXPlusSkill
    DXPlusSkill = mvarDXPlusSkill
End Property


Public Property Let Shots(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Shots = 5
    mvarShots = vdata
End Property


Public Property Get Shots() As String

'Syntax: Debug.Print X.Shots
    Shots = mvarShots
End Property



Public Property Let Loaders(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Loaders = 5
    mvarLoaders = vdata
End Property


Public Property Get Loaders() As Long

'Syntax: Debug.Print X.Loaders
    Loaders = mvarLoaders
End Property



Public Property Let PowerReqt(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PowerReqt = 5
    mvarPowerReqt = vdata
End Property


Public Property Get PowerReqt() As Double

'Syntax: Debug.Print X.PowerReqt
    PowerReqt = mvarPowerReqt
End Property



Public Property Let CPS(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPS = 5
    mvarCPS = vdata
End Property


Public Property Get CPS() As Double

'Syntax: Debug.Print X.CPS
    CPS = mvarCPS
End Property



Public Property Let VPS(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VPS = 5
    mvarVPS = vdata
End Property


Public Property Get VPS() As Double

'Syntax: Debug.Print X.VPS
    VPS = mvarVPS
End Property



Public Property Let WPS(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WPS = 5
    mvarWPS = vdata
End Property


Public Property Get WPS() As Double

'Syntax: Debug.Print X.WPS
    WPS = mvarWPS
End Property



Public Property Let sRoF(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sRoF = 5
    mvarsRoF = vdata
End Property


Public Property Get sRoF() As String

'Syntax: Debug.Print X.sRoF
    sRoF = mvarsRoF
End Property

Public Property Let dRoF(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dRoF = 5
    mvardRof = vdata
End Property


Public Property Get dRoF() As Single

'Syntax: Debug.Print X.dRoF
    dRoF = mvardRof
End Property


Public Property Let SnapShot(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SnapShot = 5
    mvarSnapShot = vdata
End Property


Public Property Get SnapShot() As Long

'Syntax: Debug.Print X.SnapShot
    SnapShot = mvarSnapShot
End Property



Public Property Let Accuracy(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Accuracy = 5
    mvarAccuracy = vdata
End Property


Public Property Get Accuracy() As Long

'Syntax: Debug.Print X.Accuracy
    Accuracy = mvarAccuracy
End Property



Public Property Let MaxRange(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxRange = 5
    mvarMaxRange = vdata
End Property


Public Property Get MaxRange() As Double

'Syntax: Debug.Print X.MaxRange
    MaxRange = mvarMaxRange
End Property



Public Property Let halfDamage(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.halfDamage = 5
    mvarhalfDamage = vdata
End Property


Public Property Get halfDamage() As Double

'Syntax: Debug.Print X.halfDamage
    halfDamage = mvarhalfDamage
End Property



Public Property Let Damage1(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Damage1 = 5
    mvarDamage1 = vdata
End Property


Public Property Get Damage1() As String

'Syntax: Debug.Print X.Damage1
    Damage1 = mvarDamage1
End Property

Public Property Let Damage2(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Damage2 = 5
    mvarDamage2 = vdata
End Property


Public Property Get Damage2() As String

'Syntax: Debug.Print X.Damage2
    Damage2 = mvarDamage2
End Property

Public Property Let BurstRadius(ByVal vdata As Long)
    mvarBurstRadius = vdata
End Property

Public Property Get BurstRadius() As Long
    BurstRadius = mvarBurstRadius
End Property

Public Property Let TypeDamage1(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TypeDamage1 = 5
    mvarTypeDamage1 = vdata
End Property


Public Property Get TypeDamage1() As String

'Syntax: Debug.Print X.TypeDamage1
    TypeDamage1 = mvarTypeDamage1
End Property

Public Property Let TypeDamage2(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TypeDamage2 = 5
    mvarTypeDamage2 = vdata
End Property


Public Property Get TypeDamage2() As String

'Syntax: Debug.Print X.TypeDamage2
    TypeDamage2 = mvarTypeDamage2
End Property

Public Property Let KEDamage(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.KEDamage = 5
    mvarKEDamage = vdata
End Property


Public Property Get KEDamage() As Double

'Syntax: Debug.Print X.KEDamage
    KEDamage = mvarKEDamage
End Property

Public Property Let Reliable(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Reliable = 5
    mvarReliable = vdata
End Property


Public Property Get Reliable() As Boolean

'Syntax: Debug.Print X.Reliable
    Reliable = mvarReliable
End Property



Public Property Let Malfunction(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Malfunction = 5
    mvarMalfunction = vdata
End Property


Public Property Get Malfunction() As String

'Syntax: Debug.Print X.Malfunction
    Malfunction = mvarMalfunction
End Property



Public Property Let Electric(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Electric = 5
    mvarElectric = vdata
End Property


Public Property Get Electric() As Boolean

'Syntax: Debug.Print X.Electric
    Electric = mvarElectric
End Property



Public Property Let Cylinders(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cylinders = 5
    Dim I As Long
    Dim bValidRoF As Boolean
    
    mvarCylinders = vdata
    If mvarZZInit = 0 Then Exit Property
    
    'check for valid Rate of Fire
    If (mvarDatatype = ElectricGatling) Or (mvarDatatype = HeavyAutomatic) Or (mvarDatatype = lightAutomatic) Then
        FillRoFList
        'determine if the ammunition type that is set is still valid
        For I = 1 To UBound(rofarray)
            If rofarray(I) = mvardRof Then
                bValidRoF = True
                Exit For
            End If
        Next
        
        If bValidRoF = False Then
            InfoPrint 1, "RoF not supported at this Bore Size.  Resetting to RoF  to 3."
            mvardRof = 3
        End If
    End If
    
End Property


Public Property Get Cylinders() As Long
'Syntax: Debug.Print X.Cylinders
    Cylinders = mvarCylinders
End Property

Public Property Let AmmunitionType(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AmmunitionType = 5
    mvarAmmunitionType = vdata
End Property


Public Property Get AmmunitionType() As String
'Syntax: Debug.Print X.AmmunitionType
    AmmunitionType = mvarAmmunitionType
End Property


Public Property Let AdvancedOption(ByVal vdata As String)
'Syntax: X.AdvancedOption = 5
    mvarAdvancedOption = vdata
End Property


Public Property Get AdvancedOption() As String

'Syntax: Debug.Print X.AdvancedOption
    AdvancedOption = mvarAdvancedOption
End Property



Public Property Let Barrel(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Barrel = 5
    mvarBarrel = vdata
End Property


Public Property Get Barrel() As String

'Syntax: Debug.Print X.Barrel
    Barrel = mvarBarrel
End Property



Public Property Let PowerOption(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PowerOption = 5
    mvarPowerOption = vdata
End Property


Public Property Get PowerOption() As String

'Syntax: Debug.Print X.PowerOption
    PowerOption = mvarPowerOption
End Property



Public Property Let Recoiless(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Recoiless = 5
    mvarRecoiless = vdata
End Property


Public Property Get Recoiless() As Boolean

'Syntax: Debug.Print X.Recoiless
    Recoiless = mvarRecoiless
End Property



Public Property Let Technology(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Technology = 5
    mvarTechnology = vdata
End Property


Public Property Get Technology() As String

'Syntax: Debug.Print X.Technology
    Technology = mvarTechnology
End Property



Public Property Let BoreSize(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BoreSize = 5
Dim I As Integer
Dim bValidAmmoType As Boolean
Dim bValidRoF As Boolean

    mvarBoreSize = vdata
    If mvarZZInit = 0 Then Exit Property
    
    FillAmmunitionList
    'determine if the ammunition type that is set is still valid
    For I = 1 To UBound(ammoarray)
        If ammoarray(I) = mvarAmmunitionType Then
            bValidAmmoType = True
            Exit For
        End If
    Next
    
    If bValidAmmoType = False Then
        InfoPrint 1, "Ammunition type not supported at this Bore Size.  Resetting to Solid ammunition type."
        mvarAmmunitionType = "Solid"
    End If

    'check for valid Rate of Fire
    If (mvarDatatype = ElectricGatling) Or (mvarDatatype = HeavyAutomatic) Or (mvarDatatype = lightAutomatic) Then
        FillRoFList
        'determine if the ammunition type that is set is still valid
        For I = 1 To UBound(rofarray)
            If Val(rofarray(I)) = mvardRof Then
                bValidRoF = True
                Exit For
            End If
        Next
        
        If bValidRoF = False Then
            InfoPrint 1, "RoF not supported at this Bore Size.  Resetting to RoF  to 3."
            mvardRof = 3
        End If
    End If
End Property


Public Property Get BoreSize() As Single

'Syntax: Debug.Print X.BoreSize
    BoreSize = mvarBoreSize
End Property



Public Property Let Mount(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Mount = 5
    mvarMount = vdata
End Property


Public Property Get Mount() As String

'Syntax: Debug.Print X.Mount
    Mount = mvarMount
End Property



Public Property Let Quality(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quality = 5
    mvarQuality = vdata
End Property


Public Property Get Quality() As String

'Syntax: Debug.Print X.Quality
    Quality = mvarQuality
End Property



Public Property Let CustomDescription(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CustomDescription = 5
    mvarCustomDescription = vdata
End Property


Public Property Get CustomDescription() As String

'Syntax: Debug.Print X.CustomDescription
    CustomDescription = mvarCustomDescription
End Property

Public Property Let Direction(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Direction = 5
    mvarDirection = vdata
End Property


Public Property Get Direction() As String

'Syntax: Debug.Print X.Direction
    Direction = mvarDirection
End Property

Public Property Let Weight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Weight = 5
    mvarWeight = vdata
End Property


Public Property Get Weight() As Double

'Syntax: Debug.Print X.Weight
    Weight = mvarWeight
End Property



Public Property Let Volume(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Volume = 5
    mvarVolume = vdata
End Property


Public Property Get Volume() As Double

'Syntax: Debug.Print X.Volume
    Volume = mvarVolume
End Property



Public Property Let TL(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TL = 5
    
    mvarTL = vdata
    
End Property


Public Property Get TL() As Integer

'Syntax: Debug.Print X.TL
    TL = mvarTL
End Property



Public Property Let SurfaceArea(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SurfaceArea = 5
    mvarSurfaceArea = vdata
End Property


Public Property Get SurfaceArea() As Double

'Syntax: Debug.Print X.SurfaceArea
    SurfaceArea = mvarSurfaceArea
End Property



Public Property Let Quantity(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quantity = 5
    mvarQuantity = vdata
End Property


Public Property Get Quantity() As Integer

'Syntax: Debug.Print X.Quantity
    Quantity = mvarQuantity
End Property



Public Property Let Parent(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Parent = 5
    mvarParent = vdata
End Property


Public Property Get Parent() As String

'Syntax: Debug.Print X.Parent
    Parent = mvarParent
End Property



Public Property Let Key(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vdata
End Property


Public Property Get Key() As String

'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property



Public Property Let SelectedImage(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SelectedImage = 5
    mvarSelectedImage = vdata
End Property


Public Property Get SelectedImage() As Integer

'Syntax: Debug.Print X.SelectedImage
    SelectedImage = mvarSelectedImage
End Property



Public Property Let Image(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Image = 5
    mvarImage = vdata
End Property


Public Property Get Image() As Integer

'Syntax: Debug.Print X.Image
    Image = mvarImage
End Property



Public Property Let Description(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vdata
End Property


Public Property Get Description() As String

'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let Datatype(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Datatype = 5
    mvarDatatype = vdata
End Property


Public Property Get Datatype() As Integer

'Syntax: Debug.Print X.Datatype
    Datatype = mvarDatatype
End Property



Public Property Let Cost(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cost = 5
    mvarCost = vdata
End Property


Public Property Get Cost() As Double

'Syntax: Debug.Print X.Cost
    Cost = mvarCost
End Property



Public Property Let Custom(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Custom = 5
    mvarCustom = vdata
End Property


Public Property Get Custom() As Boolean

'Syntax: Debug.Print X.Custom
    Custom = mvarCustom
End Property



Public Property Let DR(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR = 5
    mvarDR = vdata
End Property


Public Property Get DR() As Long

'Syntax: Debug.Print X.DR
    DR = mvarDR
End Property



Public Property Let HitPoints(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HitPoints = 5
    mvarHitPoints = vdata
End Property


Public Property Get HitPoints() As Double

'Syntax: Debug.Print X.HitPoints
    HitPoints = mvarHitPoints
End Property

Public Property Let Ruggedized(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ruggedized = 5
    mvarRuggedized = vdata
End Property


Public Property Get Ruggedized() As Boolean

'Syntax: Debug.Print X.Ruggedized
    Ruggedized = mvarRuggedized
End Property

Private Sub Class_Initialize()
 

' set the default properties
mvarCustom = False
TL = gVehicleTL
mvarQuantity = 1
mvarTechnology = "conventional smoothbore" 'others are Conventional Rifled, Electromag, Gravitic"
mvarQuality = "normal"
mvarMount = "normal"
mvarAmmunitionType = "Solid"
mvarRecoiless = False
mvarBarrel = "medium"
mvarReliable = False
mvarPowerOption = "normal"
mvarAdvancedOption = "none"
mvarBoxMagazine = False
mvarCarriage = False
mvarDirection = "front"
End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case MuzzleLoader
       
        mvarBoreSize = 30
        mvarCylinders = 1
    Case BreechLoader
        
        mvarBoreSize = 30
        mvarCylinders = 1
    Case ManualRepeater
       
        mvarBoreSize = 15
    Case Revolver
        
        mvarCylinders = 5
        mvarBoreSize = 10
    Case MechanicalGatling
        mvarDXPlusSkill = 10
        mvarBoreSize = 12
        mvarCylinders = 6
    Case SlowAutoloader
       
        mvarBoreSize = 15
    Case FastAutoloader
        
        mvarBoreSize = 15
    Case lightAutomatic
        mvardRof = 3
        mvarBoreSize = 15
    Case HeavyAutomatic
        
        mvarBoreSize = 20
        mvardRof = 3
    Case ElectricGatling
        mvarElectric = True
        mvarBoreSize = 30
        mvarCylinders = 6
        mvardRof = 3
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
    
    mvarLocation = GetLocation
    
    'determine if carriage is needed
    mvarCarriage = GetCarriage
    
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
    mvarMalfunction = GetMalfunction
    mvarKEDamage = GetDamage
    GetTypeDamages 'call sub to update Damages based on Ammunition Type
    mvarhalfDamage = GetHalfDamage
    mvarMaxRange = GetMaxRange
    'mvarMinRange = GetMinRange
    mvarAccuracy = GetAccuracy
    mvarWeight = GetWeight
    mvarVolume = GetVolume
    mvarSnapShot = GetSnapShot
    'get both Numeric and string conversion ROF's
    If (mvarDatatype <> lightAutomatic) And (mvarDatatype <> HeavyAutomatic) And (mvarDatatype <> ElectricGatling) Then
        GetRoF
    End If
    mvarCost = GetCost
    mvarWPS = GetWeightPerShot
    mvarVPS = GetVolumePerShot
    
    mvarPowerReqt = GetPowerReqt
    'make modifications based on any Advanced Options
    GetAdvancedOptionModifiers
    mvarCPS = GetCostPerShot
    mvarShots = GetShots
    mvarLoaders = GetLoaders * mvarQuantity
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea) * RugHitMod
    
    'update MaxRange and 1/2 Range depending on Ammunition Type
    '(Note: WPS,CPS,Acc updates are done within the actual functions called above)
    '(Note: MaxRange and 1/2 updates on the other hand must be done afterwards!!)
    UpdateRanges
    
    'cost, malf, and accuracy modifiers for Cheap, Fine and Very Fine quality are calced in the functions below
    'todo check that "mounting" option is taken into account
    'todo check that "recoiless" option is taken into account
    'todo check that advancedoption is taken into account
    
    '//update the cost,weight,volume, surface area and volume based on quantity and ruggedized options
    mvarCost = mvarCost * QRugMod
    mvarWeight = mvarWeight * QRugMod
    mvarVolume = mvarVolume * QRugMod
    
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    mvarPowerReqt = mvarPowerReqt * mvarQuantity
        
    'produce the print output
    If mvarRuggedized Then
        sPrint1 = "ruggedized "
    Else
        sPrint1 = ""
    End If
    
    If mvarMount <> "normal" Then
        sPrint2 = sPrint2 + ", " + mvarMount
    End If
    
    sPrint2 = sPrint2 + ", " & mvarTechnology
    
    If mvarPowerOption <> "normal" Then
        sPrint2 = sPrint2 + ", " & mvarPowerOption
    End If
    If mvarRecoiless Then
         sPrint2 = sPrint2 + ", recoiless"
    End If
    If mvarAdvancedOption <> "none" Then
        sPrint2 = sPrint2 + ", " & mvarAdvancedOption
    End If
    If mvarElectric Then
        sPrint2 = sPrint2 + ", electric loading"
    End If
    If mvarBoxMagazine Then
        sPrint2 = sPrint2 + ", box magazine"
    End If
    If mvarReliable Then
         sPrint2 = sPrint2 + ", reliable reputation"
    End If
    If mvarQuality <> "normal" Then sPrint2 = sPrint2 + ", " + mvarQuality + " construction"
    
    
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
    
    mvarPrintOutput = NumericToString(mvarQuantity) + " TL" & mvarTL & " " & sPrint1 & mvarBoreSize & "mm " & mvarCustomDescription + sPrintPlural + sPrint2 _
        + " (" + mvarLocation + sPrintDirection + ", HP " & mvarHitPoints & sPrintPlural2 + ", " + sPrintPlural3 + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ", " + Format(mvarPowerReqt, p_sFormat) + " kW)." + mvarComment
    

End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub

Private Function GetCarriage() As Boolean

' this is fairly lame since what if you want it to have a carriage anyway?
If mvarTL <= 5 Then
    If (mvarPowerOption <> "normal") And (mvarBoreSize > 120) Then
        GetCarriage = True
    ElseIf (mvarPowerOption = "normal") And (mvarBoreSize > 60) Then
        GetCarriage = True
    Else
        GetCarriage = False
    End If
Else
    GetCarriage = False
End If

End Function

Private Function GetMalfunction() As String
Dim TempMalf As String
Dim count As Integer 'holds the number of reliable traits.  >=2 increase Crit to Ver.

If mvarTL <= 3 Then
    TempMalf = 13
ElseIf mvarTL = 4 Then
    TempMalf = 14
ElseIf mvarTL = 5 Then
    TempMalf = 16
ElseIf mvarTL >= 6 Then
    TempMalf = "Crit."
End If

'breechloader CANNON built at TL4- is more likely to malfunction.  Decrease it by 3
If (mvarTL <= 4) And (mvarDatatype = BreechLoader) And (mvarBoreSize >= 20) Then
    TempMalf = DecreaseMalf(TempMalf)
    TempMalf = DecreaseMalf(TempMalf)
    TempMalf = DecreaseMalf(TempMalf)
End If

'check for other positive Malf modifiers (Todo: inquire to find out if these only apply to TL6+ weapons?)
count = 0 'init counter
If mvarElectric Then count = count + 1
If (mvarDatatype = FastAutoloader) Or (mvarDatatype = SlowAutoloader) Then
    If (mvarTechnology = "electromag") Or (mvarTechnology = "gravitic") Then
        count = count + 1
    End If
End If
If (mvarAdvancedOption = "liquid propellant") Or (mvarAdvancedOption = "caseless") Then count = count + 1
If mvarReliable Then count = count + 1
If count = 1 Then
    TempMalf = IncreaseMalf(TempMalf)
End If
If count >= 2 Then
    TempMalf = IncreaseMalf(TempMalf)
End If
'get modifier for Cheap, Fine and Very Fine quality
If mvarQuality = "cheap" Then
    TempMalf = DecreaseMalf(TempMalf)
ElseIf mvarQuality = "fine (reliable)" Then
    TempMalf = IncreaseMalf(TempMalf)
End If

GetMalfunction = TempMalf
End Function


Public Function FillAmmunitionList() As String()

ReDim ammoarray(1)
    ammoarray = mAddKey(ammoarray, "Solid") 'all weapons can use this
    If mvarBoreSize <= 2 Then
        ammoarray = mAddKey(ammoarray, "Needle") 'a sliver like projectile
        If (mvarTechnology = "gravitic") Or (mvarTechnology = "electromag") Then
            ammoarray = mAddKey(ammoarray, "Superwire")  'a superconducting monowire projectile only used for electromags or gravitic guns
        End If
    Else
        'add ammo for smallarms of >2mm
        If (mvarBoreSize > 2) And (mvarBoreSize < 20) Then
            ammoarray = mAddKey(ammoarray, "AP") 'armor piercing bullet
            ammoarray = mAddKey(ammoarray, "HP")  'hollow point, a bullet designed to expand in flesh
            ammoarray = mAddKey(ammoarray, "Plastic")  ' a plastic or rubber bullet
            ammoarray = mAddKey(ammoarray, "APS")  'an armor piercing, saboted bullet
            ammoarray = mAddKey(ammoarray, "APSHD")  'an advanced, hyperdense, saboted bullet
        End If
        'add ammo for small arms and cannons
        
        If (mvarBoreSize >= 10) And (mvarBoreSize < 20) Then
            ammoarray = mAddKey(ammoarray, "Shotshell")  ' multiple shotgun pellets
            ammoarray = mAddKey(ammoarray, "LE")  'low explosive, a fused round filled with black powder
            ammoarray = mAddKey(ammoarray, "HE")  'high explosive, a modern explosive, fragmenting round
            ammoarray = mAddKey(ammoarray, "HEC")  'HE Concussion, with bit blast and little fragmentation
            ammoarray = mAddKey(ammoarray, "HEPF")  'HE proximity fused round for anti aircraft fire
            ammoarray = mAddKey(ammoarray, "SAPLE")  'LE fused to go off after piercing armor
            ammoarray = mAddKey(ammoarray, "SAPHE")  'HE fused to go off after piercing armor
            ammoarray = mAddKey(ammoarray, "APEX")  'armor piercing explosive cannon round
            ammoarray = mAddKey(ammoarray, "Beehive")  'high-tech shrapnel using anti-personnel darts
            ammoarray = mAddKey(ammoarray, "CHEM")  'chemical, a gas or smoke-filled round
            If (mvarTL = 8) And (mvarBoreSize >= 15) Then
                ammoarray = mAddKey(ammoarray, "HEAT")  'a shaped-charged, armor piercing "anti-tank" round
                ammoarray = mAddKey(ammoarray, "HEDP")  ' HE Dual Purpose, a fragmentation and HEAT round
                ammoarray = mAddKey(ammoarray, "HESH")  ' HE Squash-Head, plastic explosive built to defeat armor
            ElseIf (mvarTL >= 9) And (mvarBoreSize >= 10) Then
                ammoarray = mAddKey(ammoarray, "HEAT")  'a shaped-charged, armor piercing "anti-tank" round
                ammoarray = mAddKey(ammoarray, "HEDP")  ' HE Dual Purpose, a fragmentation and HEAT round
                ammoarray = mAddKey(ammoarray, "HESH")  ' HE Squash-Head, plastic explosive built to defeat armor
            End If
        'add ammo for just cannons
        ElseIf mvarBoreSize >= 20 Then
            ammoarray = mAddKey(ammoarray, "Chainshot")  'two balls linked by a chain
            ammoarray = mAddKey(ammoarray, "APC")  'an armor piercing cannon round
            ammoarray = mAddKey(ammoarray, "API")  'an AP cannon round with added incendiary material
            ammoarray = mAddKey(ammoarray, "APCR")  'superior AP cannon round with a rigid tungsten core
            ammoarray = mAddKey(ammoarray, "Baton")  'a non lethal anti-riot round
            ammoarray = mAddKey(ammoarray, "APDS")  'improved APCR cannon round with discrading sabot
            ammoarray = mAddKey(ammoarray, "APDU")  'improved APCR with depleted uranium core
            ammoarray = mAddKey(ammoarray, "APDSDU")  'improved APDS with dense, depleted uranium core
            ammoarray = mAddKey(ammoarray, "APFSDS")  'hypervelocity, fin stabilized APDS dart
            ammoarray = mAddKey(ammoarray, "APFSDSDU")  'hypervelocity, fin stabilized APDSDU dart
            ammoarray = mAddKey(ammoarray, "APDSHD")  'ultra-tech APDS with hyper-dense core
            ammoarray = mAddKey(ammoarray, "APFSDSHD")  'hypervelocity, fin stabilized APDSHD dart
            ammoarray = mAddKey(ammoarray, "Canister")  'a spray of pellets, grapeshot or scrap
            ammoarray = mAddKey(ammoarray, "HEAT")  'a shaped-charged, armor piercing "anti-tank" round
            ammoarray = mAddKey(ammoarray, "HEDP")  ' HE Dual Purpose, a fragmentation and HEAT round
            ammoarray = mAddKey(ammoarray, "HESH")  ' HE Squash-Head, plastic explosive built to defeat armor
            ammoarray = mAddKey(ammoarray, "LE")  'low explosive, a fused round filled with black powder
            ammoarray = mAddKey(ammoarray, "HE")  'high explosive, a modern explosive, fragmenting round
            ammoarray = mAddKey(ammoarray, "HEC")  'HE Concussion, with bit blast and little fragmentation
            ammoarray = mAddKey(ammoarray, "HEPF")  'HE proximity fused round for anti aircraft fire
            ammoarray = mAddKey(ammoarray, "SAPLE")  'LE fused to go off after piercing armor
            ammoarray = mAddKey(ammoarray, "SAPHE")  'HE fused to go off after piercing armor
            ammoarray = mAddKey(ammoarray, "APEX")  'armor piercing explosive cannon round
            ammoarray = mAddKey(ammoarray, "Beehive")  'high-tech shrapnel using anti-personnel darts
            ammoarray = mAddKey(ammoarray, "CHEM")  'chemical, a gas or smoke-filled round
            'cannons of min boresize 60
            
            If mvarBoreSize >= 60 Then
                ammoarray = mAddKey(ammoarray, "HEDC")  'HE Depth Charge round for attacking submarines
                ammoarray = mAddKey(ammoarray, "Shrapnel")  ' a bursting shell filled with bullets
            'cannons of min boresize 80 at TL 8
            End If
            'cannons of min boresize 120 at TL7 or less
            If (mvarBoreSize >= 120) And (mvarTL <= 7) Then
                ammoarray = mAddKey(ammoarray, "FASCAM")  'drops a field of scattered anti-armor mines
                ammoarray = mAddKey(ammoarray, "ICM")  'cluster munitions that scatter grenades over an area
                ammoarray = mAddKey(ammoarray, "SICM")  'smart ICM that home in on targets
                ammoarray = mAddKey(ammoarray, "SATNUC")  'smart ICM with tiny, shaped nuclear warheads
            ElseIf (mvarBoreSize >= 80) And (mvarTL >= 8) Then
                ammoarray = mAddKey(ammoarray, "FASCAM")  'drops a field of scattered anti-armor mines
                ammoarray = mAddKey(ammoarray, "ICM")  'cluster munitions that scatter grenades over an area
                ammoarray = mAddKey(ammoarray, "SICM")  'smart ICM that home in on targets
                ammoarray = mAddKey(ammoarray, "SATNUC")  'smart ICM with tiny, shaped nuclear warheads
            End If
            
        End If
    End If
    
    'now lets add the Nukes
    '1 kiloton
    If mvarTL <= 7 Then
        If mvarBoreSize >= 250 Then ammoarray = mAddKey(ammoarray, "1 megaton Nuke")
        If mvarBoreSize >= 200 Then ammoarray = mAddKey(ammoarray, "100 kiloton Nuke")
        If mvarBoreSize >= 150 Then ammoarray = mAddKey(ammoarray, "10 kiloton Nuke")
        If mvarBoreSize >= 125 Then ammoarray = mAddKey(ammoarray, "1 kiloton Nuke")
    ElseIf mvarTL = 8 Then
        If mvarBoreSize >= 225 Then ammoarray = mAddKey(ammoarray, "1 megaton Nuke")
        If mvarBoreSize >= 175 Then ammoarray = mAddKey(ammoarray, "100 kiloton Nuke")
        If mvarBoreSize >= 125 Then ammoarray = mAddKey(ammoarray, "10 kiloton Nuke")
        If mvarBoreSize >= 100 Then ammoarray = mAddKey(ammoarray, "1 kiloton Nuke")
    ElseIf mvarTL = 9 Then
        If mvarBoreSize >= 200 Then ammoarray = mAddKey(ammoarray, "1 megaton Nuke")
        If mvarBoreSize >= 150 Then ammoarray = mAddKey(ammoarray, "100 kiloton Nuke")
        If mvarBoreSize >= 100 Then ammoarray = mAddKey(ammoarray, "10 kiloton Nuke")
        If mvarBoreSize >= 75 Then ammoarray = mAddKey(ammoarray, "1 kiloton Nuke")
    ElseIf mvarTL = 10 Then
        If mvarBoreSize >= 175 Then ammoarray = mAddKey(ammoarray, "1 megaton Nuke")
        If mvarBoreSize >= 125 Then ammoarray = mAddKey(ammoarray, "100 kiloton Nuke")
        If mvarBoreSize >= 75 Then ammoarray = mAddKey(ammoarray, "10 kiloton Nuke")
        If mvarBoreSize >= 50 Then ammoarray = mAddKey(ammoarray, "1 kiloton Nuke")
    ElseIf mvarTL >= 11 Then
        If mvarBoreSize >= 150 Then ammoarray = mAddKey(ammoarray, "1 megaton Nuke")
        If mvarBoreSize >= 100 Then ammoarray = mAddKey(ammoarray, "100 kiloton Nuke")
        If mvarBoreSize >= 50 Then ammoarray = mAddKey(ammoarray, "10 kiloton Nuke")
        If mvarBoreSize >= 25 Then ammoarray = mAddKey(ammoarray, "1 kiloton Nuke")
    End If
    'now lets add the Micronukes
    If (mvarTL >= 9) And (mvarBoreSize >= 40) Then
        ammoarray = mAddKey(ammoarray, ".1 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".01 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".001 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".0001 kiloton Nuke")
    ElseIf (mvarTL = 10) And (mvarBoreSize >= 30) Then
        ammoarray = mAddKey(ammoarray, ".1 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".01 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".001 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".0001 kiloton Nuke")
    ElseIf (mvarTL = 11) And (mvarBoreSize >= 20) Then
        ammoarray = mAddKey(ammoarray, ".1 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".01 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".001 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".0001 kiloton Nuke")
    ElseIf (mvarTL = 12) And (mvarBoreSize >= 15) Then
        ammoarray = mAddKey(ammoarray, ".1 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".01 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".001 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".0001 kiloton Nuke")
    ElseIf (mvarTL = 13) And (mvarBoreSize >= 12) Then
        ammoarray = mAddKey(ammoarray, ".1 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".01 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".001 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".0001 kiloton Nuke")
    ElseIf (mvarTL = 14) And (mvarBoreSize >= 10) Then
        ammoarray = mAddKey(ammoarray, ".1 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".01 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".001 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".0001 kiloton Nuke")
    ElseIf (mvarTL = 15) And (mvarBoreSize >= 7.5) Then
        ammoarray = mAddKey(ammoarray, ".1 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".01 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".001 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".0001 kiloton Nuke")
    ElseIf (mvarTL = 16) And (mvarBoreSize >= 5) Then
        ammoarray = mAddKey(ammoarray, ".1 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".01 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".001 kiloton Nuke")
        ammoarray = mAddKey(ammoarray, ".0001 kiloton Nuke")
    End If

FillAmmunitionList = ammoarray
End Function

Sub GetTypeDamages()
Dim I As Integer
Dim Suffix1 As String
Dim Suffix2 As String
Dim x As Single
Dim b As Single
Dim h As Single
Dim s As Single
Dim C As Single
On Error Resume Next

mvarBurstRadius = -1
b = mvarBoreSize
C = b ^ 3
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

If mvarBoreSize <= 44 Then
    s = 0.2
ElseIf mvarBoreSize <= 49 Then
    s = 0.3
ElseIf mvarBoreSize <= 54 Then
    s = 0.5
ElseIf mvarBoreSize <= 59 Then
    s = 0.7
Else
    s = 1
End If


'reset the secondary damage to 0
mvarDamage2 = Empty

'ammunition modifier updates to 1/2D and MaxRange
For I = 1 To UBound(AmmoMatrix())
    If AmmoMatrix(I).Name = mvarAmmunitionType Then
            Exit For
    End If
Next

'get suffix for fragmentation damage
If AmmoMatrix(I).Fragmentation Then
    If mvarBoreSize < 20 Then
        Suffix2 = ""
    ElseIf mvarBoreSize <= 34 Then
        Suffix2 = "[2d]"
    ElseIf mvarBoreSize <= 59 Then
        Suffix2 = "[4d]"
    ElseIf mvarBoreSize <= 94 Then
        Suffix2 = "[6d]"
    ElseIf mvarBoreSize <= 160 Then
        Suffix2 = "[10d]"
    ElseIf mvarBoreSize > 160 Then
        Suffix2 = "[12d]"
    End If
End If
'get suffix for armor divisor
If AmmoMatrix(I).Divisor <> 0 Then Suffix1 = "(" & AmmoMatrix(I).Divisor & ")"

TypeDamage1 = AmmoMatrix(I).Damage1
TypeDamage2 = AmmoMatrix(I).Damage2

Select Case AmmoMatrix(I).Name
Case "Solid", "Chainshot", "AP", "APC", "API", "APCR", "HP", "Plastic", "Baton", _
     "Needle", "APDS", "APS", "APDU", "APDSDU", "APFSDS", "APFSDSDU", "Superwire", _
     "APSHD", "APDSHD", "APFSDSHD", "Shotshell"
    
    mvarDamage1 = ConvertDamage(mvarKEDamage * AmmoMatrix(I).Multiplier) & Suffix1
Case "Shrapnel"
        If mvarKEDamage * AmmoMatrix(I).Multiplier > 1 Then
            mvarDamage1 = "1d"
        Else
            mvarDamage1 = ConvertDamage(mvarKEDamage * AmmoMatrix(I).Multiplier) & Suffix1
        End If
         
Case "Canister"
    mvarDamage1 = ConvertDamage(Sqr(mvarKEDamage))
    
Case "Beehive"
    mvarDamage1 = ConvertDamage(2.7 * mvarKEDamage / b)
    
Case "LE", "HE", "HEC", "HEDC", "HEPF", "HESH"
    mvarDamage1 = ConvertDamage(C * x / AmmoMatrix(I).Multiplier) & Suffix2

Case "SAPLE", "SAPHE", "APEX"
    mvarDamage1 = ConvertDamage(mvarKEDamage) & Suffix1
    mvarDamage2 = ConvertDamage(C * x / AmmoMatrix(I).Multiplier) & Suffix2

Case "HEAT"
    mvarDamage1 = ConvertDamage(b * s * h) & Suffix1
Case "HEDP"
    mvarDamage1 = ConvertDamage(b * s * h) & Suffix1 & Suffix2
    
Case "CHEM", "FASCAM", "ICM", "SICM", "SATNUC"
    mvarBurstRadius = AmmoMatrix(I).Multiplier * b * b
    ' these also have special damages which dont plug into the formula
    Select Case AmmoMatrix(I).Name
        Case "CHEM"
            mvarTypeDamage1 = "Special"
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
    mvarDamage1 = AmmoMatrix(I).Formula
End Select


End Sub
Private Function GetDamage() As Double
Dim fKEDamage As Double 'holds numeric value of Kinetic Energy damage before its converted to a GURPS format string
Dim l As Single 'barrel length modifier
Dim G As Integer 'gun technology modifier
Dim T As Single 'tech level modifier
Dim P As Single 'power modifier

'get L
If mvarBarrel = "extremely short" Then
    l = 0.3
ElseIf mvarBarrel = "very short" Then
    l = 0.45
ElseIf mvarBarrel = "short" Then
    l = 0.6
ElseIf mvarBarrel = "medium" Then
    l = 0.75
ElseIf mvarBarrel = "long" Then
    l = 0.9
ElseIf mvarBarrel = "very long" Then
    l = 1.05
ElseIf mvarBarrel = "extremely long" Then
    l = 1.2
End If

'get G
If (mvarTechnology = "conventional smoothbore") Or (mvarTechnology = "conventional rifled") Then
    G = 1
ElseIf mvarTechnology = "electromag" Then
    G = 4
ElseIf mvarTechnology = "gravitic" Then
    G = 6
End If

'get T
If mvarTL >= 6 Then
    T = 1
ElseIf mvarTL = 5 Then
    If mvarDatatype = MuzzleLoader Then
        T = 0.666
    Else
        T = 0.75
    End If
ElseIf mvarTL <= 4 Then
    T = 0.666
End If

'get p
If mvarPowerOption = "extra-low-powered" Then
    P = 0.25
ElseIf mvarPowerOption = "low-powered" Then
    P = 0.5
Else
    P = 1
End If

fKEDamage = mvarBoreSize * l * G * T * P

GetDamage = fKEDamage

End Function

Private Function GetHalfDamage() As Double 'in yards
Dim TempHalfDamage As Double
Dim s As Integer 'bore SIZE modifier
Dim G As Integer 'gun technology modifier
Dim l As Single 'barrel length modifier
Dim P As Single 'powered
Dim T As Single 'TL modifier
Dim I As Long

On Error Resume Next
'note: rounds doing exp or spcl damage are effected in terms of
'accuracy, but they are not effect in terms of the amount of damage
'they inflict.  I therefore dont need this since HalfDamage
'is a range which is still relevant for showing the reduced accuracy
'(page 105 top)

'ammunition modifier updates to 1/2D and MaxRange
For I = 1 To UBound(AmmoMatrix())
    If AmmoMatrix(I).Name = mvarAmmunitionType Then
            Exit For
    End If
Next


'get S
If mvarBoreSize >= 6 Then
    s = 375
Else: s = 250
End If
'get G
If (mvarTechnology = "conventional smoothbore") Or (mvarTechnology = "conventional rifled") Then
    G = 1
ElseIf mvarTechnology = "electromag" Then
    G = 4
ElseIf mvarTechnology = "gravitic" Then
    G = 6
End If

'get L
If mvarBarrel = "extremely short" Then
    l = 0.15
ElseIf mvarBarrel = "very short" Then
    l = 0.27
ElseIf mvarBarrel = "short" Then
    l = 0.33
ElseIf mvarBarrel = "medium" Then
    l = 0.5
ElseIf mvarBarrel = "long" Then
    l = 0.75
ElseIf mvarBarrel = "very long" Then
    l = 0.9
ElseIf mvarBarrel = "extremely long" Then
    l = 1.2
End If

'get p
If mvarPowerOption = "extra-low-powered" Then
    P = 0.25
ElseIf mvarPowerOption = "low-powered" Then
    P = 0.5
Else
    P = 1
End If

'get T
If mvarTL >= 7 Then
    T = 1
ElseIf mvarTL = 6 Then
    T = 0.9
ElseIf mvarTL = 5 Then
    If mvarTechnology = "conventional rifled" Then
        T = 0.75
    Else
        T = 0.5
    End If
ElseIf mvarTL < 5 Then
    T = 0.5
End If

TempHalfDamage = Sqr(mvarBoreSize) * s * G * l * P * T

' Do the final rounding
If TempHalfDamage <= 100 Then
    ' round to nearest yard
    TempHalfDamage = Round(TempHalfDamage, 0)
ElseIf TempHalfDamage <= 1000 Then ' round to nearest 10 yards
    TempHalfDamage = Round(TempHalfDamage / 10, 0) * 10
ElseIf TempHalfDamage <= 10000 Then ' round to nearest 100 yards
    TempHalfDamage = Round(TempHalfDamage / 100, 0) * 100
ElseIf TempHalfDamage > 10000 Then ' round to nearest 1000 yards
    TempHalfDamage = Round(TempHalfDamage / 1000, 0) * 1000
End If


GetHalfDamage = TempHalfDamage
End Function

Private Function GetMaxRange() As Double 'in yards
Dim R As Integer ' holds range modifier
Dim TempMax As Double

'get R
If (mvarTL <= 6) And (mvarTechnology = "conventional smoothbore") Then
    R = 80
Else
    R = 120
End If

TempMax = (Sqr(mvarhalfDamage) * R) + mvarhalfDamage

' Do the final rounding
If TempMax <= 100 Then
    ' round to nearest yard
    TempMax = Round(TempMax, 0)
ElseIf TempMax <= 1000 Then ' round to nearest 10 yards
    TempMax = Round(TempMax / 10, 0) * 10
ElseIf TempMax <= 10000 Then ' round to nearest 100 yards
    TempMax = Round(TempMax / 100, 0) * 100
ElseIf TempMax > 10000 Then ' round to nearest 1000 yards
    TempMax = Round(TempMax / 1000, 0) * 1000
End If

GetMaxRange = TempMax
End Function

Sub UpdateRanges()
Dim I As Integer

'ammunition modifier updates to 1/2D and MaxRange
For I = 1 To UBound(AmmoMatrix())
    If AmmoMatrix(I).Name = mvarAmmunitionType Then
        mvarMaxRange = mvarMaxRange * AmmoMatrix(I).Range
        mvarhalfDamage = mvarhalfDamage * AmmoMatrix(I).Range
            Exit For
    End If
Next
End Sub

'-------------------------
'Private Function GetMinRange() as single 'in yards
'
'End Function
'--------------------------

Private Function GetAccuracy() As Long
'note: CGLP rules (side bar page 181) state:
' If using cannon launched guided projectiles against a target
'that is laser illuminated (e.g by a laser designator), indirect fire
'is at a +5 bonus to hit and range penalties are ignored
'note: however, this is not something that i can model because its a real
'time game mechanic that must be noted during the course ofplay

Dim I As Integer 'ammomatrix array position
Dim R As Single
Dim Acc As Integer

R = mvarhalfDamage
 
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


'ammunition modifier
For I = 1 To UBound(AmmoMatrix())
    If AmmoMatrix(I).Name = mvarAmmunitionType Then
        Acc = Acc + AmmoMatrix(I).Accuracy
            Exit For
    End If
Next

'smoothbore modifier
If (mvarTechnology = "conventional smoothbore") And (mvarTL <= 5) Then
    Acc = Int(Acc / 2) 'devide in half and round down
ElseIf mvarTechnology = "conventional smoothbore" Then 'at TL6+
    Acc = Acc - 2
End If

'smallarms modifier for autoloaders and light automatics
If mvarBoreSize < 20 Then
    If (mvarDatatype = FastAutoloader) Or (mvarDatatype = SlowAutoloader) Or (mvarDatatype = lightAutomatic) Then
        Acc = Acc - 1
    End If
End If

'get modifier for Cheap, Fine and Very Fine quality
If mvarQuality = "cheap" Then
    Acc = Acc - 1
ElseIf mvarQuality = "fine (accurate)" Then
    Acc = Acc + 1
ElseIf mvarQuality = "very fine (accurate)" Then
    Acc = Acc + 2
End If

GetAccuracy = Acc
End Function

Private Function GetWeight() As Double
Dim l As Single 'barrel length modifier
Dim P As Single 'power modifier
Dim s As Single 'bore Size modifier OR Gun technology modifier for non conventional guns
Dim T As Single 'tech level modifier
Dim R As Single 'loading mechanism modifier
Dim Conventional As Boolean
Dim TempWeight As Double
Dim FBMod As Single 'modifier extra barrels for breech / muzzle loaders
Dim b As Single

b = mvarBoreSize
If (mvarTechnology = "conventional smoothbore") Or (mvarTechnology = "conventional rifled") Then
    Conventional = True
End If

'get L
If Conventional Then
    If mvarBarrel = "extremely short" Then
        l = 0.1
    ElseIf mvarBarrel = "very short" Then
        l = 0.2
    ElseIf mvarBarrel = "short" Then
        l = 0.3
    ElseIf mvarBarrel = "medium" Then
        l = 0.5
    ElseIf mvarBarrel = "long" Then
        l = 0.75
    ElseIf mvarBarrel = "very long" Then
        l = 1
    ElseIf mvarBarrel = "extremely long" Then
        l = 1.25
    End If
Else
    If mvarBarrel = "extremely short" Then
        l = 0.1
    ElseIf mvarBarrel = "very short" Then
        l = 0.25
    ElseIf mvarBarrel = "short" Then
        l = 0.6
    ElseIf mvarBarrel = "medium" Then
        l = 1
    ElseIf mvarBarrel = "long" Then
        l = 1.5
    ElseIf mvarBarrel = "very long" Then
        l = 2
    ElseIf mvarBarrel = "extremely long" Then
        l = 3
    End If
End If

'get P
If Conventional Then
    If (mvarDatatype = lightAutomatic) Or (mvarDatatype = HeavyAutomatic) Then
        If mvarPowerOption = "extra-low-powered" Then
            P = 0.5
        ElseIf mvarPowerOption = "low-powered" Then
            P = 0.8
        Else
            P = 1
        End If
    Else
        If mvarPowerOption = "extra-low-powered" Then
            P = 0.25
        ElseIf mvarPowerOption = "low-powered" Then
            P = 0.5
        Else
            P = 1
        End If
    End If
Else ' electormags and gravitic
    If mvarPowerOption = "extra-low-powered" Then
        P = 0.05
    ElseIf mvarPowerOption = "low-powered" Then
        P = 0.25
    Else
        P = 1
    End If
End If

'get T
If Conventional Then
    If mvarTL <= 3 Then
        T = 2.5
    ElseIf mvarTL = 4 Then
        T = 2
    ElseIf mvarTL = 5 Then
        T = 1.5
    ElseIf mvarTL = 6 Then
        T = 1.2
    ElseIf mvarTL = 7 Then
        T = 1
    Else
        T = 0.6
    End If
Else
    If mvarTL <= 8 Then
        T = 1
    ElseIf mvarTL = 9 Then
        T = 0.6
    Else
        T = 0.4
    End If
End If

' get G or S (actually S is for conventionals but will be used in place of S for calc purposes only)
If Conventional Then
    ' update bore size per errata2
    ' At TL6+ treat bore size as half its actual diameter if low-power or 1/4 diameter if extra-low-power.''
    If mvarTL >= 6 Then
        If mvarPowerOption = "extra-low-powered" Then
            b = b / 4
        ElseIf mvarPowerOption = "low-powered" Then
            b = b / 2
        End If
    End If

    If b <= 25 Then
        s = 0.25
    ElseIf b <= 40 Then
        s = 0.375
    ElseIf b <= 50 Then
        s = 0.5
    ElseIf b <= 60 Then
        s = 0.75
    ElseIf b > 60 Then
        s = 1
    End If
Else
    If mvarTechnology = "electromag" Then
        s = 1
    ElseIf mvarTechnology = "gravitic" Then
        s = 1.2
    End If
End If


'get R
If (mvarDatatype = MuzzleLoader) Then
    If mvarCarriage Then
        R = 0.8
    Else
        R = 0.4
    End If
ElseIf mvarDatatype = BreechLoader Then
    R = 0.6
ElseIf mvarDatatype = Revolver Then
    R = 0.6 + (0.05 * mvarCylinders)
ElseIf (mvarDatatype = ManualRepeater) Or (mvarDatatype = SlowAutoloader) Then
    R = 0.75
ElseIf (mvarDatatype = FastAutoloader) Or (mvarDatatype = lightAutomatic) Then
    R = 1
ElseIf (mvarDatatype = HeavyAutomatic) Or (mvarDatatype = MechanicalGatling) Then
    R = 1.5
ElseIf mvarDatatype = ElectricGatling Then
    R = 1.2 + (0.3 * mvarCylinders)
End If

If (mvarDatatype <> ElectricGatling) And (mvarDatatype <> MechanicalGatling) Then
    If mvarCylinders > 1 Then
        FBMod = 0.7 + (0.3 * mvarCylinders)
    Else
        FBMod = 1
    End If
Else
    FBMod = 1
End If



TempWeight = mvarBoreSize * mvarBoreSize * l * P * s * T * R * FBMod
    
GetWeight = Round(TempWeight, 2)

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

If mvarWeight < 2.5 Then
    TSS = 11
ElseIf mvarWeight < 10 Then
    TSS = 12
ElseIf mvarWeight < 15 Then
    TSS = 14
ElseIf mvarWeight < 26 Then
    TSS = 17
ElseIf mvarWeight < 401 Then
    TSS = 20
ElseIf mvarWeight < 2001 Then
    TSS = 25
Else
    TSS = 30
End If

GetSnapShot = TSS
End Function

Private Function GetRoF() As String
Dim sTempRoF As String
Dim Denominator As Long
Dim Numerator As String
Dim bFraction As Boolean

Numerator = ""
Denominator = 0

Select Case mvarDatatype
Case MuzzleLoader
    If mvarCylinders > 1 Then
        Numerator = Format(mvarCylinders) & ":"
    Else
        Numerator = "1/"
        
    End If

    If mvarBoreSize < 120 Then
        If mvarTL <= 3 Then
            If mvarPowerOption <> "normal" Then
                Denominator = 45
            Else
                Denominator = 90
            End If
        ElseIf mvarTL = 4 Then
            If mvarPowerOption <> "normal" Then
                Denominator = 30
            Else
                Denominator = 60
            End If
        ElseIf mvarTL = 5 Then
            If mvarPowerOption <> "normal" Then
                Denominator = 10
            Else
                Denominator = 20
            End If
        ElseIf mvarTL >= 6 Then
            If mvarPowerOption <> "normal" Then
                Denominator = 4
            Else
                Denominator = 8
            End If
        End If
    Else
        If mvarTL <= 3 Then
            Denominator = mvarBoreSize / 1.33
        ElseIf mvarTL = 4 Then
            Denominator = mvarBoreSize / 2
        ElseIf mvarTL = 5 Then
            Denominator = mvarBoreSize / 6
        Else
            Denominator = mvarBoreSize / 15
        End If
        'adjust for low or extra low powered
        If mvarPowerOption <> "normal" Then
            Denominator = Round(Denominator, 0)
        Else
            Denominator = Round(Denominator / 2, 0)
        End If
    End If

Case BreechLoader
    If mvarCylinders > 1 Then
        Numerator = Format(mvarCylinders) & ":"
    Else
        Numerator = "1/"
    End If
    
    If mvarBoreSize < 60 Then
        If mvarTL <= 3 Then
            If mvarPowerOption <> "normal" Then
                Denominator = 20
            Else
                Denominator = 40
            End If
        ElseIf mvarTL = 4 Then
            If mvarPowerOption <> "normal" Then
                Denominator = 10
            Else
                Denominator = 20
            End If
        ElseIf mvarTL = 5 Then
            If mvarPowerOption <> "normal" Then
                Denominator = 3
            Else
                Denominator = 5
            End If
        ElseIf mvarTL >= 6 Then
            If mvarPowerOption <> "normal" Then
                Denominator = 2
            Else
                Denominator = 3
            End If
        End If
    Else
        If mvarTL <= 3 Then
            Denominator = mvarBoreSize / 1.5
        ElseIf mvarTL = 4 Then
            Denominator = mvarBoreSize / 3
        ElseIf mvarTL = 5 Then
            Denominator = mvarBoreSize / 12
        Else
            Denominator = mvarBoreSize / 20
        End If
        'adjust for low or extra low powered
        If mvarPowerOption <> "normal" Then
            Denominator = Round(Denominator, 0)
        Else
            Denominator = Round(Denominator / 2, 0)
        End If
    End If

Case Revolver

    If mvarBoreSize < 60 Then
        If mvarTL <= 5 Then
            Denominator = 1
        ElseIf mvarTL >= 6 Then
            Denominator = 3
        End If
    Else
        If mvarTL <= 5 Then
            Denominator = 2
            Numerator = "1/"
        Else
            Denominator = 1
        End If
    End If

Case ManualRepeater
    If (mvarBoxMagazine) And (mvarPowerOption = "normal") Then
        Denominator = 1
    ElseIf mvarPowerOption <> "normal" Then
        Denominator = 3
    Else
        Numerator = "1/"
        Denominator = 2
    End If
    
Case SlowAutoloader
    If mvarBoreSize <= 30 Then
        Denominator = 1
    Else
        Denominator = Round(mvarBoreSize / 20, 0)
        Numerator = "1/"
    End If
    
Case FastAutoloader
    If mvarBoreSize <= 20 Then
        Denominator = 3
    ElseIf mvarBoreSize <= 30 Then
        Denominator = 2
    ElseIf mvarBoreSize <= 60 Then
        Denominator = 1
    Else
        Denominator = Round(mvarBoreSize / 40, 0)
        Numerator = "1/"
    End If

Case MechanicalGatling
    If mvarBoreSize <= 60 Then
        Denominator = Round(mvarDXPlusSkill / 2, 0)
    Else
        Denominator = Round(mvarDXPlusSkill / 3, 0)
    End If
End Select

'save the String ROF
If Numerator = "" Then
    mvarsRoF = Format(Denominator)
    mvardRof = Denominator
Else
    mvarsRoF = Numerator & Format(Denominator)
    mvardRof = 1 / Denominator
End If
End Function

Public Function FillRoFList() As String()
    Dim iTempRof As Long
    Dim I As Long
    
    ReDim rofarray(1)
    
    Select Case mvarDatatype
        Case lightAutomatic
            If mvarBoreSize <= 10 Then
                iTempRof = 16
            Else
                If mvarBoreSize > 160 Then
                    iTempRof = 1
                Else
                    iTempRof = Round(160 / mvarBoreSize, 0)
                End If
            End If
            If iTempRof < 3 Then iTempRof = 3
            ReDim rofarray(iTempRof - 2)
            For I = 1 To iTempRof - 2
                rofarray(I) = 2 + I
            Next
            
        Case HeavyAutomatic
            If mvarBoreSize <= 20 Then
                iTempRof = 20
            Else
                If mvarBoreSize > 400 Then
                    iTempRof = 1
                Else
                    iTempRof = Round(400 / mvarBoreSize, 0)
                End If
            End If
            If iTempRof < 3 Then iTempRof = 3
            ReDim rofarray(iTempRof - 2)
            For I = 1 To iTempRof - 2
                rofarray(I) = 2 + I
            Next
        
        Case ElectricGatling
            If mvarBoreSize <= 20 Then
                iTempRof = Round(16.67 * mvarCylinders, 0)
            Else
                If mvarBoreSize > 334 Then
                    iTempRof = mvarCylinders
                Else
                    iTempRof = Round((334 / mvarBoreSize) * mvarCylinders, 0)
                End If
            End If
            ReDim rofarray(iTempRof)
            For I = 1 To iTempRof
                rofarray(I) = I
            Next
        End Select
    
    FillRoFList = rofarray
End Function
Private Function GetCost() As Double
Dim TempCost As Double
Dim D As Integer
Dim M As Single

If mvarWeight < 10 Then
    TempCost = (50 * mvarWeight) + 250
ElseIf mvarWeight <= 100 Then
    TempCost = 75 * mvarWeight
Else
    TempCost = (25 * mvarWeight) + 5000
End If

'get D modifier for conventional guns
If (mvarTechnology = "conventional smoothbore") Or (mvarTechnology = "conventional rifled") Then
    M = 1
    If mvarTL <= 5 Then
        D = 10
    ElseIf mvarTL = 6 Then
        D = 5
    ElseIf mvarTL <= 8 Then
        D = 1 'no modifier at TL 7 and 8
    ElseIf mvarTL = 9 Then
        D = 2 'guns start to get cheaper again
    Else
        D = 4
    End If
    'modifiers for smoothbore
    If mvarTechnology = "conventional smoothbore" Then
        M = 0.8
    End If
    'modifier for electric loading or electric gatling
    If (mvarElectric) Or (mvarDatatype = ElectricGatling) Then
        M = M * 1.5
    End If
ElseIf mvarTechnology = "electromag" Then
    D = 1
    If mvarTL <= 8 Then
        M = 16
    ElseIf mvarTL = 9 Then
        M = 8
    Else
        M = 4
    End If
ElseIf mvarTechnology = "gravitic" Then
    D = 1
    If mvarTL <= 11 Then
        M = 16
    ElseIf mvarTL = 12 Then
        M = 8
    Else
        M = 4
    End If
Else
    D = 1
    M = 1
End If

TempCost = TempCost / D * M


'todo double cost for any gun/launcher combo weapon

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

Private Function GetWeightPerShot() As Double
Dim D As Single 'small versus large arms modifier
Dim P As Single 'power option modifier
Dim l As Single 'barrel length modifier
Dim Divisor As Long 'tl5- smoothbore modifier
Dim sngRet As Double
Dim OtherMod As Single
Dim I As Integer

'set divisor value and D value
If (mvarTL <= 5) And (mvarTechnology = "conventional smoothbore") Then
    Divisor = 125000
    D = 1
Else
    Divisor = 1
    If mvarBoreSize < 20 Then
        D = 0.00005
    Else
        D = 0.0000375
    End If
End If

'get P
If mvarPowerOption = "low-powered" Then
    P = 0.5
ElseIf mvarPowerOption = "extra-low-powered" Then
    P = 0.2
Else
    P = 1
End If

'get L
If (mvarTechnology = "electromag") Or (mvarTechnology = "gravitic") Then
    l = 1
ElseIf mvarBoreSize < 20 Then
    If (mvarBarrel = "very long") Or (mvarBarrel = "extremely long") Then
        l = 3
    ElseIf (mvarBarrel = "medium") Or (mvarBarrel = "long") Then
        l = 2.5
    Else
        l = 1.5
    End If
Else
    If (mvarBarrel = "very long") Or (mvarBarrel = "extremely long") Then
        l = 2
    ElseIf (mvarBarrel = "medium") Or (mvarBarrel = "long") Then
        l = 1.7
    Else
        l = 1.5
    End If
End If

'TODO part of the "other modifiers" should include rocket assisted ammunition
'other modifiers
If mvarRecoiless Then 'or (mvarammuntiontype =rocketassisted???
    OtherMod = 1.5

Else
    OtherMod = 1
End If


mvarTempWPS = mvarBoreSize ^ 3 / Divisor * D * P * l * OtherMod

'ammunition modifier
For I = 1 To UBound(AmmoMatrix())
    If AmmoMatrix(I).Name = mvarAmmunitionType Then
        sngRet = mvarTempWPS / AmmoMatrix(I).WPS
            Exit For
    End If
Next

GetWeightPerShot = Round(sngRet, 5)
End Function

Private Function GetVolumePerShot() As Double

GetVolumePerShot = mvarWPS / 150
End Function

Private Function GetCostPerShot() As Double
Dim C As Integer ' technology modifier
Dim T As Integer 'tech level modifier
Dim I As Integer
Dim TempCPS As Double

'get C
If (mvarTechnology = "electromag") Or (mvarTechnology = "gravitic") Then
    C = 8
Else
    C = 2
End If

'get T
If mvarTL <= 5 Then
    T = 10
ElseIf mvarTL = 6 Then
    T = 4
Else
    T = 1
End If

TempCPS = mvarWPS * C / T

'ammunition modifier
For I = 1 To UBound(AmmoMatrix())
    If AmmoMatrix(I).Name = mvarAmmunitionType Then
        Select Case mvarAmmunitionType
            Case "1 megaton Nuke", "100 kiloton Nuke", "10 kiloton Nuke", "1 kiloton Nuke", _
                ".1 kiloton Nuke", ".01 kiloton Nuke", ".001 kiloton Nuke", ".0001 kiloton Nuke"
                TempCPS = (TempCPS * 3) + AmmoMatrix(I).CPS
            Case Else
                TempCPS = TempCPS * AmmoMatrix(I).CPS
            End Select
            Exit For
    End If
Next

GetCostPerShot = Round(TempCPS, 2)
End Function

Private Function GetPowerReqt() As Double
Dim TempPower As Double
Dim l As Single
Dim P As Single

'get L
If mvarBarrel = "extremely short" Then
    l = 0.1
ElseIf mvarBarrel = "very short" Then
    l = 0.2
ElseIf mvarBarrel = "short" Then
    l = 0.4
ElseIf mvarBarrel = "medium" Then
    l = 0.6
ElseIf mvarBarrel = "long" Then
    l = 0.8
ElseIf mvarBarrel = "very long" Then
    l = 1
ElseIf mvarBarrel = "extremely long" Then
    l = 1.2
End If

'get P
If mvarPowerOption = "extra-low-powered" Then
    P = 0.0625
ElseIf mvarPowerOption = "low-powered" Then
    P = 0.25
Else
    P = 1
End If

'do final calcs
If mvarAdvancedOption = "electrothermal" Then 'rules for advanced option
    TempPower = (5000 * mvarWPS * l * P * mvardRof) / 10
ElseIf mvarTechnology = "electromag" Then
    TempPower = 5000 * mvarWPS * l * P * mvardRof
ElseIf mvarTechnology = "gravitic" Then
    TempPower = 500 * mvarWPS * l * P * mvardRof
ElseIf mvarElectric Then
    TempPower = 0.1 * mvarWPS * mvardRof
Else
    TempPower = 0
End If

'

GetPowerReqt = TempPower

End Function

Private Function GetShots() As String
'number of shots the weapon has ready to fire.
If (mvarDatatype = Revolver) Or (mvarDatatype = BreechLoader) Or (mvarDatatype = MuzzleLoader) Then
    GetShots = Str(mvarCylinders)
Else
    GetShots = "var."
End If
End Function

Private Sub GetAdvancedOptionModifiers()

If mvarAdvancedOption = "none" Then
    Exit Sub
ElseIf mvarAdvancedOption = "plastic-cased ammunition" Then
    mvarWPS = mvarWPS / 1.5
    mvarVPS = mvarVPS / 1.5
    mvarCPS = mvarCPS * 2
ElseIf mvarAdvancedOption = "caseless" Then
    IncreaseMalf (mvarMalfunction) 'reduce the chance of malfunction
    mvarWPS = mvarWPS / 2
    mvarVPS = mvarVPS / 2
    mvarCost = mvarCost * 2
ElseIf mvarAdvancedOption = "liquid propellant" Then
    IncreaseMalf (mvarMalfunction) 'reduce the chance of malfunction
    mvarWPS = mvarWPS / 2
    mvarVPS = mvarVPS / 2
    mvarCost = mvarCost * 3
ElseIf mvarAdvancedOption = "electrothermal" Then
'note: i perform power reqt. modifiers in the actual function call for electrothermal option
    mvarAccuracy = mvarAccuracy + 1
    mvarMaxRange = mvarMaxRange * 1.5
    mvarhalfDamage = mvarhalfDamage * 1.5
    mvarKEDamage = mvarKEDamage * 1.5
    '//todo: delete the following line? (page 111 sidebar bottom)
    GetTypeDamages 'i should verify that with new KEdamage, i need to redo damage.
    'Im thinking i dont have to call GetTypeDamages again. that i should delete the
    'above line and justkeep the below as is
    mvarWPS = mvarWPS / 2.5
    mvarCPS = mvarCPS / 2.5
    mvarVPS = mvarVPS / 2.5
    mvarCost = mvarCost * 4
    
End If


End Sub


Private Function GetLoaders() As Long
Dim C As Integer
Dim T As Integer
Dim TempLoaders As Single

If (mvarDatatype = MuzzleLoader) Or (mvarDatatype = BreechLoader) Or (mvarDatatype = Revolver) Or (mvarDatatype = ManualRepeater) Then
    If mvarBoreSize > 66 Then
        If (mvarDatatype = MuzzleLoader) And (mvarCarriage) Then
            C = 2
        Else
            C = 1
        End If
        If mvarTL <= 3 Then
            T = 4
        ElseIf mvarTL = 4 Then
            T = 3
        ElseIf mvarTL = 5 Then
            T = 2
        Else
            T = 1
        End If
        
        TempLoaders = ((mvarBoreSize / 66) - 1) * C * T
        GetLoaders = RoundUP(TempLoaders)
    Else
        GetLoaders = 0
    End If
End If


End Function
