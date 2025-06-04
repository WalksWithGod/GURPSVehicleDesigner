VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWeaponLiquidProjector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarHitPoints As Double
Private mvarDR As Long
Private mvarQuality As String
Private mvarCustom As Boolean
Private mvarCost As Double
Private mvarDatatype As Integer
Private mvarDescription As String
Private mvarImage As Integer
Private mvarSelectedImage As Integer
Private mvarRuggedized As Boolean
Private mvarKey As String
Private mvarParent As String
Private mvarQuantity As Integer
Private mvarSurfaceArea As Double
Private mvarTL As Integer
Private mvarVolume As Double
Private mvarWeight As Double
Private mvarWPS As Single
Private mvarCPS As Single
Private mvarCustomDescription As String
Private mvarDamage As String
Private mvarTypeDamage As String
Private mvarhalfDamage As Double
Private mvarMaxRange As Double
Private mvarAccuracy As Long
Private mvarSnapShot As Long
Private mvarShots As Long
Private mvarRoF As String
Private mvarPowerReqt As Double
Private mvarMount As String
Private mvarStyle As String
Private mvarAmmunitionType As String
Private mvarLoaders As Long
Private mvarMalfunction As String
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
            InfoPrint 1, "Liquid Projectors must be placed in Body, Superstructure, Pod, equipment Pod,Turret, Popturret, Arm, Wing, Open Mount, Leg or Module."
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

Public Property Let Style(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Style = 5
    mvarStyle = vdata
End Property

Public Property Get Style() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Style
    Style = mvarStyle
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
mvarTL = gVehicleTL
mvarQuantity = 1
mvarQuality = "normal"
mvarMount = "normal"
mvarStyle = "medium"
mvarTypeDamage = "Spcl."
mvarShots = 10
mvarDirection = "front"
End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case FlameThrower
       
    Case WaterCannon
       
        mvarAmmunitionType = "water"
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
    mvarDamage = ConvertDamage(GetDamage)
    mvarMalfunction = GetMalfunction
    mvarLoaders = GetLoaders
    mvarSnapShot = GetSnapShot
    mvarAccuracy = GetAccuracy
    mvarhalfDamage = GetHalfDamage
    mvarMaxRange = GetMaxRange
    mvarWeight = GetWeight
    mvarVolume = GetVolume
    mvarRoF = GetRoF
    mvarCost = GetCost
    mvarWPS = GetWPS
    mvarCPS = GetCPS
    mvarSurfaceArea = CalcSurfaceArea(Volume)
    mvarHitPoints = CalcComponentHitpoints(SurfaceArea) * RugHitMod
    'cost, malf, and accuracy modifiers for Cheap, Fine and Very Fine quality are calced in the functions below
        
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

    sPrint1 = sPrint1 & mvarStyle + " "
    
    If mvarMount <> "normal" Then
        sPrint2 = sPrint2 + ", " + mvarMount
    End If
    
    If mvarDatatype = WaterCannon Then
        sPrint2 = sPrint2 + ", fires " + mvarAmmunitionType
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
    
    mvarPrintOutput = NumericToString(mvarQuantity) + " TL" & mvarTL & " " + sPrint1 + mvarCustomDescription + sPrintPlural + sPrint2 _
        + " (" + mvarLocation + sPrintDirection + ", HP " & mvarHitPoints & sPrintPlural2 + ", " + sPrintPlural3 + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ", " + Format(mvarPowerReqt, p_sFormat) + " kW)." + mvarComment
    

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


If mvarTL <= 3 Then
    TempMalf = 14
ElseIf mvarTL = 4 Then
    TempMalf = 15
ElseIf mvarTL = 5 Then
    TempMalf = 16
ElseIf mvarTL >= 6 Then
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


Private Function GetDamage() As Single
Dim TDam As Single

If mvarDatatype = FlameThrower Then
    If mvarTL <= 4 Then
        TDam = 1
    ElseIf mvarTL = 5 Then
        TDam = 2
    ElseIf mvarTL <= 8 Then
        TDam = 3
    ElseIf mvarTL >= 9 Then
        TDam = 5
    End If
ElseIf mvarDatatype = WaterCannon Then
    If mvarAmmunitionType = "acid" Then
        TDam = 0.8 'this translates into 1d-1 damage
    ElseIf mvarTL <= 4 Then
        TDam = 2
    ElseIf mvarTL = 5 Then
        TDam = 3
    ElseIf mvarTL <= 8 Then
        TDam = 4
    ElseIf mvarTL >= 9 Then
        TDam = 4
    End If
End If

GetDamage = TDam

End Function

Private Function GetHalfDamage() As Single 'in yards
Dim TempHD As Single

If mvarStyle = "light" Then
    TempHD = mvarTL * 5
ElseIf mvarStyle = "medium" Then
    TempHD = mvarTL * 7
ElseIf mvarStyle = "heavy" Then
    TempHD = mvarTL * 10
End If

If mvarTL <= 4 Then TempHD = CLng(TempHD / 10)

GetHalfDamage = TempHD
End Function

Private Function GetMaxRange() As Double 'in yards
Dim TempMax As Single

If mvarStyle = "light" Then
    TempMax = mvarTL * 7
ElseIf mvarStyle = "medium" Then
    TempMax = mvarTL * 10
ElseIf mvarStyle = "heavy" Then
    TempMax = mvarTL * 15
End If

If mvarTL <= 4 Then TempMax = TempMax / 10

GetMaxRange = TempMax
End Function

Private Function GetAccuracy() As Long

Dim i As Integer 'ammomatrix array position
Dim R As Single
Dim Acc As Integer


Acc = mvarTL
'TODO find out if this max of 8 should be done first or last
If Acc > 8 Then Acc = 8

If mvarStyle = "medium" Then
    Acc = Acc + 1
ElseIf mvarStyle = "heavy" Then
    Acc = Acc + 2
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

Private Function GetWeight() As Single
Dim l As Single
Dim W As Single
Dim T As Single
Dim TempWeight As Single

'get L
If mvarDatatype = WaterCannon Then
    l = 160
ElseIf mvarDatatype = FlameThrower Then
    l = 200
End If

'get W
If mvarStyle = "heavy" Then
    W = 2
ElseIf mvarStyle = "medium" Then
    W = 1
ElseIf mvarStyle = "light" Then
    W = 0.5
End If

'get T
If mvarTL <= 5 Then
    T = 1.25
ElseIf mvarTL = 6 Then
    T = 0.75
ElseIf mvarTL = 7 Then
    T = 0.5
ElseIf mvarTL >= 8 Then
    T = 0.25
End If


TempWeight = l * W * T + (T * mvarShots)
    
GetWeight = Round(TempWeight, 2)

End Function

Private Function GetVolume() As Single
If mvarMount = "normal" Then
    GetVolume = mvarWeight / 50
Else
    GetVolume = mvarWeight / 20 'concealed weapons take up more space
End If

End Function

Private Function GetSnapShot() As Long
Dim TSS As Long

If mvarTL <= 5 Then
    TSS = 10
ElseIf mvarTL >= 6 Then
    TSS = 5
End If

GetSnapShot = TSS
End Function

Private Function GetRoF() As String

If mvarDatatype = WaterCannon Then
    If mvarTL <= 5 Then
        GetRoF = 1
    ElseIf mvarTL = 6 Then
        GetRoF = 3
    ElseIf mvarTL >= 7 Then
        GetRoF = 4
    End If
ElseIf mvarDatatype = FlameThrower Then
    If mvarTL <= 5 Then
        GetRoF = 4
    ElseIf mvarTL = 6 Then
        GetRoF = 8
    ElseIf mvarTL >= 7 Then
        GetRoF = 8
    End If
End If

End Function

Private Function GetCost() As Single
Dim TempCost As Single


If mvarTL <= 6 Then
    TempCost = mvarWeight * 5
ElseIf mvarTL >= 7 Then
    TempCost = mvarWeight * 25
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

GetCost = TempCost
End Function

Private Function GetLoaders() As Long

If mvarTL <= 4 Then
    If mvarStyle = "heavy" Then
        GetLoaders = 2
    ElseIf mvarStyle = "medium" Then
        GetLoaders = 1
    Else
        GetLoaders = 0
    End If
Else
    GetLoaders = 0
End If
    
End Function

Private Function GetWPS() As Single
If mvarDatatype = WaterCannon Then
    GetWPS = 4.25
ElseIf mvarDatatype = FlameThrower Then
    GetWPS = 3
End If
End Function

Private Function GetCPS() As Single
If mvarDatatype = FlameThrower Then
    GetCPS = 0.5
ElseIf mvarDatatype = WaterCannon Then
    If mvarAmmunitionType = "water" Then
        GetCPS = 0
    ElseIf mvarAmmunitionType = "acid" Then
        GetCPS = 0.5
    ElseIf mvarAmmunitionType = "foam" Then
        GetCPS = 0.5
    End If
End If

End Function
