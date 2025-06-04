VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWeaponLauncher"
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
Private mvarRuggedized As Boolean
Private mvarWeight As Double
Private mvarCustomDescription As String
Private mvarQuality As String
Private mvarDiameter As Single
Private mvarMaxLoad As Single
Private mvarSnapShot As Long
Private mvarLoaders As Long
Private mvarRoF As String
Private mvarShots As String
Private mvarCylinders As Long
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


        If (InstallPoint = HardPoint) Or (InstallPoint = WeaponBay) Or (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = equipmentPod) Or (InstallPoint = Module) Then
            TempCheck = True
        Else
            InfoPrint 1, "Launchers must be placed in Body, Superstructure, Pod, equipment Pod,Turret, Popturret, Arm, Wing, Open Mount, Leg, Hardpoint, Weaponbay or Module."
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

Public Property Let Cylinders(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cylinders = 5
    mvarCylinders = vdata
End Property


Public Property Get Cylinders() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cylinders
    Cylinders = mvarCylinders
End Property

Public Property Let Shots(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Shots = 5
    mvarShots = vdata
End Property


Public Property Get Shots() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Shots
    Shots = mvarShots
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



Public Property Let MaxLoad(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxLoad = 5
    mvarMaxLoad = vdata
End Property


Public Property Get MaxLoad() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxLoad
    MaxLoad = mvarMaxLoad
End Property



Public Property Let Diameter(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Diameter = 5
    
    mvarDiameter = vdata
End Property


Public Property Get Diameter() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Diameter
    Diameter = mvarDiameter
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
mvarQuality = "normal"
mvarMount = "normal"
mvarMaxLoad = 1000
mvarDirection = "front"
End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case MuzzleloadingLauncher
        
        mvarDiameter = 30
        mvarCylinders = 1
    Case BreechloadingLauncher
       
         mvarDiameter = 30
        mvarCylinders = 1
    Case ManualRepeaterLauncher
       
         mvarDiameter = 30
         mvarCylinders = 1
    Case RevolverLauncher
        
        mvarCylinders = 5
         mvarDiameter = 30
    Case SlowAutoLoaderLauncher
       
         mvarDiameter = 30
         mvarCylinders = 1
    Case FastAutoLoaderLauncher
        
         mvarDiameter = 30
         mvarCylinders = 1
    Case lightAutomaticLauncher
        
         mvarDiameter = 30
         mvarCylinders = 1
    Case HeavyAutomaticLauncher
       
         mvarDiameter = 30
         mvarCylinders = 1
    Case DisposableLauncher
       
         mvarDiameter = 30
        mvarCylinders = 1
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
    mvarWeight = GetWeight
    mvarVolume = GetVolume
    mvarSnapShot = GetSnapShot
    mvarRoF = GetRoF
    mvarCost = GetCost
    mvarShots = GetShots
    mvarLoaders = GetLoaders
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)

'cost, malf, and accuracy modifiers for Cheap, Fine and Very Fine quality are calced in the functions below
'todo check that "mounting" option is taken into account
'todo check that "recoiless" option is taken into account
'todo check that advancedoption is taken into account

    '//update the cost,weight,volume, surface area and volume based on quantity and ruggedized options
    mvarCost = mvarCost * QRugMod
    mvarWeight = mvarWeight * QRugMod
    mvarVolume = mvarVolume * QRugMod
    
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    mvarLoaders = mvarLoaders * mvarQuantity
    
    'produce the print output
    If mvarRuggedized Then
        sPrint1 = "ruggedized "
    Else
        sPrint1 = ""
    End If
    
    
    'produce the print output
    If mvarMount <> "normal" Then
        sPrint2 = sPrint2 + ", " + mvarMount
    End If
    
    sPrint2 = sPrint2 + ", " + Format(mvarMaxLoad, p_sFormat) + " lbs max load"
    
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

mvarPrintOutput = NumericToString(mvarQuantity) + " TL" & mvarTL & " " + sPrint1 + " " & mvarDiameter & "mm " & mvarCustomDescription + sPrintPlural + sPrint2 _
    + " (" + mvarLocation + sPrintDirection + ", HP " & mvarHitPoints & sPrintPlural2 + ", " + sPrintPlural3 + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ")." + mvarComment

End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(mvarParent).StatsUpdate
End If
End Sub


Private Function GetWeight() As Single
Dim mWPS As Single
Dim R As Single
Dim T As Single
Dim M As Single
Dim TempWeight As Single

'get mWPS
mWPS = mvarMaxLoad

'get R
If mvarDatatype = DisposableLauncher Then
    R = 1
ElseIf mvarDatatype = MuzzleloadingLauncher Then
    R = 1.5
ElseIf mvarDatatype = BreechloadingLauncher Then
    R = 2
ElseIf (mvarDatatype = ManualRepeaterLauncher) Or (mvarDatatype = SlowAutoLoaderLauncher) Then
    R = 2.5
ElseIf mvarDatatype = RevolverLauncher Then
    R = 1.9 + (0.1 * mvarCylinders)
ElseIf mvarDatatype = FastAutoLoaderLauncher Then
    R = 3
ElseIf mvarDatatype = lightAutomaticLauncher Then
    R = 3.5
ElseIf mvarDatatype = HeavyAutomaticLauncher Then
    R = 4.5
End If

' get T
If mvarTL <= 7 Then
    T = 0
Else
    T = 0.5
End If

'get M
If mvarCylinders = 1 Then
    M = 1
ElseIf mvarCylinders > 1 Then
    M = 1 + ((mvarCylinders - 1) * 0.4)
End If

TempWeight = (R - T) * mWPS * M

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
Dim TempRoF As String
Dim iRof As Single
Dim Numerator As String

iRof = 0

Select Case mvarDatatype
Case MuzzleloadingLauncher
    If mvarCylinders > 1 Then
        Numerator = Str(mvarCylinders) & ":"
    Else
        Numerator = "1/"
    End If

    If mvarDiameter < 120 Then
        If mvarTL <= 3 Then
            TempRoF = Numerator & "45"
        ElseIf mvarTL = 4 Then
            TempRoF = Numerator & "30"
        ElseIf mvarTL = 5 Then
            TempRoF = Numerator & "10"
        ElseIf mvarTL >= 6 Then
            TempRoF = Numerator & "6"
        End If
    Else
        If mvarTL <= 3 Then
            iRof = Round(mvarDiameter / 2.66, 0) 'note: i round the iROF here
        ElseIf mvarTL = 4 Then
            iRof = mvarDiameter / 4
        ElseIf mvarTL = 5 Then
            iRof = Round(mvarDiameter / 12, 0)
        Else
            iRof = Round(mvarDiameter / 20, 0)
        End If
        TempRoF = Numerator & Str(iRof)
        iRof = 0
    End If

Case DisposableLauncher
    If mvarCylinders = 1 Then
        TempRoF = "1NR"
    Else
        TempRoF = mvarCylinders & ":1NR"
    End If
    iRof = 0
Case BreechloadingLauncher
    If mvarCylinders > 1 Then
        Numerator = Str(mvarCylinders) & ":"
    Else
        Numerator = "1/"
    End If
    
    If mvarDiameter <= 60 Then
        If mvarTL <= 3 Then
            TempRoF = Numerator & "20"
        ElseIf mvarTL = 4 Then
            TempRoF = Numerator & "10"
        ElseIf mvarTL = 5 Then
            TempRoF = Numerator & "5"
        ElseIf mvarTL >= 6 Then
            TempRoF = Numerator & "2"
        End If
    Else
        If mvarTL <= 3 Then
            iRof = Round(mvarDiameter / 3, 0)
        ElseIf mvarTL = 4 Then
            iRof = Round(mvarDiameter / 6, 0)
        ElseIf mvarTL = 5 Then
            iRof = Round(mvarDiameter / 12, 0)
        Else
            iRof = Round(mvarDiameter / 30, 0)
        End If
        TempRoF = Numerator & Str(iRof)
        iRof = 0
    End If

Case RevolverLauncher

    If mvarDiameter < 40 Then
        If mvarTL <= 5 Then
            TempRoF = "1"
        ElseIf mvarTL >= 6 Then
            TempRoF = "3"
        End If
    Else
        If mvarTL <= 5 Then
            TempRoF = "1/2"
        Else
            TempRoF = 1
        End If
    End If

Case ManualRepeaterLauncher
    If mvarTL <= 5 Then
        TempRoF = "1"
    Else
        TempRoF = "2"
    End If
    
Case SlowAutoLoaderLauncher
    If mvarDiameter <= 40 Then
        TempRoF = "1"
    Else
        iRof = Round(mvarDiameter / 40, 0)
        TempRoF = iRof
        If iRof = 1 Then
            iRof = 0
        End If

    End If
    
Case FastAutoLoaderLauncher
    If mvarDiameter <= 15 Then
        TempRoF = "3"
    ElseIf mvarDiameter <= 20 Then
        TempRoF = "2"
    ElseIf mvarDiameter <= 60 Then
        TempRoF = "1"
    Else
        iRof = Round(mvarDiameter / 60, 0)
        TempRoF = iRof
        If iRof = 1 Then
            iRof = 0
        End If
    End If
    
Case lightAutomaticLauncher
        TempRoF = "Up to " & Str(CLng(160 / mvarDiameter))

Case HeavyAutomaticLauncher
    If mvarDiameter <= 20 Then
        TempRoF = "3 to 20"
    Else
        TempRoF = "Up to " & Str(Round(400 / mvarDiameter, 0))
    End If
End Select


'pass final results
If iRof <> 0 Then
    GetRoF = "1/" & TempRoF
Else
    GetRoF = TempRoF
End If
End Function

Private Function GetCost() As Single
Dim TempCost As Single


'get weight
If mvarWeight < 10 Then
    TempCost = (50 * mvarWeight) + 250
ElseIf mvarWeight <= 100 Then
    TempCost = 75 * mvarWeight
ElseIf mvarWeight > 100 Then
    TempCost = (25 * mvarWeight) + 5000
End If

'get tl  modifier
If mvarTL <= 5 Then
    TempCost = TempCost / 10
ElseIf mvarTL = 6 Then
    TempCost = TempCost / 5
ElseIf mvarTL = 9 Then
    TempCost = TempCost / 2
ElseIf mvarTL >= 10 Then
    TempCost = TempCost / 4
End If

'double cost for automatic or divide by 10 if disposable
If (mvarDatatype = HeavyAutomaticLauncher) Or (mvarDatatype = lightAutomaticLauncher) Then
    TempCost = TempCost * 2
ElseIf mvarDatatype = DisposableLauncher Then
    TempCost = TempCost / 10
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


Private Function GetShots() As String
'number of shots the weapon has ready to fire.
If (mvarDatatype = DisposableLauncher) Or (mvarDatatype = RevolverLauncher) Or (mvarDatatype = BreechloadingLauncher) Or (mvarDatatype = MuzzleloadingLauncher) Then
    GetShots = Str(mvarCylinders)
Else
    GetShots = "var."
End If
End Function

Private Function GetLoaders() As Long
Dim TempLoaders As Single

If (mvarDatatype = MuzzleloadingLauncher) Or (mvarDatatype = BreechloadingLauncher) Or (mvarDatatype = RevolverLauncher) Or (mvarDatatype = ManualRepeaterLauncher) Then
    TempLoaders = (mvarDiameter / 250) - 1
    'TODO rules (page 122 top) say round to nearest whole  number.  Im rounding up, need to decide if that's final
    GetLoaders = RoundUP(TempLoaders)
Else
    GetLoaders = 0
End If

End Function



