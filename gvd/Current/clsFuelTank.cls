VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFuelTank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarTL As Integer
Private mvarWeight As Double
Private mvarVolume As Double
Private mvarCost As Double
Private mvarFire As Integer
Private mvarLocation As String
Private mvarParent As String
Private mvarKey As String
Private mvarDR As Long
Private mvarRuggedized As Boolean
Private mvarSurfaceArea As Double
Private mvarHitPoints As Double

Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarCapacity As Single
Private mvarFailSafePoints As Long
Private mvarFuelWeight As Single
Private mvarFuelCost As Single
Private mvarFuelType As Integer
Private mvarFuelFire As Integer
Private mvarFuel As String   'this is the string that appears in property's dialog for user

Private mvarImage As Integer
Private mvarSelectedImage As Integer
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
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrintOutput = 5
    mvarPrintOutput = vdata
End Property


Public Property Get PrintOutput() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrintOutput
    PrintOutput = mvarPrintOutput
End Property


Public Property Let MatrixPos2(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MatrixPos2 = 5
    mvarMatrixPos2 = vdata
End Property


Public Property Get MatrixPos2() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MatrixPos2
    MatrixPos2 = mvarMatrixPos2
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




Public Function LocationCheck() As Boolean
Dim TempCheck As Boolean
Dim InstallPoint As Integer
InstallPoint = Veh.Components(mvarParent).Datatype

        If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = equipmentPod) Then
            TempCheck = True
        Else
            InfoPrint 1, "Fuel Tanks must be placed in Body, Superstructure, Pod, equipment Pod, Turret, Popturret, Arm, Wing, Open Mount or Leg."
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



Public Property Let ParentDatatype(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ParentDatatype = 5
    mvarParentDatatype = vdata
End Property


Public Property Get ParentDatatype() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ParentDatatype
    ParentDatatype = mvarParentDatatype
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

Public Property Let Fuel(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Fuel = 5
    mvarFuel = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If mvarFuel = "ethanol" Then
        mvarFuelType = EthanolAlchohol
    ElseIf mvarFuel = "methanol" Then mvarFuelType = MethanolAlchohol
    ElseIf mvarFuel = "aviation gas" Then mvarFuelType = AviationGas
    ElseIf mvarFuel = "cadmium" Then mvarFuelType = Cadmium
    ElseIf mvarFuel = "diesel" Then mvarFuelType = Diesel
    ElseIf mvarFuel = "gasoline" Then mvarFuelType = Gasoline
    ElseIf mvarFuel = "jet fuel" Then mvarFuelType = JetFuel
    ElseIf mvarFuel = "rocket fuel" Then mvarFuelType = RocketFuel
    ElseIf mvarFuel = "water" Then mvarFuelType = Water
    ElseIf mvarFuel = "hydrogen" Then mvarFuelType = LiquidHydrogen
    ElseIf mvarFuel = "metal/LOX" Then mvarFuelType = MetalLOX
    ElseIf mvarFuel = "oxygen (LOX)" Then mvarFuelType = LiquidOxygen
    ElseIf mvarFuel = "propane/LNG" Then mvarFuelType = Propane
    End If
    
    GetMatrixIndex
End Property


Public Property Get Fuel() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Fuel
    Fuel = mvarFuel
End Property


Public Property Let FuelType(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FuelType = 5
    mvarFuelType = vdata
    GetMatrixIndex
End Property


Public Property Get FuelType() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FuelType
    FuelType = mvarFuelType
End Property

Public Property Let Fire(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Fire = 5
    mvarFire = vdata
End Property


Public Property Get Fire() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Fire
    Fire = mvarFire
End Property

Public Property Let FuelFire(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FuelFire = 5
    mvarFuelFire = vdata
End Property


Public Property Get FuelFire() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FuelFire
    FuelFire = mvarFuelFire
End Property

Public Property Let FuelCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FuelCost = 5
    mvarFuelCost = vdata
End Property


Public Property Get FuelCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FuelCost
    FuelCost = mvarFuelCost
End Property


Public Property Let FuelWeight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FuelWeight = 5
    mvarFuelWeight = vdata
End Property


Public Property Get FuelWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FuelWeight
    FuelWeight = mvarFuelWeight
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

Public Property Let capacity(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Capacity = 5
    mvarCapacity = vdata
End Property


Public Property Get capacity() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Capacity
    capacity = mvarCapacity
End Property

Public Property Let FailSafePoints(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FailSafePoints = 5
    mvarFailSafePoints = vdata
End Property


Public Property Get FailSafePoints() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FailSafePoints
    FailSafePoints = mvarFailSafePoints
End Property

Public Property Let TL(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TL = 5
    mvarTL = vdata
    GetMatrixIndex
End Property


Public Property Get TL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    TL = mvarTL
End Property





Private Sub Class_Initialize()
ReDim mvarFuelUsingSystemKeyChain(1)

' set the default properties
mvarCustom = False
TL = gVehicleTL
mvarRuggedized = False
mvarCapacity = 1000


End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case CoalBunker
       
        mvarCapacity = 100
        mvarFuelType = Coal
    Case WoodBunker
        
        mvarCapacity = 100
        mvarFuelType = Wood
    Case StandardTank
       
        mvarFuel = "gasoline"
        mvarFuelType = Gasoline
        mvarCapacity = 10
    Case lightTank
       
        mvarFuel = "gasoline"
        mvarFuelType = Gasoline
        mvarCapacity = 20
    Case StandardSelfSealingTank
       
        mvarFuel = "gasoline"
        mvarFuelType = Gasoline
        mvarCapacity = 20
    Case UltralightTank
       
        mvarFuel = "gasoline"
        mvarFuelType = Gasoline
        mvarCapacity = 20
    Case lightSelfSealingTank
      
        mvarFuel = "gasoline"
        mvarFuelType = Gasoline
        mvarCapacity = 20
    Case UltralightSelfSealingTank
        
        mvarCapacity = 20
        mvarFuel = "gasoline"
        mvarFuelType = Gasoline
    Case AntiMatterBay
        
        mvarFailSafePoints = 0
        mvarCapacity = 1
        mvarFuelType = AntiMatter
End Select


End Sub

Public Sub GetMatrixIndex()
Dim i As Integer
If mvarDatatype = 0 Then Exit Sub
'first load the matrix for the Storage Tanks
mvarMatrixPos = 0 'init the counter
For i = 1 To UBound(FuelTankMatrix())
    If FuelTankMatrix(i).ID = mvarDatatype Then
        If FuelTankMatrix(i).TL >= mvarTL Then
            mvarMatrixPos = i
            Exit For
        Else
            mvarMatrixPos = i
        End If
    End If
Next


'Now load the matrix for the fuel itself
mvarMatrixPos2 = 0 'init the counter
For i = 1 To UBound(FuelMatrix())
    If FuelMatrix(i).ID = mvarFuelType Then
        If FuelMatrix(i).TL >= mvarTL Then
            mvarMatrixPos2 = i
            Exit For
        Else
            mvarMatrixPos2 = i
        End If
    End If
Next

End Sub


Public Sub StatsUpdate()
    mvarZZInit = 1
If mvarMatrixPos = 0 Then Exit Sub

Dim TrueWeight As Single
Dim TempCost As Single
Dim TempFuel As Single
Dim TempVolume As Single
Dim QRugMod As Single 'combined quantity and ruggedized multipliers
Dim RugHitMod As Integer 'ruggedized hit point multiplier
Dim sPrint1 As String
Dim sPrint2 As String
Dim sPrint3 As String
Dim sPrintPlural As String
Dim sPrintPlural2 As String
Dim sPrintPlural3 As String
Dim sPrintPlural4 As String

mvarLocation = GetLocation

'set the ruggedized and quantity multipliers
If mvarRuggedized Then
    QRugMod = 1.5
    RugHitMod = 2
Else
    QRugMod = 1
    RugHitMod = 1
End If

'determine if the weight is above or below 5kw and then make adjustments
TrueWeight = mvarCapacity * FuelTankMatrix(mvarMatrixPos).Weight

'Find the volume
'If (mvarDatatype = CoalBunker) Or (mvarDatatype = WoodBunker) Then
'    TempVolume = mvarCapacity
'Else 'MPJ 07/07/00  Coal bunkers and Wood bunkers also need to use the
      'data files.  Volume is not a 1:1 relationship
      
    TempVolume = mvarCapacity * FuelTankMatrix(mvarMatrixPos).Volume
'End If

'find the cost
TempCost = mvarCapacity * FuelTankMatrix(mvarMatrixPos).Cost

'calc stats for Failsafes NOTE: ive placed failsafes before ruggedized calcs
If (mvarDatatype = AntiMatterBay) And (mvarFailSafePoints > 0) Then
    TempVolume = TempVolume * mvarFailSafePoints
    TempCost = TempCost * mvarFailSafePoints
    TrueWeight = TrueWeight * mvarFailSafePoints
End If

'get base stats
mvarWeight = TrueWeight
mvarCost = TempCost
mvarVolume = TempVolume
mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)

'get finals
mvarWeight = Round(QRugMod * mvarWeight, 2)
mvarCost = Round(QRugMod * mvarCost, 2)
mvarVolume = Round(QRugMod * mvarVolume, 2)
mvarSurfaceArea = CalcSurfaceArea(mvarVolume)

'Get fire
mvarFire = FuelTankMatrix(mvarMatrixPos).Fire

'Get the cost for the fuel based on the fueltype
'NOTE: These should not be added to the actual cost of the tank. Thats why its last!
mvarFuelWeight = FuelMatrix(mvarMatrixPos2).Weight * mvarCapacity
mvarFuelCost = FuelMatrix(mvarMatrixPos2).Cost * mvarCapacity
mvarFuelFire = FuelMatrix(mvarMatrixPos2).Fire

'produce the print output
 If mvarRuggedized Then
     sPrint1 = "ruggedized "
 Else
     sPrint1 = ""
 End If
 
 Select Case mvarDatatype
    Case CoalBunker
        sPrint1 = Format(mvarCapacity, p_sFormat) + " cf. " + sPrint1
        sPrint2 = " Holds " + Format(mvarCapacity, p_sFormat) + " cf. coal" + " (" + Format(mvarFuelWeight, p_sFormat) + " lbs)"
    Case WoodBunker
        sPrint1 = Format(mvarCapacity, p_sFormat) + " cf. " + sPrint1
        sPrint2 = " Holds " + Format(mvarCapacity, p_sFormat) + " cf. wood" + " (" + Format(mvarFuelWeight, p_sFormat) + " lbs)"
    Case StandardTank, lightTank, StandardSelfSealingTank, UltralightTank, lightSelfSealingTank, UltralightSelfSealingTank
        sPrint1 = Format(mvarCapacity, p_sFormat) + " gal. " + sPrint1
        sPrint2 = " Holds " + Format(mvarCapacity, p_sFormat) + " gal. " + mvarFuel + " (" + Format(mvarFuelWeight, p_sFormat) + " lbs, fire +" + Format(mvarFuelFire) + ")"
        sPrint3 = ", fire " + Format(mvarFire)
    Case AntiMatterBay
        sPrint1 = Format(mvarCapacity, p_sFormat) + " gram " + sPrint1
        sPrint3 = ", fail safe " + Format(mvarFailSafePoints)
End Select

 mvarPrintOutput = "TL" & mvarTL & " " + sPrint1 + mvarCustomDescription _
     + " (" + mvarLocation + ", HP " & mvarHitPoints & ", " + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + sPrint3 + ")." + sPrint2 + mvarComment

End Sub


Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub
