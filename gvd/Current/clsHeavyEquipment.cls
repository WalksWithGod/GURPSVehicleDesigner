VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHeavyEquipment"
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
Private mvarPowerReqt As Double
Private mvarHeight As Single
Private mvarST As Long
Private mvarLength As Single
Private mvarTunnelingAbility As Single
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
Private mvarQuantity As Integer
Private mvarDesiredWeight As Single


Private mvarImage As Integer
Private mvarSelectedImage As Integer
Private mvarComment As String
Private mvarCName As String
Private mvarMatrixPos As Long

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
Dim InstallPoint As Integer
Dim TempCheck As Boolean
InstallPoint = Veh.Components(mvarParent).Datatype

Select Case mvarDatatype

    Case ExtendableLadder, TractorBeam, PressorBeam, CombinationBeam
        If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Pod) Or (InstallPoint = Arm) Or (InstallPoint = Leg) Or (InstallPoint = Wing) Then
            TempCheck = True
        Else
            InfoPrint 1, "This particular Heavy equipment item must be placed in Body, Superstructure, Turret, Popturret, Pod, Wing, Arm, or Leg."
            TempCheck = False
        End If
    Case Crane, CraneWithElectroMagnet, WreckingCrane
        If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Turret) Or (InstallPoint = Pod) Or (InstallPoint = OpenMount) Then
            TempCheck = True
        Else
            InfoPrint 1, "Cranes must be placed in Body, Superstructure, Turret, Pod or Open Mount."
            TempCheck = False
        End If
    Case PowerShovel
        If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Then
            TempCheck = True
        Else
            InfoPrint 1, "Power Shovels must be placed in Body, Superstructure, Turret, or Popturret."
            TempCheck = False
        End If
    Case Bore, SuperBore, ForkLift, LaunchCatapult, SkyHook, Winch
        If InstallPoint = Body Then
            TempCheck = True
        Else
            InfoPrint 1, "This particular Heavy equipment item must be placed in hull."
            TempCheck = False
        End If
        
    Case VehicularBridge
        If InstallPoint = Superstructure Then
            TempCheck = True
        Else
            InfoPrint 1, "Vehicular Bridges must be placed in a Superstructure."
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

Public Property Let TunnelingAbility(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TunnelingAbility = 5
    mvarTunnelingAbility = vdata
End Property


Public Property Get TunnelingAbility() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TunnelingAbility
    TunnelingAbility = mvarTunnelingAbility
End Property



Public Property Let Length(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Length = 5
    mvarLength = vdata
End Property


Public Property Get Length() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Length
    Length = mvarLength
End Property

Public Property Let DesiredWeight(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DesiredWeight = 5
    mvarDesiredWeight = vdata
End Property


Public Property Get DesiredWeight() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DesiredWeight
    DesiredWeight = mvarDesiredWeight
End Property

Public Property Let ST(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ST = 5
    mvarST = vdata
End Property


Public Property Get ST() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ST
    ST = mvarST
End Property



Public Property Let Height(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Height = 5
    mvarHeight = vdata
End Property


Public Property Get Height() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Height
    Height = mvarHeight
End Property



Public Property Let PowerReqt(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Power = 5
    mvarPowerReqt = vdata
End Property


Public Property Get PowerReqt() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Power
    PowerReqt = mvarPowerReqt
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
 

' set the default properties
mvarCustom = False
TL = gVehicleTL
mvarRuggedized = False
mvarQuantity = 1

End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case ExtendableLadder
       
        mvarHeight = 25
    Case Crane
       
        mvarHeight = 50
    Case Winch
       
        mvarST = 100
    Case CraneWithElectroMagnet
        
        mvarHeight = 50
    Case PowerShovel
       
        mvarST = 100
    Case WreckingCrane
        
        mvarHeight = 50
    Case ForkLift
        
        mvarST = 100
    Case VehicularBridge
        
        mvarLength = 10
        mvarDesiredWeight = 10000
    Case LaunchCatapult
       
    Case SkyHook
       
    Case Bore
       
        mvarTunnelingAbility = 1
    Case SuperBore
        
        mvarTunnelingAbility = 1
           
    Case TractorBeam
       
        mvarST = 100
    Case PressorBeam
        
        mvarST = 100
    Case CombinationBeam
        
        mvarST = 100
End Select


End Sub

Public Sub GetMatrixIndex()
Dim i As Integer
If mvarDatatype = 0 Then Exit Sub
mvarMatrixPos = 0 'init the counter
For i = 1 To UBound(HeavyequipmentMatrix())
    If HeavyequipmentMatrix(i).ID = mvarDatatype Then
        If HeavyequipmentMatrix(i).TL >= mvarTL Then
            mvarMatrixPos = i
            Exit For
        Else
            mvarMatrixPos = i
        End If
    End If
Next
End Sub


Public Sub StatsUpdate()
Dim HeightMod As Single
Dim STMod As Single
Dim TempWeight As Single
Dim TempCost As Single
Dim TempVolume As Single
Dim TempPower As Single
Dim QRugMod As Single 'combined quantity and ruggedized multipliers
Dim RugHitMod As Integer 'ruggedized hit point multiplier
Dim sPrint1 As String
Dim sPrint2 As String
Dim sPrintPlural As String
Dim sPrintPlural2 As String
Dim sPrintPlural3 As String

    mvarZZInit = 1
    
    mvarLocation = GetLocation
    
    'set the ruggedized and quantity multipliers
    If mvarRuggedized Then
        QRugMod = 1.5 * mvarQuantity
        RugHitMod = 2
    Else
        QRugMod = 1 * mvarQuantity
        RugHitMod = 1
    End If
    
    
    Select Case mvarDatatype
    Case ExtendableLadder
    
    HeightMod = mvarHeight / 6
    
        mvarWeight = HeightMod * HeavyequipmentMatrix(mvarMatrixPos).Weight
        mvarCost = HeightMod * HeavyequipmentMatrix(mvarMatrixPos).Cost
        mvarVolume = HeightMod * HeavyequipmentMatrix(mvarMatrixPos).Volume
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
        mvarPowerReqt = HeavyequipmentMatrix(mvarMatrixPos).Power
        
    Case Crane, CraneWithElectroMagnet, WreckingCrane
    
        HeightMod = mvarHeight / 6
    
    
       
        mvarWeight = HeightMod * HeavyequipmentMatrix(mvarMatrixPos).Weight
        mvarCost = HeightMod * HeavyequipmentMatrix(mvarMatrixPos).Cost
        mvarVolume = HeightMod * HeavyequipmentMatrix(mvarMatrixPos).Volume
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
        mvarPowerReqt = HeightMod * HeavyequipmentMatrix(mvarMatrixPos).Power
        If mvarDatatype = CraneWithElectroMagnet Then mvarPowerReqt = mvarPowerReqt * 2
        
    Case VehicularBridge
    
        TempWeight = mvarDesiredWeight / 10000
        
        mvarWeight = TempWeight * mvarLength * HeavyequipmentMatrix(mvarMatrixPos).Weight
        mvarCost = TempWeight * mvarLength * HeavyequipmentMatrix(mvarMatrixPos).Cost
        mvarVolume = TempWeight * mvarLength * HeavyequipmentMatrix(mvarMatrixPos).Volume
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
        mvarPowerReqt = mvarLength * TempWeight * HeavyequipmentMatrix(mvarMatrixPos).Power
    
    Case Bore, SuperBore
    
        mvarWeight = mvarTunnelingAbility * HeavyequipmentMatrix(mvarMatrixPos).Weight
        mvarCost = mvarTunnelingAbility * HeavyequipmentMatrix(mvarMatrixPos).Cost
        mvarVolume = mvarTunnelingAbility * HeavyequipmentMatrix(mvarMatrixPos).Volume
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
        mvarPowerReqt = mvarTunnelingAbility * HeavyequipmentMatrix(mvarMatrixPos).Power
    
    Case TractorBeam, PressorBeam, CombinationBeam
    
        
        STMod = mvarST / 100
        TempWeight = 200 * STMod
        TempVolume = 4 * STMod
        TempCost = 200 * STMod
        TempPower = 100 * STMod
        
        mvarWeight = HeavyequipmentMatrix(mvarMatrixPos).Weight + TempWeight
        mvarCost = HeavyequipmentMatrix(mvarMatrixPos).Cost + TempCost
        mvarVolume = HeavyequipmentMatrix(mvarMatrixPos).Volume + TempVolume
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
        mvarPowerReqt = TempPower
        
    Case PowerShovel, Winch, ForkLift
    
        STMod = mvarST / 10
    
        mvarWeight = STMod * HeavyequipmentMatrix(mvarMatrixPos).Weight
        mvarCost = STMod * HeavyequipmentMatrix(mvarMatrixPos).Cost
        mvarVolume = STMod * HeavyequipmentMatrix(mvarMatrixPos).Volume
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
        mvarPowerReqt = STMod * HeavyequipmentMatrix(mvarMatrixPos).Power
    
    Case LaunchCatapult
        
        mvarWeight = HeavyequipmentMatrix(mvarMatrixPos).Weight
        mvarCost = HeavyequipmentMatrix(mvarMatrixPos).Cost
        mvarVolume = HeavyequipmentMatrix(mvarMatrixPos).Volume
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
        mvarPowerReqt = HeavyequipmentMatrix(mvarMatrixPos).Power
    
    Case SkyHook
        
        mvarWeight = HeavyequipmentMatrix(mvarMatrixPos).Weight
        mvarCost = HeavyequipmentMatrix(mvarMatrixPos).Cost
        mvarVolume = HeavyequipmentMatrix(mvarMatrixPos).Volume
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
        mvarPowerReqt = HeavyequipmentMatrix(mvarMatrixPos).Power
    End Select
    
    'get finals
    mvarWeight = Round(QRugMod * mvarWeight, 2)
    mvarCost = Round(QRugMod * mvarCost, 2)
    mvarVolume = Round(QRugMod * mvarVolume, 2)
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    mvarPowerReqt = Round(mvarQuantity * mvarPowerReqt, 2)
    
    'produce the print output
    If mvarRuggedized Then
        sPrint1 = "ruggedized "
    Else
        sPrint1 = ""
    End If
    
    If mvarHeight <> 0 Then
        sPrint1 = sPrint1 + Format(mvarHeight) + "ft "
    End If
    
    If mvarST <> 0 Then
        sPrint1 = sPrint1 + "ST " & Format(mvarST) & " " 'mpj 06/29/2000 added space to end
    End If
    
    If mvarLength Then
        sPrint1 = sPrint1 + Format(mvarLength) + "yd "
    End If
    
    If mvarTunnelingAbility <> 0 Then
        sPrint2 = ", " + Format(mvarTunnelingAbility, p_sFormat) + " cf per hour tunneling ability"
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
    mvarPrintOutput = NumericToString(mvarQuantity) + " TL" & mvarTL & " " + sPrint1 + mvarCustomDescription + sPrintPlural + sPrint2 _
        + " (" + mvarLocation + ", HP " & mvarHitPoints & sPrintPlural2 + ", " + sPrintPlural3 + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ", " + Format(mvarPowerReqt, p_sFormat) + " kW)." + mvarComment
    
End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub

