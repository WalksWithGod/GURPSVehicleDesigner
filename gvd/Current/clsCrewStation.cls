VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrewStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarWeight As Double
Private mvarVolume As Double
Private mvarCost As Double
Private mvarStationFunction As String
Private mvarLocation As String
Private mvarParent As String
Private mvarKey As String
Private mvarDR As Long
Private mvarRuggedized As Boolean
Private mvarSurfaceArea As Double
Private mvarHitPoints As Double
Private mvarTL As Integer
Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarExposed As Boolean
Private mvarExtraSpace As Single
Private mvarGSeat As Boolean
Private mvarBridgeAccessSpace As Boolean
Private mvarQuantity As Long
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




Public Property Let BridgeAccessSpace(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BridgeAccessSpace = 5
    mvarBridgeAccessSpace = vdata
End Property


Public Property Get BridgeAccessSpace() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BridgeAccessSpace
    BridgeAccessSpace = mvarBridgeAccessSpace
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

Select Case mvarDatatype

    Case CrampedCrewStation, NormalCrewStation, RoomyCrewStation
        If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = OpenMount) Or (InstallPoint = Pod) Or (InstallPoint = Arm) Or (InstallPoint = Leg) Or (InstallPoint = Wing) Or (InstallPoint = equipmentPod) Then
            TempCheck = True
        Else
            InfoPrint 1, "Cramped, Normal, and Roomy Crew Stations must be placed in Body, Superstructure, Turret, Popturret, Open Mount, Pod, equipment Pod, Wing, Arm, or Leg."
            TempCheck = False
        End If
    Case CycleCrewStation, HarnessCrewStation
        If InstallPoint = Body Then
            TempCheck = True
        Else
            InfoPrint 1, "Cycle and Harness Crew Stations must be placed in hull."
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
    If mvarQuantity < 1 Then mvarQuantity = 1
    
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

Public Property Let Exposed(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Exposed = 5
    mvarExposed = vdata
End Property


Public Property Get Exposed() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Exposed
    Exposed = mvarExposed
End Property

Public Property Let GSeat(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GSeat = 5
    mvarGSeat = vdata
End Property


Public Property Get GSeat() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GSeat
    GSeat = mvarGSeat
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


Public Property Let StationFunction(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StationFunction = 5
    mvarStationFunction = vdata
End Property


Public Property Get StationFunction() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StationFunction
    StationFunction = mvarStationFunction
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

Public Property Let ExtraSpace(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExtraSpace = 5
    mvarExtraSpace = vdata
End Property


Public Property Get ExtraSpace() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExtraSpace
    ExtraSpace = mvarExtraSpace
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



Private Sub Class_Initialize()
' set the default properties
mvarCustom = False
TL = gVehicleTL
mvarRuggedized = False
mvarGSeat = False
mvarExposed = False
mvarStationFunction = "Pilot controls everything from"
mvarQuantity = 1
End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case CrampedCrewStation
      
    Case NormalCrewStation
       
    Case RoomyCrewStation
      
    Case CycleCrewStation
       
    Case HarnessCrewStation
          
End Select

End Sub

Public Sub GetMatrixIndex()
Dim i As Integer
If mvarDatatype = 0 Then Exit Sub
mvarMatrixPos = 0 'init the counter
For i = 1 To UBound(CrewStationMatrix())
    If CrewStationMatrix(i).ID = mvarDatatype Then
        If CrewStationMatrix(i).TL >= mvarTL Then
            mvarMatrixPos = i
            Exit For
        Else
            mvarMatrixPos = i
        End If
    End If
Next
End Sub

Public Sub StatsUpdate()
mvarZZInit = 1
If mvarMatrixPos = 0 Then Exit Sub

Dim VolumeMod As Single
Dim element As Object
Dim TempVolume As Single
Dim CostMod As Single
Dim sPrint1 As String
Dim sPrint2 As String
Dim sPrintPlural As String
Dim sPrintPlural2 As String
Dim sPrintPlural3 As String
Dim sPrintPlural4 As String

mvarLocation = GetLocation

If mvarExposed Then
    VolumeMod = 0.5
Else
    VolumeMod = 1
End If
If mvarGSeat Then
    CostMod = 500
Else
    CostMod = 0
End If

If mvarBridgeAccessSpace Then
    VolumeMod = VolumeMod * 3
End If

mvarWeight = Round(CrewStationMatrix(mvarMatrixPos).Weight, 2)
mvarCost = Round(CrewStationMatrix(mvarMatrixPos).Cost + CostMod, 2)
mvarVolume = Round(VolumeMod * CrewStationMatrix(mvarMatrixPos).Volume, 2)
mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)

'//now factor in the quantity
mvarWeight = Round(mvarWeight * mvarQuantity, 2)
mvarCost = Round(mvarCost * mvarQuantity, 2)
mvarVolume = Round(mvarVolume * mvarQuantity, 2)
mvarSurfaceArea = Round(mvarSurfaceArea * mvarQuantity, 2)

'produce the print output
 If mvarRuggedized Then
     sPrint1 = "ruggedized "
 Else
     sPrint1 = ""
 End If

 If mvarQuantity > 1 Then
     sPrintPlural = "s"
     sPrintPlural2 = " with total "
     sPrintPlural3 = " each"
     sPrintPlural4 = " total of "
 Else
     sPrintPlural = ""
     sPrintPlural2 = " with "
     sPrintPlural3 = ""
     sPrintPlural4 = ""
 End If
 
  
 If mvarExposed Then
     sPrint2 = ", exposed"
 End If
 
 If mvarGSeat Then
    If mvarQuantity > 1 Then
        sPrint2 = sPrint2 + ", with g-seats"
    Else
        sPrint2 = sPrint2 + ", with g-seat"
    End If
 End If
 
 'todo: make a "bridge" container object and when crew stations are added to it, bridge access space gets
 ' added by the bridge based on number of children it has.
 If mvarBridgeAccessSpace Then
    sPrint2 = sPrint2 & ", with " + Format(mvarVolume / 3, p_sFormat) + " cf bridge access space"
 End If
 
 mvarPrintOutput = mvarStationFunction & NumericToString(mvarQuantity) + " TL" & mvarTL & " " + sPrint1 + mvarCustomDescription + sPrintPlural + sPrint2 _
     + " (" + mvarLocation + ", HP " & mvarHitPoints & sPrintPlural3 + ", " + sPrintPlural4 + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ")." + mvarComment


End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub



