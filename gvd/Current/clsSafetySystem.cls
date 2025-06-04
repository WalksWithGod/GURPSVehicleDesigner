VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSafetySystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarTL As Integer
Private mvarWeight As Double
Private mvarVolume As Double
Private mvarCost As Double
Private mvarPowerReqt As Double
Private mvarLocation As String
Private mvarParent As String
Private mvarKey As String
Private mvarDR As Long
Private mvarSurfaceArea As Double
Private mvarHitPoints As Double
Private mvarOccupancy As Long
Private mvarGReduction As Single
Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
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

Public Property Let Occupancy(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Occupancy = 5
    mvarOccupancy = vdata
End Property

Public Property Get Occupancy() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MatrixPos
    Occupancy = mvarOccupancy
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


Select Case mvarDatatype
    Case GravityWeb, EjectionSeat, Airbag, CrashWeb
       If (InstallPoint = Cabin) Or (InstallPoint = LuxuryCabin) Or (InstallPoint = Suite) Or (InstallPoint = LuxurySuite) Or (InstallPoint = CrampedCrewStation) Or (InstallPoint = NormalCrewStation) Or (InstallPoint = RoomyCrewStation) Or (InstallPoint = CycleCrewStation) Or (InstallPoint = HarnessCrewStation) Or (InstallPoint = CycleSeat) Or (InstallPoint = CrampedSeat) Or (InstallPoint = NormalSeat) Or (InstallPoint = RoomySeat) Then
          TempCheck = True
        Else
            InfoPrint 1, "Ejection Seat, Airbag, Crashweb, and Gravity Webs must be added to a Crew Station, Seat, Cabin or Suite.  Note: Cabins and Suites are assumed to have the same number of seats as its Occupancy."
            TempCheck = False
        End If
                  
    Case WombTank
       If (InstallPoint = Suite) Or (InstallPoint = LuxurySuite) Or (InstallPoint = Cabin) Or (InstallPoint = LuxuryCabin) Or (InstallPoint = Hammock) Or (InstallPoint = Bunk) Or (InstallPoint = SmallGalley) Or (InstallPoint = CrampedCrewStation) Or (InstallPoint = NormalCrewStation) Or (InstallPoint = RoomyCrewStation) Or (InstallPoint = CrampedSeat) Or (InstallPoint = NormalSeat) Or (InstallPoint = RoomySeat) Or (InstallPoint = CrampedStandingRoom) Or (InstallPoint = NormalStandingRoom) Or (InstallPoint = RoomyStandingRoom) Then
          TempCheck = True
        Else
            InfoPrint 1, "Womb Tanks must be added to a Bunks, Cabins, Suites,Hammocks, Galleys, Crew Station, Seat or Standing Room with the exception of Cycle and Harness versions."
            TempCheck = False
        End If
    Case CrewEscapeCapsule
        If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = equipmentPod) Then
            TempCheck = True
        Else
            InfoPrint 1, "Crew Escape Capsule must be placed in Body, Superstructure, Pod, Equipment Pod, Turret, Popturret, Arm, Wing, Open Mount or Leg."
            TempCheck = False
        End If
        
        
    Case GravCompensator
        If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = equipmentPod) Or (InstallPoint = Module) Then
            TempCheck = True
        Else
            InfoPrint 1, "Grav Compensators must be placed in Body, Superstructure, Pod, Equipment Pod, Turret, Popturret, Arm, Wing, Open Mount, Leg or Module."
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

Public Property Let Quantity(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quantity = 5
    mvarQuantity = vdata
End Property


Public Property Get Quantity() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quantity
    Quantity = mvarQuantity
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

Public Property Let GReduction(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GReduction = 5
    mvarGReduction = vdata
End Property

Public Property Get GReduction() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GReduction
    GReduction = mvarGReduction
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
mvarQuantity = 1
End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case EjectionSeat
        
    Case Airbag
        
    Case CrashWeb
       
    Case WombTank
        
    Case GravityWeb
      
    Case GravCompensator
    Case CrewEscapeCapsule
        
       
End Select
mvarOccupancy = 1
End Sub

Public Sub GetMatrixIndex()
Dim i As Integer
If mvarDatatype = 0 Then Exit Sub
mvarMatrixPos = 0 'init the counter
For i = 1 To UBound(SafetyMatrix())
    If SafetyMatrix(i).ID = mvarDatatype Then
        If SafetyMatrix(i).TL >= mvarTL Then
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
    Dim ParentComponent As Integer
    Dim OC As Long
    Dim lngNumberofSystems As Long
    Dim sPrint1 As String
    Dim sPrint2 As String
    Dim sPrintPlural As String
    Dim sPrintPlural2 As String
    Dim sPrintPlural3 As String
    Dim sPrintPlural4 As String
    
    ParentComponent = Veh.Components(mvarParent).Datatype
    
    mvarLocation = GetLocation
    
    If mvarMatrixPos = 0 Then Exit Sub
    
    If mvarDatatype = WombTank Then
        
        mvarWeight = SafetyMatrix(mvarMatrixPos).Weight * Veh.Components(Parent).Volume
        mvarCost = SafetyMatrix(mvarMatrixPos).Cost * mvarWeight
        mvarVolume = Weight / SafetyMatrix(mvarMatrixPos).Volume
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)
    ElseIf (ParentComponent = Cabin) Or (ParentComponent = LuxuryCabin) Or (ParentComponent = Suite) Or (ParentComponent = LuxurySuite) Then
        mvarQuantity = Veh.Components(mvarParent).Occupancy * Veh.Components(mvarParent).Quantity
        mvarWeight = SafetyMatrix(mvarMatrixPos).Weight
        mvarCost = SafetyMatrix(mvarMatrixPos).Cost
        mvarVolume = SafetyMatrix(mvarMatrixPos).Volume
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)
    ElseIf mvarDatatype = CrewEscapeCapsule Then
        mvarWeight = SafetyMatrix(mvarMatrixPos).Weight * mvarOccupancy
        mvarCost = SafetyMatrix(mvarMatrixPos).Cost * mvarOccupancy
        mvarVolume = SafetyMatrix(mvarMatrixPos).Volume * mvarOccupancy
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)
    ElseIf mvarDatatype = GravCompensator Then
        mvarWeight = SafetyMatrix(mvarMatrixPos).Weight
        mvarCost = SafetyMatrix(mvarMatrixPos).Cost
        mvarVolume = SafetyMatrix(mvarMatrixPos).Volume
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)
    Else
        mvarQuantity = Veh.Components(Parent).Quantity
        mvarWeight = SafetyMatrix(mvarMatrixPos).Weight
        mvarCost = SafetyMatrix(mvarMatrixPos).Cost
        mvarVolume = SafetyMatrix(mvarMatrixPos).Volume
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)
    End If
    
    'do final rounding
    mvarWeight = Round(mvarWeight * mvarQuantity, 2)
    mvarCost = Round(mvarCost * mvarQuantity, 2)
    mvarVolume = Round(mvarVolume * mvarQuantity, 2)
    mvarSurfaceArea = Round(mvarSurfaceArea * mvarQuantity, 2)
    
    mvarPowerReqt = SafetyMatrix(mvarMatrixPos).Power * mvarQuantity
    
    'find out G Reduction if gravcompensator
    If mvarDatatype = GravCompensator Then
        Dim MaxG As Long
        Dim BaseReduction As Double 'MPJ 07/07/00 increased from long to double
        Dim NumCompensators As Long
        Dim element As Object
        Dim TempReduction As Double 'MPJ 07/07/00 increased from single to double
        Dim i As Integer
        Dim Lweight As Double 'MPJ 07/07/00 increasd from single to double
        
        '//for backward compatibility, make sure quantity is at least 1
        If mvarQuantity < 1 Then mvarQuantity = 1
        
        'determine maximum g reduction
        If mvarTL > 12 Then
            i = mvarTL - 12
            MaxG = 2 * 2 ^ i
            BaseReduction = 4000000 * 2 ^ i
        Else
            MaxG = 2
            BaseReduction = 4000000
        End If
        
        'find the number of compensators
        'todo: assuming i used recursion, i would send an iterator object to bring me
        'back a reference via:
        'set o = Body.Itterator.getObject(GravCompesator) <-- with that being a class ID
        
        For Each element In Veh
            If element.Datatype = GravCompensator Then
                NumCompensators = NumCompensators + element.Quantity
            End If
        Next
        
        'find the reduction
        Lweight = Veh.Stats.LoadedWeight
        If Lweight <= 0 Then Lweight = mvarWeight
        TempReduction = (BaseReduction * NumCompensators) / Lweight
        
        If TempReduction > MaxG Then
            mvarGReduction = MaxG
        Else
            mvarGReduction = Round(TempReduction, 2)
        End If
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
     
    'print output
    If mvarGReduction <> 0 Then
        sPrint1 = ", " + Format(GReduction) + " G reduction"
    End If
   
    mvarPrintOutput = NumericToString(mvarQuantity) + " TL" & mvarTL & " " + mvarCustomDescription + sPrintPlural + sPrint1 + _
    " (" + mvarLocation + ", HP " & mvarHitPoints & sPrintPlural3 + ", " + sPrintPlural4 + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ", " + Format(mvarPowerReqt, p_sFormat) + " kW " + sPrint2 + ")." + mvarComment
End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub
