VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArmMotor"
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
Private mvarBadGrip As Boolean
Private mvarPoorCoordination As Boolean
Private mvarCheap As Boolean
Private mvarExtendable As Boolean
Private mvarStriker As Boolean
Private mvarST As Long
Private mvarPowerReqt As Double
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
Dim TempCheck As Boolean
Dim InstallPoint As Integer
InstallPoint = Veh.Components(mvarParent).Datatype

        If InstallPoint = Arm Then
            TempCheck = True
        Else
            InfoPrint 1, "Arm Motors must be placed in Arms."
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



Public Property Let Striker(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Striker = 5
    mvarStriker = vdata
End Property


Public Property Get Striker() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Striker
    Striker = mvarStriker
End Property



Public Property Let Extendable(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Extendable = 5
    mvarExtendable = vdata
End Property


Public Property Get Extendable() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Extendable
    Extendable = mvarExtendable
End Property



Public Property Let Cheap(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cheap = 5
    mvarCheap = vdata
End Property


Public Property Get Cheap() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cheap
    Cheap = mvarCheap
End Property



Public Property Let PoorCoordination(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PoorCoordination = 5
    mvarPoorCoordination = vdata
End Property


Public Property Get PoorCoordination() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PoorCoordination
    PoorCoordination = mvarPoorCoordination
End Property



Public Property Let BadGrip(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BadGrip = 5
    mvarBadGrip = vdata
End Property


Public Property Get BadGrip() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BadGrip
    BadGrip = mvarBadGrip
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


Private Sub Class_Initialize()
 

' set the default properties
mvarCustom = False
mvarRuggedized = False

mvarST = 12
End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()


End Sub

Public Sub GetMatrixIndex()
Dim i As Integer
If mvarDatatype = 0 Then Exit Sub

mvarMatrixPos = 0 'init the counter

mvarLocation = GetLocation

For i = 1 To UBound(ArmMotorMatrix())
    If ArmMotorMatrix(i).ID = mvarDatatype Then
        ' arm motors are dependant on the ARM's (the parents) TL and NOT the vehicles TL.
        ' arm motors do not have their own TL either.
        If ArmMotorMatrix(i).TL >= Veh.Components(Parent).TL Then
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
'Note:  Must call Getmatrixindex here because we need to update
'the Counter in case the parent TL has changed
GetMatrixIndex

Dim WeightMod As Single
Dim VolumeMod As Single
Dim CostMod As Single
Dim RugMod As Single
Dim RugHitMod As Integer
Dim sPrint1 As String
Dim sPrint2 As String

'set the ruggedized modifiers
If mvarRuggedized Then
    RugMod = 1.5
    RugHitMod = 2
Else
    RugMod = 1
    RugHitMod = 1
End If

'inti variables
WeightMod = 1
VolumeMod = 1
CostMod = 1

'bad grip option modifier
If mvarBadGrip Then
    CostMod = 0.5
End If

'Cheap option modifier
If mvarCheap Then
    WeightMod = 2
    VolumeMod = 2
    CostMod = CostMod * 0.5
End If

'Extendable option modifier
If mvarExtendable Then
    WeightMod = WeightMod * 2
    VolumeMod = VolumeMod * 2
    CostMod = CostMod * 2
End If

'poor coordination option modifier
If mvarPoorCoordination Then
    CostMod = CostMod * 0.5
End If

'Striker option modifier
If mvarStriker Then
    WeightMod = WeightMod * 0.5
    VolumeMod = VolumeMod * 0.5
    CostMod = CostMod * 0.2
End If

'calc final stats
mvarWeight = Round(mvarST * RugMod * WeightMod * ArmMotorMatrix(mvarMatrixPos).Weight, 2)
mvarCost = Round(mvarST * RugMod * CostMod * ArmMotorMatrix(mvarMatrixPos).Cost, 2)
mvarVolume = Round(mvarST * RugMod * VolumeMod * ArmMotorMatrix(mvarMatrixPos).Volume, 2)
mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)

mvarPowerReqt = Round(mvarST / ArmMotorMatrix(mvarMatrixPos).Power, 2)

'update the stats for its parent
QueryParent

'produce the print output
If mvarRuggedized Then
    sPrint1 = "ruggedized "
Else
    sPrint1 = ""
End If

If mvarBadGrip Then
    sPrint2 = ", " + "bad grip"
End If

If mvarCheap Then
    sPrint2 = sPrint2 + ", " + "cheap"
End If

If mvarExtendable Then
    sPrint2 = sPrint2 + ", extendable"
End If

If mvarPoorCoordination Then
    sPrint2 = sPrint2 + ", poor coordination"
End If

If mvarStriker Then
    sPrint2 = sPrint2 + ", striker"
End If

mvarPrintOutput = "ST " + Format(mvarST) + " " + sPrint1 + mvarCustomDescription + sPrint2 _
    + " (" + mvarLocation + ", HP " & mvarHitPoints & ", " + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ", " + Format(mvarPowerReqt, p_sFormat) + " kW)." + mvarComment

End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub

