VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBattlesuitSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarWeight As Double
Private mvarVolume1 As Single
Private mvarVolume2 As Single
Private mvarVolume3 As Single
Private mvarVolume4 As Single
Private mvarCost As Double
Private mvarLocation As String
Private mvarParent As String
Private mvarKey As String
Private mvarDR As Long
Private mvarSurfaceArea As Double
Private mvarHitPoints As Double
Private mvarTL As Integer
Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarImage As Integer
Private mvarSelectedImage As Integer
Private mvarPilotWeight As Single
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




Public Property Let PilotWeight(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PilotWeight = 5
    mvarPilotWeight = vdata
End Property


Public Property Get PilotWeight() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PilotWeight
    PilotWeight = mvarPilotWeight
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

If InstallPoint = Body Then
    TempCheck = True
Else
    InfoPrint 1, "Battlesuit systems must be placed in hull."
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

Public Property Let Volume1(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Volume1= 5
    mvarVolume1 = vdata
End Property


Public Property Get Volume1() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Volume1
    Volume1 = mvarVolume1
End Property

Public Property Let Volume2(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Volume2 = 5
    mvarVolume2 = vdata
End Property


Public Property Get Volume2() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Volume2
    Volume2 = mvarVolume2
End Property

Public Property Let Volume3(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Volume3 = 5
    mvarVolume3 = vdata
End Property


Public Property Get Volume3() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Volume3
    Volume3 = mvarVolume3
End Property

Public Property Let Volume4(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Volume4 = 5
    mvarVolume4 = vdata
End Property


Public Property Get Volume4() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Volume4
    Volume4 = mvarVolume4
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
mvarPilotWeight = 200

End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case BattlesuitSystem
      
    Case FormFittingBattleSuitSystem
       
   
End Select
StatsUpdate
End Sub

Public Sub StatsUpdate()
Dim TotalVolume As Single
Dim CostMod As Single
Dim WeightMod As Single
mvarZZInit = 1

mvarLocation = GetLocation

If mvarTL <= 9 Then
    CostMod = 1
ElseIf mvarTL = 10 Then
    CostMod = 0.5
ElseIf mvarTL >= 11 Then
    CostMod = 0.25
End If

    
    If mvarDatatype = BattlesuitSystem Then
        mvarVolume1 = mvarPilotWeight / 50
        mvarVolume2 = 0
        mvarVolume3 = 0
        mvarVolume4 = 0
    Else
        mvarVolume1 = mvarPilotWeight / 100 'body
        mvarVolume2 = mvarPilotWeight / 400 'turret
        mvarVolume3 = mvarPilotWeight / 1000 'arms
        mvarVolume4 = mvarPilotWeight / 400 'legs
    End If
    
    mvarWeight = Round(0.2 * mvarPilotWeight, 2)
    mvarCost = Round(CostMod * (3000 + (mvarPilotWeight * 20)), 2)
    
    TotalVolume = Round(mvarVolume1 + mvarVolume2 + mvarVolume3 + mvarVolume4, 2)
    
    mvarSurfaceArea = CalcSurfaceArea(TotalVolume)
    mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)

'update the stats for the Body
QueryParent
End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub





