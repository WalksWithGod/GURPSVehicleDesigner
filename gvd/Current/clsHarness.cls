VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHarness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarQuantity As Integer
Private mvarCombinedST As Single
Private mvarTotalHexes As Single
Private mvarTL As Integer
Private mvarWeight As Double
Private mvarCost As Double
Private mvarEfficiency As String
Private mvarMotivePower As Single
Private mvarMotiveThrust As Single
Private mvarSubType As String
Private mvarLocation As String
Private mvarParent As String
Private mvarKey As String
Private mvarDR As Long


Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarAnimalDescription As String
Private mvarHexes As Long
Private mvarSpeed As Long
Private mvarBeastST As Long
Private mvarMove As Long
 'holds the index value of the array item!
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

        If (InstallPoint = Body) Then
            TempCheck = True
        Else
            InfoPrint 1, "A harness must be placed attached to the hull."
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

Public Property Let AnimalDescription(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AnimalDescription = 5
    mvarAnimalDescription = vdata
End Property


Public Property Get AnimalDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AnimalDescription
    AnimalDescription = mvarAnimalDescription
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



Public Property Let Move(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Move = 5
    mvarMove = vdata
End Property

Public Property Get Move() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Move
    Move = mvarMove
End Property

Public Property Let BeastST(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BeastST = 5
    mvarBeastST = vdata
End Property


Public Property Get BeastST() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BeastST
    BeastST = mvarBeastST
End Property



Public Property Let Speed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Speed = 5
    mvarSpeed = vdata
End Property


Public Property Get Speed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Speed
    Speed = mvarSpeed
End Property

Public Property Let Hexes(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Hexes = 5
    mvarHexes = vdata
End Property


Public Property Get Hexes() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Hexes
    Hexes = mvarHexes
End Property

Public Property Let TotalHexes(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalHexes = 5
    mvarTotalHexes = vdata
End Property


Public Property Get TotalHexes() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalHexes
    TotalHexes = mvarTotalHexes
End Property

Public Property Let CombinedST(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalHexes = 5
    mvarCombinedST = vdata
End Property


Public Property Get CombinedST() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalHexes
    CombinedST = mvarCombinedST
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

Public Property Let SubType(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.

    mvarSubType = vdata
End Property


Public Property Get SubType() As String
'used when retrieving value of a property, on the right side of an assignment.

    SubType = mvarSubType
End Property

Public Property Let MotiveThrust(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MotiveThrust = 5
    mvarMotiveThrust = vdata
End Property


Public Property Get MotiveThrust() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MotiveThrust
    MotiveThrust = mvarMotiveThrust
End Property

Public Property Let MotivePower(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MotivePower = 5
    mvarMotivePower = vdata
End Property


Public Property Get MotivePower() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MotivePower
    MotivePower = mvarMotivePower
End Property

Public Property Let Efficiency(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Efficiency = 5
    mvarEfficiency = vdata
End Property


Public Property Get Efficiency() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Efficiency
    Efficiency = mvarEfficiency
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


Private Sub Class_Initialize()
' set the default properties

mvarSubType = "Land Animal"
'mvarCustom = False
mvarQuantity = 2
TL = gVehicleTL
mvarCombinedST = 0
mvarHexes = 3
mvarAnimalDescription = "large horse"
mvarBeastST = 45
mvarMove = 5
mvarSpeed = 15 'speed = move * 2
End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
'set default descirption values for Datatype
If mvarDatatype = RopeHarness Then
   
ElseIf mvarDatatype = YokeandPoleHarness Then
   
ElseIf mvarDatatype = ShaftandCollarHarness Then
   
ElseIf mvarDatatype = WhiffletreeHarness Then
  
End If

End Sub

Public Sub GetMatrixIndex()
Dim i As Integer
If mvarDatatype = 0 Then Exit Sub
mvarMatrixPos = 0 'init the counter
For i = 1 To UBound(HarnessMatrix())
    If HarnessMatrix(i).ID = mvarDatatype Then
        If HarnessMatrix(i).TL >= mvarTL Then
            mvarMatrixPos = i
            Exit For
        Else
            mvarMatrixPos = i
        End If
    End If
Next
End Sub


Public Sub StatsUpdate()
Dim sPrintPlural As String
Dim sPrint1 As String

    mvarZZInit = 1
    If mvarMatrixPos = 0 Then Exit Sub

    mvarLocation = GetLocation
    
    mvarTotalHexes = mvarHexes * mvarQuantity
    mvarCombinedST = mvarQuantity * mvarBeastST
    'calculate the Beasts speed
    mvarSpeed = mvarMove * 2
    
    
    ' Calculate the Weight
    mvarWeight = Round(HarnessMatrix(mvarMatrixPos).Weight * mvarTotalHexes, 2)
    mvarCost = Round(HarnessMatrix(mvarMatrixPos).Cost * mvarTotalHexes, 2)
    If mvarSubType = "Land Animal" Then
        mvarMotivePower = Round(mvarCombinedST * HarnessMatrix(mvarMatrixPos).Efficiency, 2)
    ElseIf mvarSubType = "Swimming Animal" Then
        mvarMotiveThrust = Round(mvarCombinedST * HarnessMatrix(mvarMatrixPos).Efficiency * 20, 2)
    ElseIf mvarSubType = "Flying Animal" Then
        mvarMotiveThrust = Round(mvarCombinedST * HarnessMatrix(mvarMatrixPos).Efficiency * 2, 2)
    End If
    
    'produce print output
    If mvarSubType = "Land Animal" Then
        sPrint1 = Format(mvarMotivePower, p_sFormat) + " kW motive power"
    ElseIf mvarSubType = "Swimming Animal" Then
        sPrint1 = Format(mvarMotiveThrust, p_sFormat) + " lbs. aquatic motive thrust"
    ElseIf mvarSubType = "Flying Animal" Then
        sPrint1 = Format(mvarMotiveThrust, p_sFormat) + " lbs. motive thrust"
    End If
    
    If mvarQuantity > 1 Then
        sPrintPlural = "s"
    Else
        sPrintPlural = ""
    End If
    
    mvarPrintOutput = "TL" & mvarTL & " " + mvarCustomDescription + " connected to " + Format(mvarQuantity) + " " + mvarAnimalDescription + sPrintPlural + " providing " + sPrint1 _
        + " (" + mvarLocation + ", " + Format(mvarWeight, p_sFormat) + " lbs., " + "$" + Format(mvarCost, p_sFormat) + ")." + mvarComment
        
End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub

