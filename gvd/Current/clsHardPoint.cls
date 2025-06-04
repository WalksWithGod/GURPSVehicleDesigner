VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHardPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarLoadCapacity As Single
Private mvarControl As String
Private mvarCost As Double
Private mvarWeight As Double
Private mvarTapped As Boolean
Private mvarDirection As String

'local variable(s) to hold property value(s)
Private mvarImage As Integer
Private mvarKey As String
Private mvarParent As String
Private mvarSelectedImage As Integer
Private mvarLocation As String
Private mvarDR As Long
Private mvarHitPoints As Double
Private mvarTL As Integer
Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarQuantity As Integer
Private mvarVolume As Double
Private mvarSurfaceArea As Double

Private mvarComment As String
Private mvarCName As String
Private mvarMatrixPos As Long
Private mvarPrintOutput As String
Private mvarZZInit As Byte
Private mvarAbbrev As String
Private mvarIndex As Long
Private mvarLogicalParent As String

Public Property Let LogicalParent(ByVal vdata As String)
    mvarLogicalParent = vdata
End Property

Public Property Get LogicalParent() As String
    LogicalParent = mvarLogicalParent
End Property

Public Property Let index(ByVal vdata As Long)
    mvarIndex = vdata
End Property

Public Property Get index() As Long
    index = mvarIndex
End Property

Public Property Let Abbrev(ByVal vdata As String)
    mvarAbbrev = vdata
End Property

Public Property Get Abbrev() As String
    Abbrev = mvarAbbrev
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





Public Function LocationCheck() As Boolean
Dim TempCheck As Boolean
If mvarDatatype = WeaponBay Then

   If (Veh.Components(Parent).Datatype = Body) Then
            TempCheck = True
        Else
            InfoPrint 1, "Weapon Bays must be installed in the Body"
            TempCheck = False
        End If
Else
    If (Veh.Components(Parent).Datatype = Body) Or (Veh.Components(Parent).Datatype = Wing) Or (Veh.Components(Parent).Datatype = Arm) Then
            TempCheck = True
        Else
            InfoPrint 1, "Hardpoints must be installed in the Body, Wing or Arm"
            TempCheck = False
        End If
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


Public Property Let Tapped(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tapped = 5
    mvarTapped = vdata
End Property


Public Property Get Tapped() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tapped
    Tapped = mvarTapped
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



Public Property Let Control(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Control = 5
    mvarControl = vdata
End Property


Public Property Get Control() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Control
    Control = mvarControl
End Property



Public Property Let LoadCapacity(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LoadCapacity = 5
    mvarLoadCapacity = vdata
End Property


Public Property Get LoadCapacity() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LoadCapacity
    LoadCapacity = mvarLoadCapacity
End Property



Private Sub Class_Initialize()
' set the default properties
mvarCustom = False
TL = gVehicleTL 'Note: the TL here is not really necessary.
mvarQuantity = 1
mvarTapped = True
mvarLoadCapacity = 500
mvarDirection = "front"
End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()

Select Case mvarDatatype
Case HardPoint
   
Case WeaponBay
   
End Select
End Sub

Public Sub GetMatrixIndex()
Dim i As Integer
If mvarDatatype = 0 Then Exit Sub
mvarMatrixPos = 0 'init the counter
For i = 1 To UBound(HardpointMatrix())
    If HardpointMatrix(i).ID = mvarDatatype Then
        If HardpointMatrix(i).TL >= mvarTL Then
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
Dim sPrintPlural2 As String
Dim sPrintPlural3 As String
Dim sPrintPlural4 As String
    
    mvarZZInit = 1
    If mvarMatrixPos = 0 Then Exit Sub
    
    mvarLocation = GetLocation
    
    mvarAbbrev = "Hardpoint"
    If mvarIndex > 0 Then mvarAbbrev = mvarAbbrev & mvarIndex
    
    mvarWeight = Round(mvarQuantity * HardpointMatrix(mvarMatrixPos).Weight * mvarLoadCapacity, 2)
    mvarCost = Round(mvarQuantity * HardpointMatrix(mvarMatrixPos).Cost * mvarLoadCapacity, 2)
    If mvarDatatype = WeaponBay Then
        mvarVolume = Round(mvarQuantity * HardpointMatrix(mvarMatrixPos).Volume * mvarLoadCapacity, 2)
        mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
        mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea)
    End If
    
    If mvarQuantity > 1 Then
        sPrintPlural = "s"
        sPrintPlural2 = " with total "
        sPrintPlural3 = " each "
        sPrintPlural4 = " total of "
    Else
        sPrintPlural = ""
        sPrintPlural2 = " with "
        sPrintPlural3 = ""
        sPrintPlural4 = ""
    End If
    If mvarDatatype = WeaponBay Then
        mvarPrintOutput = NumericToString(mvarQuantity) + sPrintPlural3 + "rated " + Format(mvarLoadCapacity, p_sFormat) + " lbs." _
            + " (" + mvarLocation + ", HP " & mvarHitPoints & sPrintPlural3 + ", " + sPrintPlural4 + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ")." + mvarComment
    Else
        mvarPrintOutput = NumericToString(mvarQuantity) + sPrintPlural3 + "rated " + Format(mvarLoadCapacity, p_sFormat) + " lbs." _
           + " (" + mvarLocation + ", " + Format(mvarWeight, p_sFormat) + " lbs., " + "$" + Format(mvarCost, p_sFormat) + ")." + mvarComment
    End If

End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub


