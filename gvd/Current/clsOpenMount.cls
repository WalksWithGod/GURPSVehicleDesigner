VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOpenMount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarOrientation As String
Private mvarImage As Integer
Private mvarSelectedImage As Integer
Private mvarParent As String
Private mvarKey As String
Private mvarRotation As String
Private mvarTL As Integer
Private mvarWeight As Double
Private mvarCost As Double
Private mvarSurfaceArea As Double
Private mvarVolume As Double
Private mvarHitPoints As Double
Private mvarLocation As String
Private mvarDR As Long
Private mvarEmptySpace As Single
Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarQuantity As Integer
Private mvarComment As String
Private mvarCName As String
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




Public Property Let Orientation(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Orientation = 5
  
  mvarOrientation = vdata

End Property


Public Property Get Orientation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Orientation
    Orientation = mvarOrientation
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
Dim InstallPoint As Integer
InstallPoint = Veh.Components(mvarParent).Datatype

        If (InstallPoint = Body) Or (InstallPoint = Mast) Or (InstallPoint = Turret) Or _
            (InstallPoint = Superstructure) Or (InstallPoint = Arm) Or (InstallPoint = Leg) Then
            TempCheck = True
        Else
            InfoPrint 1, "Open Mounts must be placed on Body, Superstructure,Turret, Mast, Leg or Arm."
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



Public Property Let EmptySpace(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmptySpace = 5
    mvarEmptySpace = vdata
End Property


Public Property Get EmptySpace() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmptySpace
    EmptySpace = mvarEmptySpace
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
End Property

Public Property Get TL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    TL = mvarTL
End Property

Public Property Let Rotation(ByVal vdata As String)
Attribute Rotation.VB_Description = "Fixed, Limited (180 pivot), or full (360 pivot)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Rotation = 5
    mvarRotation = vdata
End Property

Public Property Get Rotation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Rotation
    Rotation = mvarRotation
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


Private Sub Class_Initialize()
' set the default properties

mvarRotation = "none"
mvarTL = Veh.Components(1).TL
mvarQuantity = 1
mvarOrientation = "top"

   
End Sub


Public Sub StatsUpdate()
        mvarZZInit = 1
    Dim sRotation As String
    Dim sPrint1 As String
    Dim sPrintPlural1 As String
    
    mvarLocation = GetLocation
    
    mvarAbbrev = "Open Mount"
    If mvarIndex > 0 Then mvarAbbrev = mvarAbbrev & mvarIndex
    
    sRotation = mvarRotation
    mvarVolume = 0 ' reset it back to zero
    ' Calculate the volume
    mvarVolume = CalcCombinedVolume(mvarKey) + (EmptySpace * mvarQuantity)
    ' calculate the final volume based on turret rotation setting
    If sRotation = "full" Then
        mvarVolume = mvarVolume * 1.2
    Else
        If sRotation = "limited" Then
            mvarVolume = mvarVolume * 1.1
        End If
    End If
    mvarVolume = Round(mvarVolume, 2)
    ' calculate the surface
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    ' Calculate the cost using special Mast and Open Mount rules
    mvarCost = mvarSurfaceArea * 10
    ' Calculate the weight using special Mast and Open Mount rules
    mvarWeight = mvarSurfaceArea * TechLevelModifier(mvarTL)
    ' Calculate the Hit Points
    mvarHitPoints = mvarSurfaceArea * 2
    
    'generate print output
    If mvarRotation <> "none" Then
        sPrint1 = "with " + mvarRotation + " rotation"
    End If
    If mvarQuantity > 1 Then sPrintPlural1 = "s"
    mvarPrintOutput = NumericToString(mvarQuantity) + " " + mvarCustomDescription + sPrintPlural1 + " " + sPrint1 + " (on " + mvarOrientation + " of " + Veh.Components(mvarParent).CustomDescription + ")."
End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If

End Sub

Public Sub QueryChild()
' if the object has children, query it and check to see if
' more stats/property updates are needed for other objects in the collection
' (see the StatusUpdate property for help on checking for childeren.  Can i use that one in place of this?)

End Sub

