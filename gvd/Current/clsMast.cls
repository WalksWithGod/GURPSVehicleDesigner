VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit


Private mvarImage As Integer
Private mvarSelectedImage As Integer
Private mvarParent As String
Private mvarKey As String
Private mvarTL As Integer
Private mvarWeight As Double
Private mvarCost As Double
Private mvarMaterials As String
Private mvarSurfaceArea As Double
Private mvarVolume As Double
Private mvarHeight As Long
Private mvarDependantStructure As Boolean
Private mvarHitPoints As Double
Private mvarLocation As String
Private mvarDR As Long
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

        If (InstallPoint = Body) Or (InstallPoint = Superstructure) Then
            TempCheck = True
        Else
            InfoPrint 1, "Masts must be placed on Body or Superstructure."
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



Public Property Let DependantStructure(ByVal vdata As Boolean)
Attribute DependantStructure.VB_Description = "The only structural subassembly that can be attached atop it is a single open mount and that open mount may not house components whose combined volume in cf is greater than half the mast's height."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DependantStructure = 5
    mvarDependantStructure = vdata
End Property


Public Property Get DependantStructure() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DependantStructure
    DependantStructure = mvarDependantStructure
End Property



Public Property Let Height(ByVal vdata As Long)
Attribute Height.VB_Description = "This will need to be like Armor in that it brings up a mini dialog so that users can set seperate height values for each mass.  The ""number"" property will be read only and will be calculated and displayed to user when user defines each mast height."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Height = 5

    Dim MaxHeight As Long
    mvarHeight = vdata
    If mvarZZInit = 0 Then Exit Property
    
    ' check to see if the user has entered an invalid mast height
    If TL <= 4 Then
        MaxHeight = 150
    Else
        If Materials = "wood" Then
            MaxHeight = 150
        Else
            MaxHeight = 200
        End If
   End If
    ' If the user has entered invalid height, post error message then set Mast to
    ' maximum allowable height
    If mvarHeight > MaxHeight Then
        InfoPrint 1, "Maximum Mast height at this tech level using " & mvarMaterials & " materials is " & MaxHeight & " ft."
        mvarHeight = MaxHeight
    End If
        
End Property


Public Property Get Height() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Height

   Height = mvarHeight
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
    'check that material is wood for tech levels 1 thru 4
    mvarTL = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If (mvarTL < 5) And (mvarMaterials = "metal") Then
        InfoPrint 1, "Metal masts unavailable prior to tech level 5. Material reset to 'wood'."
        mvarMaterials = "wood"
    End If
        
End Property


Public Property Get TL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    
    
    TL = mvarTL
    
End Property



Public Property Let Materials(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Materials = 5
    mvarMaterials = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If (mvarMaterials = "metal") And (mvarTL < 5) Then
        InfoPrint 1, "Metal masts not allowed prior to tech level 5."
        mvarMaterials = "wood"
    End If
    
    
End Property


Public Property Get Materials() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Materials
    Materials = mvarMaterials
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

mvarTL = Veh.Components(1).TL
mvarHeight = 20
mvarQuantity = 1
mvarMaterials = "wood"
End Sub



Public Sub StatsUpdate()
Dim sRotation As String
Dim sPrintPlural1 As String

    mvarZZInit = 1
    
    mvarLocation = GetLocation
    
    mvarAbbrev = "Mast"
    If mvarIndex > 0 Then mvarAbbrev = mvarAbbrev & mvarIndex
    
    ' Calculate the volume for the Mast using the Mast specific formula
    mvarVolume = Round(((mvarHeight ^ 3) / 10000) * mvarQuantity, 2)
    
    ' calculate the surface
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    ' Calculate the cost using special Mast and Open Mount rules
    mvarCost = mvarSurfaceArea * 10
    ' Calculate the weight using special Mast and Open Mount rules
    mvarWeight = mvarSurfaceArea * TechLevelModifier(mvarTL)
    ' Calculate the Hit Points
    mvarHitPoints = mvarQuantity * 2 * mvarSurfaceArea
    
    'generate print output
    If mvarQuantity > 1 Then sPrintPlural1 = "s"
    mvarPrintOutput = NumericToString(mvarQuantity) + " " + Format(mvarHeight) + " ft. " + mvarMaterials + " " + mvarCustomDescription + sPrintPlural1 + " " + " (atop " + Veh.Components(mvarParent).CustomDescription + ")."


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


