VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWing"
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
Private mvarSubType As String
Private mvarWeight As Double
Private mvarCost As Double
Private mvarFrameStrength As String
Private mvarMaterials As String
Private mvarResponsive As Boolean
Private mvarRobotic As Boolean
Private mvarBiomechanical As Boolean
Private mvarLivingMetal As Boolean
Private mvarSurfaceArea As Double
Private mvarVolume As Double
Private mvarTL As Integer
Private mvarControlledInstability As Boolean
Private mvarControlledInstabilityCost As Single
Private mvarFoldingCost As Single
Private mvarFoldingWeight As Single
Private mvarVariableSweepCost As Single
Private mvarVariableSweepWeight As Single
Private mvarFolding As Boolean
Private mvarVariableSweep As String
Private mvarEmptySpace As Single
Private mvarHitPoints As Double
Private mvarLocation As String
Private mvarDR As Long
Private mvarAccessSpace As Single
Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarSiblingKey As String
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
Dim siblingorientation As String

  mvarOrientation = vdata
  If mvarZZInit = 0 Then Exit Property
  
On Error GoTo errorhandler
siblingorientation = Veh.Components(mvarSiblingKey).Orientation

  'update the sibling as well
    If mvarSiblingKey <> "" Then
        If siblingorientation <> vdata Then
        ElseIf siblingorientation = "right" Then
            Veh.Components(mvarSiblingKey).Orientation = "left"
        Else
            Veh.Components(mvarSiblingKey).Orientation = "right"
        End If
    End If
Exit Property
errorhandler:
'when loading a save vehicle, it will try to update the sibling
'which has not yet been created
    Exit Property
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

        If InstallPoint = Body Then
            TempCheck = True
        Else
            InfoPrint 1, "Wings must be placed on hull."
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



Public Property Let AccessSpace(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccessSpace = 5
    mvarAccessSpace = vdata
End Property


Public Property Get AccessSpace() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccessSpace
    AccessSpace = mvarAccessSpace
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



Public Property Let VariableSweep(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VariableSweep = 5
    mvarVariableSweep = vdata
    If mvarZZInit = 0 Then Exit Property
    
    On Error GoTo errorhandler
     'update the sibling as well
    If mvarSiblingKey <> "" Then
        If Veh.Components(mvarSiblingKey).VariableSweep <> vdata Then
            Veh.Components(mvarSiblingKey).VariableSweep = vdata
            Veh.Components(mvarSiblingKey).StatsUpdate
        End If
    End If
    Exit Property
errorhandler:
'when loading a save vehicle, it will try to update the sibling
'which has not yet been created
    Exit Property
End Property


Public Property Get VariableSweep() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VariableSweep
    VariableSweep = mvarVariableSweep
End Property



Public Property Let Folding(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Folding = 5
    mvarFolding = vdata
    If mvarZZInit = 0 Then Exit Property
    
    'update the sibling as well
    If mvarSiblingKey <> "" Then
        If Veh.Components(mvarSiblingKey).Folding <> vdata Then
            Veh.Components(mvarSiblingKey).Folding = vdata
            Veh.Components(mvarSiblingKey).StatsUpdate
        End If
    End If
End Property


Public Property Get Folding() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Folding
    Folding = mvarFolding
End Property



Public Property Let ControlledInstability(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ControlledInstability = 5
    mvarControlledInstability = vdata
    If mvarZZInit = 0 Then Exit Property
    
    'update the sibling as well
    If mvarSiblingKey <> "" Then
        If Veh.Components(mvarSiblingKey).ControlledInstability <> vdata Then
            Veh.Components(mvarSiblingKey).ControlledInstability = vdata
            Veh.Components(mvarSiblingKey).StatsUpdate
        End If
    End If
End Property


Public Property Get ControlledInstability() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ControlledInstability
    ControlledInstability = mvarControlledInstability
End Property



Public Property Let TL(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TL = 5
    mvarTL = vdata
    If mvarZZInit = 0 Then Exit Property
    
On Error GoTo errorhandler
    'update the sibling as well
    If mvarSiblingKey <> "" Then
        If Veh.Components(mvarSiblingKey).TL <> vdata Then
            Veh.Components(mvarSiblingKey).TL = vdata
            Veh.Components(mvarSiblingKey).StatsUpdate
        End If
    End If
Exit Property
errorhandler:
'when loading a save vehicle, it will try to update the sibling
'which has not yet been created
    Exit Property
End Property


Public Property Get TL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    TL = mvarTL
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



Public Property Let LivingMetal(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LivingMetal = 5
    mvarLivingMetal = vdata
    If mvarZZInit = 0 Then Exit Property
    
     'update the sibling as well
    If mvarSiblingKey <> "" Then
        If Veh.Components(mvarSiblingKey).LivingMetal <> vdata Then
            Veh.Components(mvarSiblingKey).LivingMetal = vdata
            Veh.Components(mvarSiblingKey).StatsUpdate
        End If
    End If
    
End Property


Public Property Get LivingMetal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LivingMetal
    LivingMetal = mvarLivingMetal
End Property



Public Property Let Biomechanical(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Biomechanical = 5
    mvarBiomechanical = vdata
    If mvarZZInit = 0 Then Exit Property
    
    'update the sibling as well
    If mvarSiblingKey <> "" Then
        If Veh.Components(mvarSiblingKey).Biomechanical <> vdata Then
            Veh.Components(mvarSiblingKey).Biomechanical = vdata
            Veh.Components(mvarSiblingKey).StatsUpdate
        End If
    End If
End Property


Public Property Get Biomechanical() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Biomechanical
    Biomechanical = mvarBiomechanical
End Property



Public Property Let Robotic(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Robotic = 5
    mvarRobotic = vdata
    If mvarZZInit = 0 Then Exit Property
    
      'update the sibling as well
    If mvarSiblingKey <> "" Then
        If Veh.Components(mvarSiblingKey).Robotic <> vdata Then
            Veh.Components(mvarSiblingKey).Robotic = vdata
            Veh.Components(mvarSiblingKey).StatsUpdate
        End If
    End If
End Property


Public Property Get Robotic() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Robotic
    Robotic = mvarRobotic
End Property



Public Property Let Responsive(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Responsive = 5
    mvarResponsive = vdata
    If mvarZZInit = 0 Then Exit Property
    
      'update the sibling as well
    If mvarSiblingKey <> "" Then
        If Veh.Components(mvarSiblingKey).Responsive <> vdata Then
            Veh.Components(mvarSiblingKey).Responsive = vdata
            Veh.Components(mvarSiblingKey).StatsUpdate
        End If
    End If
End Property


Public Property Get Responsive() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Responsive
    Responsive = mvarResponsive
End Property



Public Property Let Materials(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Materials = 5
    mvarMaterials = vdata
    If mvarZZInit = 0 Then Exit Property
    
     'update the sibling as well
    If mvarSiblingKey <> "" Then
        If Veh.Components(mvarSiblingKey).Materials <> vdata Then
            Veh.Components(mvarSiblingKey).Materials = vdata
            Veh.Components(mvarSiblingKey).StatsUpdate
        End If
    End If
End Property


Public Property Get Materials() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Materials
    Materials = mvarMaterials
End Property



Public Property Let FrameStrength(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FrameStrength = 5
    mvarFrameStrength = vdata
    If mvarZZInit = 0 Then Exit Property
    
     'update the sibling as well
    If mvarSiblingKey <> "" Then
        If Veh.Components(mvarSiblingKey).FrameStrength <> vdata Then
            Veh.Components(mvarSiblingKey).FrameStrength = vdata
            Veh.Components(mvarSiblingKey).StatsUpdate
        End If
    End If
End Property


Public Property Get FrameStrength() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FrameStrength
    FrameStrength = mvarFrameStrength
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


Public Property Let FoldingCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FoldingCost = 5
    mvarFoldingCost = vdata
End Property


Public Property Get FoldingCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FoldingCost
    FoldingCost = mvarFoldingCost
End Property

Public Property Let VariableSweepCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VariableSweepCost = 5
    mvarVariableSweepCost = vdata
End Property


Public Property Get VariableSweepCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VariableSweepCost
    VariableSweepCost = mvarVariableSweepCost
End Property

Public Property Let ControlledInstabilityCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ControlledInstabilityCost = 5
    mvarControlledInstabilityCost = vdata
End Property


Public Property Get ControlledInstabilityCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ControlledInstabilityCost
    ControlledInstabilityCost = mvarControlledInstabilityCost
End Property

Public Property Let FoldingWeight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FoldingWeight = 5
    mvarFoldingWeight = vdata
End Property


Public Property Get FoldingWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FoldingWeight
    FoldingWeight = mvarFoldingWeight
End Property

Public Property Let VariableSweepWeight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VariableSweepWeight = 5
    mvarVariableSweepWeight = vdata
End Property


Public Property Get VariableSweepWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VariableSweepWeight
    VariableSweepWeight = mvarVariableSweepWeight
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



Public Property Let SubType(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.

    mvarSubType = vdata
    If mvarZZInit = 0 Then Exit Property
    
On Error GoTo errorhandler

     'update the sibling as well
    If mvarSiblingKey <> "" Then
        If Veh.Components(mvarSiblingKey).SubType <> vdata Then
            Veh.Components(mvarSiblingKey).SubType = vdata
            Veh.Components(mvarSiblingKey).StatsUpdate
        End If
    End If
Exit Property
errorhandler:
'when loading a save vehicle, it will try to update the sibling
'which has not yet been created
    Exit Property
End Property


Public Property Get SubType() As String
'used when retrieving value of a property, on the right side of an assignment.

    SubType = mvarSubType
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


Public Property Let SiblingKey(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SiblingKey = 5
    mvarSiblingKey = vdata
End Property


Public Property Get SiblingKey() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SiblingKey
    SiblingKey = mvarSiblingKey
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
mvarTL = gVehicleTL
mvarSubType = "standard"
mvarOrientation = "left"
mvarFolding = False
mvarVariableSweep = "none"
mvarControlledInstability = False
mvarFrameStrength = Veh.Components(BODY_KEY).FrameStrength
mvarMaterials = Veh.Components(BODY_KEY).Materials
mvarResponsive = Veh.Components(BODY_KEY).Responsive
mvarRobotic = Veh.Components(BODY_KEY).Robotic
mvarBiomechanical = Veh.Components(BODY_KEY).Biomechanical
mvarLivingMetal = Veh.Components(BODY_KEY).LivingMetal
mvarCost = 0
mvarWeight = 0
mvarVolume = 0
mvarSurfaceArea = 0
mvarHitPoints = 0

End Sub


Private Sub Class_Terminate()
    'the class is being destroyed

    
   
End Sub

Public Sub StatsUpdate()
Dim sPrint2 As String
Dim element As Object

    mvarZZInit = 1
    mvarLocation = GetLocation
    
    If mvarOrientation = "left" Then
        mvarAbbrev = "WiL"
    Else
        mvarAbbrev = "WiR"
    End If
        
    If mvarIndex > 0 Then mvarAbbrev = mvarAbbrev & mvarIndex
    
    'get the accessspace
    mvarAccessSpace = CalcAccessSpace(mvarKey)
    ' Calculate the component volume
    mvarVolume = CalcCombinedVolume(mvarKey) + mvarEmptySpace + mvarAccessSpace
    ' Calculate volume due to retracts that are of the Body & Wing variety
    If GetRetractLocation = "body & wings" Then
        mvarVolume = mvarVolume + (0.025 * Veh.Components(BODY_KEY).Volume)
    End If
    mvarVolume = Round(mvarVolume, 2)
    ' calculate the surface
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    
    ' apply the surface area multiplier
    Select Case mvarSubType
        Case "standard", "flarecraft"
            mvarSurfaceArea = mvarSurfaceArea * 1.5
            
        Case "high agility", "STOL"
            mvarSurfaceArea = mvarSurfaceArea * 2
        Case "biplane"
            mvarSurfaceArea = mvarSurfaceArea * 3
        Case "triplane"
            mvarSurfaceArea = mvarSurfaceArea * 4
        
        Case "stub"
            mvarSurfaceArea = mvarSurfaceArea * 1 ' no change according to rules page 18)
    
    End Select
    
    ' Calculate the cost
    mvarCost = BasicDesignCost(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials, mvarResponsive, mvarRobotic, mvarBiomechanical, mvarLivingMetal)
    ' Calculate the weight
    mvarWeight = BasicDesignWeight(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials)
    ' Calculate the Hit Points
    mvarHitPoints = CalcHitPoints(TypeName(Me), mvarFrameStrength, mvarSurfaceArea)
    'print output
    If mvarFolding Then
        sPrint2 = ", folding"
    End If
    If mvarVariableSweep <> "none" Then
        sPrint2 = sPrint2 + ", variable sweep"
    End If
    If mvarControlledInstability Then
        sPrint2 = sPrint2 + ", controlled instability"
    End If
    
    mvarPrintOutput = mvarSubType + " " + mvarCustomDescription + " (" + mvarOrientation + " side of " + Veh.Components(Parent).CustomDescription + sPrint2 + ")."

End Sub

Public Sub CalcWingRotorOptionWeightsAndCosts()
Dim structarea As Single
Dim structweight As Single
Dim structcost As Single
Dim divisor1 As Single
Dim divisor2 As Single


structarea = Veh.Stats.StructuralSurfaceArea
structweight = Veh.Stats.StructuralWeight
structcost = Veh.Stats.StructuralCost

divisor1 = structarea * structweight
divisor2 = structarea * structcost

'folding cost and weight
If (mvarFolding) And (divisor1 <> 0) Then
    mvarFoldingWeight = 0.05 * mvarSurfaceArea / divisor1
    mvarFoldingCost = 25 * mvarFoldingWeight
Else
    mvarFoldingWeight = 0
    mvarFoldingCost = 0
End If

'variable sweep cost and weight
If (mvarVariableSweep <> "none") And (divisor1 <> 0) And (divisor2 <> 0) Then
    mvarVariableSweepWeight = 0.1 * mvarSurfaceArea / divisor1
    mvarVariableSweepCost = 0.25 * mvarSurfaceArea / divisor2
    If mvarVariableSweep = "automatic" Then mvarVariableSweepCost = mvarVariableSweepCost * 2
Else
    mvarVariableSweepCost = 0
    mvarVariableSweepWeight = 0
End If

'controlled instability cost
If mvarControlledInstability Then
    mvarControlledInstabilityCost = 100 * mvarSurfaceArea
Else
    mvarControlledInstabilityCost = 0
End If
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


