VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarSubType As String
Private mvarParent As String
Private mvarKey As String

Private mvarImage As Integer
Private mvarSelectedImage As Integer
Private mvarTL As Integer
Private mvarImprovedSuspension As Boolean
Private mvarImprovedSuspensionCost As Single
Private mvarWheelblades As String
Private mvarSnowTires As Boolean
Private mvarRacingTires As Boolean
Private mvarPunctureResistant As Boolean
Private mvarRetractLocation As String
Private mvarImprovedBrakes As Boolean
Private mvarAllwheelSteering As Boolean
Private mvarSmartwheels As Boolean
Private mvarFrameStrength As String
Private mvarMaterials As String
Private mvarResponsive As Boolean
Private mvarRobotic As Boolean
Private mvarBiomechanical As Boolean
Private mvarLivingMetal As Boolean
Private mvarCost As Double
Private mvarImprovedBrakesCost As Single
Private mvarAllWheelSteeringCost As Single
Private mvarSmartWheelsCost As Single
Private mvarSnowTiresCost As Single
Private mvarRacingTiresCost As Single
Private mvarPunctureResistantCost As Single
Private mvarWheelBladesCost As Single
Private mvarWheelBladesWeight As Single
Private mvarEmptySpace As Single   'MPJ 6/30/2000  added so users can have Monster Tires which also decrease ground pressure
Private mvarWeight As Double
Private mvarSurfaceArea As Double
Private mvarVolume As Double
Private mvarHitPoints As Double
Private mvarLocation As String
Private mvarDR As Long
Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarQuantity As Integer
'local variable(s) to hold property value(s)
Private mvarComment As String
Private mvarCName As String
'local variable(s) to hold property value(s)
Private mvarPrintOutput As String
Private mvarZZInit As Byte
Private mvarAbbrev As String
Private mvarLogicalParent As String

Public Property Let LogicalParent(ByVal vdata As String)
    mvarLogicalParent = vdata
End Property

Public Property Get LogicalParent() As String
    LogicalParent = mvarLogicalParent
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

        If InstallPoint = Body Then
            TempCheck = True
        Else
            InfoPrint 1, "Wheels must be placed on hull."
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



Public Property Let SubType(ByVal vdata As String)
Attribute SubType.VB_Description = "Name of the wheels"
'used when assigning a value to the property, on the left side of an assignment.
    mvarSubType = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If vdata <> "retractable" Then
        If mvarRetractLocation <> "none" Then
            InfoPrint 1, "Rectract location invalid with this wheel type.  Retract location has been reset to 'none'"
            mvarRetractLocation = "none"
        End If
    End If

End Property


Public Property Get SubType() As String
'used when retrieving value of a property, on the right side of an assignment.

    SubType = mvarSubType
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
End Property


Public Property Get FrameStrength() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FrameStrength
    FrameStrength = mvarFrameStrength
End Property

Public Property Let ImprovedSuspensionCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ImprovedSuspensionCost = 5
    mvarImprovedSuspensionCost = vdata
End Property


Public Property Get ImprovedSuspensionCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ImprovedSuspensionCost
    ImprovedSuspensionCost = mvarImprovedSuspensionCost
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

Public Property Let ImprovedBrakesCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ImprovedBrakesCost = 5
    mvarImprovedBrakesCost = vdata
End Property


Public Property Get ImprovedBrakesCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ImprovedBrakesCost
    ImprovedBrakesCost = mvarImprovedBrakesCost
End Property

Public Property Let AllWheelSteeringCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AllWheelSteeringCost = 5
    mvarAllWheelSteeringCost = vdata
End Property


Public Property Get AllWheelSteeringCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AllWheelSteeringCost
    AllWheelSteeringCost = mvarAllWheelSteeringCost
End Property

Public Property Let SmartWheelsCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SmartWheelscost = 5
    mvarSmartWheelsCost = vdata
End Property


Public Property Get SmartWheelsCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SmartWheelscost
    SmartWheelsCost = mvarSmartWheelsCost
End Property


Public Property Let PunctureResistantCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PunctureResistantCost = 5
    mvarPunctureResistantCost = vdata
End Property


Public Property Get PunctureResistantCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PunctureResistantCost
    PunctureResistantCost = mvarPunctureResistantCost
End Property

Public Property Let RacingTiresCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RacingTiresCost = 5
    mvarRacingTiresCost = vdata
End Property


Public Property Get RacingTiresCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RacingTiresCost
    RacingTiresCost = mvarRacingTiresCost
End Property

Public Property Let SnowTiresCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SnowTiresCost = 5
    mvarSnowTiresCost = vdata
End Property


Public Property Get SnowTiresCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SnowTiresCost
    SnowTiresCost = mvarSnowTiresCost
End Property

Public Property Let WheelBladesCost(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WheelBladesCost = 5
    mvarWheelBladesCost = vdata
End Property


Public Property Get WheelBladesCost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WheelBladesCost
    WheelBladesCost = mvarWheelBladesCost
End Property

Public Property Let WheelBladesWeight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WheelBladesWeight = 5
    mvarWheelBladesWeight = vdata
End Property


Public Property Get WheelBladesWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WheelBladesWeight
    WheelBladesWeight = mvarWheelBladesWeight
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



Public Property Let Smartwheels(ByVal vdata As Boolean)
   

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Smartwheels = 5
    mvarSmartwheels = vdata
    Exit Property

End Property


Public Property Get Smartwheels() As Boolean
    


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Smartwheels
    Smartwheels = mvarSmartwheels
    Exit Property


End Property



Public Property Let AllwheelSteering(ByVal vdata As Boolean)
  


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AllwheelSteering = 5
    mvarAllwheelSteering = vdata



End Property


Public Property Get AllwheelSteering() As Boolean
   

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AllwheelSteering
    AllwheelSteering = mvarAllwheelSteering

End Property

Public Property Let ImprovedBrakes(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ImprovedBrakes = 5
    mvarImprovedBrakes = vdata
End Property


Public Property Get ImprovedBrakes() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ImprovedBrakes
    ImprovedBrakes = mvarImprovedBrakes


End Property



Public Property Let RetractLocation(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RetractLocation = 5
    mvarRetractLocation = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If vdata <> "none" Then
        If mvarSubType <> "retractable" Then
            mvarSubType = "retractable"
        End If
    End If
End Property


Public Property Get RetractLocation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RetractLocation
    RetractLocation = mvarRetractLocation
    
End Property



Public Property Let PunctureResistant(ByVal vdata As Boolean)
  


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PunctureResistant = 5
    mvarPunctureResistant = vdata


End Property


Public Property Get PunctureResistant() As Boolean
  

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PunctureResistant
    PunctureResistant = mvarPunctureResistant
  End Property



Public Property Let RacingTires(ByVal vdata As Boolean)
   


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RacingTires = 5
    mvarRacingTires = vdata
End Property


Public Property Get RacingTires() As Boolean
  


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RacingTires
    RacingTires = mvarRacingTires
   End Property



Public Property Let SnowTires(ByVal vdata As Boolean)
 

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SnowTires = 5
    mvarSnowTires = vdata
End Property


Public Property Get SnowTires() As Boolean
   
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SnowTires
    SnowTires = mvarSnowTires
End Property

Public Property Let Wheelblades(ByVal vdata As String)
 

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Wheelblades = 5
    mvarWheelblades = vdata
End Property


Public Property Get Wheelblades() As String
  


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Wheelblades
    Wheelblades = mvarWheelblades
End Property



Public Property Let ImprovedSuspension(ByVal vdata As Boolean)
Attribute ImprovedSuspension.VB_Description = "Does the wheel sub assembly have improved suspension or not?"
    


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ImprovedSuspension = 5
    mvarImprovedSuspension = vdata
End Property


Public Property Get ImprovedSuspension() As Boolean
    


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ImprovedSuspension
    ImprovedSuspension = mvarImprovedSuspension
End Property





Private Sub Class_Initialize()
' set the default properties
mvarTL = gVehicleTL
mvarSubType = "standard"
mvarQuantity = "4"
mvarImprovedSuspension = False
mvarWheelblades = "none"
mvarSnowTires = False
mvarRacingTires = False
mvarPunctureResistant = False
mvarRetractLocation = "none"
mvarImprovedBrakes = False
mvarAllwheelSteering = False
mvarSmartwheels = False
mvarFrameStrength = Veh.Components(BODY_KEY).FrameStrength
mvarMaterials = Veh.Components(BODY_KEY).Materials
mvarResponsive = Veh.Components(BODY_KEY).Responsive
mvarRobotic = Veh.Components(BODY_KEY).Robotic
mvarBiomechanical = Veh.Components(BODY_KEY).Biomechanical
mvarLivingMetal = Veh.Components(BODY_KEY).LivingMetal
mvarCost = 0
mvarWeight = 0
mvarSurfaceArea = 0
mvarVolume = 0
mvarHitPoints = 0

   
End Sub


Private Sub Class_Terminate()
    'the class is being destroyed

    
   
End Sub

Public Sub StatsUpdate()

Dim sPrint1 As String
Dim sPrint2 As String
    mvarZZInit = 1
        
    mvarLocation = GetLocation
    mvarAbbrev = "Wheel"
    
    ' Calculate the component volume
    'TODO: NOTE- The rule is simply, we get the volume of the body. But,
    '      when using sequenced 'only necessary' object updates, we need to
    '      get the Body's volume after the body's .Update has been performed.
    '      The body is always the "second to last" then so to speak.  Since wheels/tracks
    '      and a few selected others, will update afterwards.  We should remember that
    '      these subassemblies dont influence the stats of the body (obviously) since if it
    '     did, you'd have an infinite loop of updates.
    Select Case mvarSubType
        Case "small", "retractable"
            mvarVolume = Veh.Components(BODY_KEY).Volume * 0.05
        Case "standard"
            mvarVolume = Veh.Components(BODY_KEY).Volume * 0.1
        Case "heavy", "off-road", "railway"
            mvarVolume = Veh.Components(BODY_KEY).Volume * 0.2
    End Select
    
    
    
    mvarVolume = Round(mvarVolume + mvarEmptySpace, 2)
    ' calculate the surface
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    
    'Improved Suspension cost
    If mvarImprovedSuspension Then mvarImprovedSuspensionCost = mvarSurfaceArea * 100 Else mvarImprovedSuspensionCost = 0
    'improvedbrakes cost
    If mvarImprovedBrakes Then
        mvarImprovedBrakesCost = 20 * mvarSurfaceArea
        If mvarTL = 8 Then
            mvarImprovedBrakesCost = mvarImprovedBrakesCost / 2
        ElseIf mvarTL >= 9 Then
            mvarImprovedBrakesCost = mvarImprovedBrakesCost / 4
        End If
    Else
        mvarImprovedBrakesCost = 0
    End If
    'allwheel steering cost
    If mvarAllwheelSteering Then
        If mvarTL <= 7 Then
            mvarAllWheelSteeringCost = 100 * mvarSurfaceArea
            mvarAllWheelSteeringCost = Maximum(mvarAllWheelSteeringCost, 5000)
        ElseIf mvarTL = 8 Then
            mvarAllWheelSteeringCost = 50 * mvarSurfaceArea
            mvarAllWheelSteeringCost = Maximum(mvarAllWheelSteeringCost, 2500)
        ElseIf mvarTL >= 9 Then
            mvarAllWheelSteeringCost = 25 * mvarSurfaceArea
            mvarAllWheelSteeringCost = Maximum(mvarAllWheelSteeringCost, 1250)
        End If
    Else
        mvarAllWheelSteeringCost = 0
    End If
    
    'smartwheels cost
    If mvarSmartwheels Then
        mvarSmartWheelsCost = Maximum(80 * mvarSurfaceArea, 4000)
        If mvarTL = 9 Then
            mvarSmartWheelsCost = mvarSmartWheelsCost / 2
        ElseIf mvarTL >= 10 Then
            mvarSmartWheelsCost = mvarSmartWheelsCost / 4
        End If
    Else
        mvarSmartWheelsCost = 0
    End If
    
    'snowtires costs
    If mvarSnowTires Then
        If mvarSurfaceArea >= 200 Then
            mvarSnowTiresCost = 200 * mvarQuantity
        Else
            mvarSnowTiresCost = 100 * mvarQuantity
        End If
    Else
        mvarSnowTiresCost = 0
    End If
    
    'racing tires cost
    If mvarRacingTires Then
        If mvarSurfaceArea >= 200 Then
            mvarRacingTiresCost = 500 * mvarQuantity
        Else
            mvarRacingTiresCost = 250 * mvarQuantity
        End If
    Else
        mvarRacingTiresCost = 0
    End If
    
    'puncture resistance tirres cost
    If mvarPunctureResistant Then
        If mvarSurfaceArea >= 200 Then
            mvarPunctureResistantCost = 500 * mvarQuantity
        Else
            mvarPunctureResistantCost = 250 * mvarQuantity
        End If
    Else
        mvarPunctureResistantCost = 0
    End If
    
    'wheelbladescost and weight
    If mvarWheelblades <> "none" Then
        If mvarWheelblades = "rectractable" Then
            mvarWheelBladesWeight = 0.2 * mvarSurfaceArea
            mvarWheelBladesCost = 100 * mvarWheelBladesWeight
        Else
            mvarWheelBladesWeight = 0.1 * mvarSurfaceArea
            mvarWheelBladesCost = 100 * mvarWheelBladesWeight
        End If
    Else
        mvarWheelBladesCost = 0
        mvarWheelBladesWeight = 0
    End If
    ' Calculate the cost
    mvarCost = BasicDesignCost(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials, mvarResponsive, mvarRobotic, mvarBiomechanical, mvarLivingMetal)
    ' Calculate the weight
    mvarWeight = BasicDesignWeight(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials)
    ' Calculate the Hit Points
    mvarHitPoints = CalcHitPoints(TypeName(Me), mvarFrameStrength, mvarSurfaceArea, mvarQuantity)

    'generate the print output
    If mvarRetractLocation <> "none" Then
        sPrint1 = ", rectract into " + mvarRetractLocation
    End If
    If mvarImprovedBrakes Then
        sPrint2 = sPrint2 + ", improved brakes"
    End If
    If mvarAllwheelSteering Then
        sPrint2 = sPrint2 + ", all-wheel steering"
    End If
    If mvarSmartwheels Then
        sPrint2 = sPrint2 + ", smart wheels"
    End If
    mvarPrintOutput = mvarSubType + " " + mvarCustomDescription + " (" + Format(mvarQuantity) + " wheels " + sPrint2 + sPrint1 + ")."

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



