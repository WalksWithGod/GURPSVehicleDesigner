VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSnorkel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarWeight As Double
Private mvarCost As Double
Private mvarVolume As Double
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
Private mvarImage As Integer
Private mvarSelectedImage As Integer
Private mvarPowerPlants As String

Private mvarPlantKey As String

Private mvarLocation As String
Private mvarComment As String
Private mvarCName As String

Private mvarMatrixPos As Long
Private mvarPrintOutput As String
Private combustionenginearray() As String
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

Public Property Let PlantKey(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PlantKey = 5
    mvarPlantKey = vdata
End Property


Public Property Get PlantKey() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PlantKey
    PlantKey = mvarPlantKey
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




Public Function LocationCheck() As Boolean
Dim TempCheck As Boolean
Dim InstallPoint As Integer
InstallPoint = Veh.Components(mvarParent).Datatype

    If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Pod) Or (InstallPoint = Wing) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Superstructure) Or (InstallPoint = Arm) Or (InstallPoint = Leg) Or (InstallPoint = equipmentPod) Then
         TempCheck = True
     Else
         InfoPrint 1, "Snorkels must go in the Body, Pod, equipment Pod, Wing, Turret, Popturret, Superstructure, Arm or Leg."
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


Public Property Let PowerPlants(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PowerPlants = 5
   Dim i, j As Long
    
    mvarPowerPlants = vdata
    If mvarZZInit = 0 Then Exit Property
    
    If mvarPowerPlants = "none" Then
        mvarPlantKey = "-1"
    Else
        i = InStr(1, mvarPowerPlants, " (")
        j = InStr(1, mvarPowerPlants, ")")
        i = i + 2
        mvarPlantKey = (Mid(mvarPowerPlants, i, j - i))
    End If
    
End Property


Public Property Get PowerPlants() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PowerPlants
    PowerPlants = mvarPowerPlants
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


Private Sub Class_Initialize()
' set the default properties
ReDim combustionenginearray(1)
mvarCustom = False
TL = gVehicleTL 'Note: the TL here is not really necessary.
mvarPowerPlants = "none"
mvarPlantKey = "-1"
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
For i = 1 To UBound(SnorkelMatrix())
    If SnorkelMatrix(i).ID = mvarDatatype Then
        If SnorkelMatrix(i).TL >= mvarTL Then
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
    
    Dim PowerPlantWeight As Single
    Dim PowerPlantVolume As Single
    Dim PowerPlantCost As Single
    Dim QRugMod As Single 'combined quantity and ruggedized multipliers
    Dim RugHitMod As Integer 'ruggedized hit point multiplier
    
    On Error GoTo errorhandler
    
    mvarLocation = GetLocation
    
    'set the ruggedized and quantity multipliers
    If mvarRuggedized Then
        QRugMod = 1.5
        RugHitMod = 2
    Else
        QRugMod = 1
        RugHitMod = 1
    End If
    
    If mvarPlantKey = "-1" Then
        mvarWeight = 0
        mvarCost = 0
        mvarVolume = 0
        mvarSurfaceArea = 0
        mvarHitPoints = 0
        Exit Sub
    Else
        FillCombustionEngineList
        If combustionenginearray(1) = "" Then
            PowerPlants = "none"
        End If
        PowerPlantWeight = Veh.Components(mvarPlantKey).Weight
        PowerPlantVolume = Veh.Components(mvarPlantKey).Volume
        PowerPlantCost = Veh.Components(mvarPlantKey).Cost
    End If
    
    'get base stats
    mvarWeight = SnorkelMatrix(mvarMatrixPos).Weight * PowerPlantWeight
    mvarCost = SnorkelMatrix(mvarMatrixPos).Cost * PowerPlantCost
    mvarVolume = SnorkelMatrix(mvarMatrixPos).Volume * PowerPlantVolume
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
    'get finals
    mvarWeight = Round(QRugMod * mvarWeight, 2)
    mvarCost = Round(QRugMod * mvarCost, 2)
    mvarVolume = Round(QRugMod * mvarVolume, 2)
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    
    Exit Sub
    
errorhandler:
        mvarPlantKey = "-1"
        mvarWeight = 0
        mvarCost = 0
        mvarVolume = 0
        mvarSurfaceArea = 0
        mvarHitPoints = 0
        Exit Sub
End Sub
Public Function FillCombustionEngineList() As String()
    ReDim combustionenginearray(1)
    Dim tempkeys() As String
    Dim i As Long
    Dim current As Long
    
    tempkeys = Veh.KeyManager.GetCurrentPowerSystemKeys
    If tempkeys(1) <> "" Then
        For i = 1 To UBound(tempkeys)
            If TypeOf Veh.Components(tempkeys(i)) Is clsInternalCombustionEngine Then
                current = current + 1
                ReDim Preserve combustionenginearray(current)
                combustionenginearray(current) = Veh.Components(tempkeys(i)).CustomDescription & " (" & tempkeys(i) & ")"
            End If
        Next
    End If
    
    FillCombustionEngineList = combustionenginearray
End Function

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub
