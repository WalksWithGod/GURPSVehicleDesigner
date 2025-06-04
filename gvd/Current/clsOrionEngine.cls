VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrionEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarTL As Integer
Private mvarWeight As Double
Private mvarVolume As Double
Private mvarMotiveThrust As Single
Private mvarCost As Double
Private mvarPulseRate As Single
Private mvarThrustTime As Single
Private mvarBombSize As Single
Private mvarLiftEngine As Boolean
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
Private mvarQuantity As Integer
Private mvarNumBombs As Long
Private mvarBombCost As Single
Private mvarBombVolume As Single
Private mvarBombWeight As Single
Private mvarBombSurfaceArea As Single

Private mvarImage As Integer
Private mvarSelectedImage As Integer
Private mvarComment As String
Private mvarCName As String
Private mvarMatrixPos As Long
Private mvarMatrixPos2 As Long
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



Public Property Let MatrixPos2(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MatrixPos2 = 5
    mvarMatrixPos2 = vdata
End Property


Public Property Get MatrixPos2() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MatrixPos2
    MatrixPos2 = mvarMatrixPos2
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

        If InstallPoint = Body Then
            TempCheck = True
        Else
            InfoPrint 1, "Orion Engines must be placed in Body"
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


Public Property Let LiftEngine(ByVal vdata As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LiftEngine = 5
    mvarLiftEngine = vdata
End Property


Public Property Get LiftEngine() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LiftEngine
    LiftEngine = mvarLiftEngine
End Property



Public Property Let PulseRate(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PulseRate = 5
    mvarPulseRate = vdata
End Property


Public Property Get PulseRate() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PulseRate
    PulseRate = mvarPulseRate
End Property

Public Property Let ThrustTime(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ThrustTime = 5
    mvarThrustTime = vdata
End Property


Public Property Get ThrustTime() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ThrustTime
    ThrustTime = mvarThrustTime
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

Public Property Get BombSize() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BombSize
    BombSize = mvarBombSize
End Property

Public Property Let BombSize(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BombSize = 5
    mvarBombSize = vdata
End Property

Public Property Get BombCost() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BombCost
    BombCost = mvarBombCost
End Property

Public Property Let BombCost(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BombCost = 5
    mvarBombCost = vdata
End Property


Public Property Get BombWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BombWeight
    BombWeight = mvarBombWeight
End Property

Public Property Let BombWeight(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BombWeight = 5
    mvarBombWeight = vdata
End Property


Public Property Get BombVolume() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BombVolume
    BombVolume = mvarBombVolume
End Property

Public Property Let BombVolume(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BombVolume = 5
    mvarBombVolume = vdata
End Property

Public Property Get BombSurfaceArea() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BombSurfaceArea
    BombSurfaceArea = mvarBombSurfaceArea
End Property

Public Property Let BombSurfaceArea(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BombSurfaceArea = 5
    mvarBombSurfaceArea = vdata
End Property


Public Property Let NumBombs(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumBombs = 5
    mvarNumBombs = vdata
End Property


Public Property Get NumBombs() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumBombs
    NumBombs = mvarNumBombs
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
mvarRuggedized = False
mvarBombSize = 100
mvarQuantity = 1
mvarPulseRate = 10
mvarNumBombs = 1000

End Sub


Public Sub GetMatrixIndex()
Dim i As Integer
If mvarDatatype = 0 Then Exit Sub
mvarMatrixPos = 0 'init the counter
For i = 1 To UBound(OrionEngineMatrix())
    If OrionEngineMatrix(i).ID = mvarDatatype Then
        If OrionEngineMatrix(i).TL >= mvarTL Then
            mvarMatrixPos = i
            Exit For
        Else
            mvarMatrixPos = i
        End If
    End If
Next

mvarMatrixPos2 = 0 'init the counter
For i = 1 To UBound(ThrustBombMatrix())
    If ThrustBombMatrix(i).ID = ThrustBomb Then
        If ThrustBombMatrix(i).TL >= mvarTL Then
            mvarMatrixPos2 = i
            Exit For
        Else
            mvarMatrixPos2 = i
        End If
    End If
Next
End Sub


Public Sub StatsUpdate()
    mvarZZInit = 1
    If mvarMatrixPos = 0 Then Exit Sub
    
    Dim QRugMod As Single 'combined quantity and ruggedized multipliers
    Dim RugHitMod As Integer 'ruggedized hit point multiplier
    Dim sPrint1 As String
    Dim sPrint2 As String
    Dim sPrint3 As String
    Dim sPrint4 As String
    Dim sPrintPlural As String
    Dim sPrintPlural2 As String
    Dim sPrintPlural3 As String
    Dim sPrintPlural4 As String
    
    mvarLocation = GetLocation
    
    'set the ruggedized and quantity multipliers
    If mvarRuggedized Then
        QRugMod = 1.5 * mvarQuantity
        RugHitMod = 2
    Else
        QRugMod = 1 * mvarQuantity
        RugHitMod = 1
    End If
    
    'Get intial weight/cost /volume for Engine
    'NOTE: There are no mods for Liftengines
    mvarWeight = OrionEngineMatrix(mvarMatrixPos).Weight1 + (OrionEngineMatrix(mvarMatrixPos).Weight2 * mvarBombSize) + (mvarPulseRate * OrionEngineMatrix(mvarMatrixPos).Weight3)
    mvarCost = mvarWeight * (mvarPulseRate ^ 1 / 2)
    mvarVolume = mvarWeight / OrionEngineMatrix(mvarMatrixPos).Volume
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)
            
    'Get weight/cost/volume for bombs
    mvarBombWeight = Round(mvarQuantity * mvarNumBombs * (ThrustBombMatrix(mvarMatrixPos2).Weight1 + (ThrustBombMatrix(mvarMatrixPos2).Weight2 * mvarBombSize)), 2)
    mvarBombVolume = Round(mvarQuantity * mvarNumBombs * (mvarBombWeight / ThrustBombMatrix(mvarMatrixPos2).Volume), 2)
    mvarBombCost = Round(mvarQuantity * (mvarBombWeight * ThrustBombMatrix(mvarMatrixPos2).Cost), 2)
    mvarBombSurfaceArea = Round(CalcSurfaceArea(mvarBombVolume), 2)
    
    'get stats BEFORE adding in weight/cost/volume of the bombs
    mvarWeight = Round(QRugMod * mvarWeight, 2)
    mvarCost = Round(QRugMod * mvarCost, 2)
    mvarVolume = Round(QRugMod * mvarVolume, 2)
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    
    'now add in weight/volume/cost of the bombs
    mvarWeight = mvarWeight + mvarBombWeight
    mvarCost = mvarCost + mvarBombCost
    mvarVolume = mvarVolume + mvarBombVolume
    mvarSurfaceArea = mvarSurfaceArea + mvarBombSurfaceArea
    
    'Get motive thrust and thrust time
    mvarMotiveThrust = mvarQuantity * 400000 * mvarBombSize * mvarPulseRate
    mvarThrustTime = mvarNumBombs / mvarPulseRate
    
    'produce the print output
    If mvarRuggedized Then
        sPrint1 = "ruggedized "
    Else
        sPrint1 = ""
    End If
    
    sPrint2 = Format(mvarPulseRate, p_sFormat) + "bps " + Format(mvarBombSize, p_sFormat) + " kt "
    If mvarLiftEngine Then
        sPrint3 = " lift engine"
    Else
        sPrint3 = ""
    End If
    
    sPrint4 = " its " + Format(mvarNumBombs) + " bombs yeilds a thrust time of " + Format(mvarThrustTime, p_sFormat) + " seconds"
    
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
    
    mvarPrintOutput = NumericToString(mvarQuantity) + " TL" & mvarTL & " " + sPrint2 + sPrint1 + mvarCustomDescription + sPrintPlural + sPrint3 + sPrintPlural2 + Format(mvarMotiveThrust, p_sFormat) + " lbs. motive thrust" _
        + " (" + mvarLocation + ", HP " & mvarHitPoints & sPrintPlural3 + ", " + sPrintPlural4 + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ")" + sPrint4 + "." + mvarComment
    

End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub

