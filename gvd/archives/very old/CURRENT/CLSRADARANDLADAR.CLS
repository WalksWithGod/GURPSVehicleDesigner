VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRadarandLadar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 1


#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    Private mlClassDebugID As Long
#End If
Private mvarNoTargeting As String 'local copy
Private mvarSurfaceSearch As String 'local copy
Private mvarAirSearch As String 'local copy
Private mvarLowResImagingRadar As String 'local copy
Private mvarHighResImagingRadar As String 'local copy
Private mvarLadar As String 'local copy
Private mvarActiveElectromagneticSensorArray As String 'local copy
Private mvarFTLRadar As String 'local copy
Private mvarRange As String 'local copy
Private mvarDirectionFacing As String 'local copy
Private mvarScanRating As String 'local copy
Private mvarTL As Integer 'local copy
Private mvarWeight As Double 'local copy
Private mvarVolume As Double 'local copy
Private mvarCost As Double 'local copy
Private mvarPowerReqt As Double 'local copy
Private mvarSubType As String 'local copy
Private mvarLocation As String 'local copy
Private mvarParent As String 'local copy
Private mvarKey As String 'local copy
Private mvarDR As Variant 'local copy
Private mvarRuggedized As Boolean 'local copy
Private mvarSurfaceArea As Double 'local copy
Private mvarHitPoints As Integer 'local copy
Private mvarMinTL As Integer 'local copy
Private mvarText As String 'local copy
Public Property Let Text(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Text = 5
    mvarText = vData
End Property


Public Property Get Text() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Text
    Text = mvarText
End Property



Public Property Let MinTL(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MinTL = 5
    mvarMinTL = vData
End Property


Public Property Get MinTL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinTL
    MinTL = mvarMinTL
End Property



Public Property Let HitPoints(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HitPoints = 5
    mvarHitPoints = vData
End Property


Public Property Get HitPoints() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HitPoints
    HitPoints = mvarHitPoints
End Property



Public Property Let SurfaceArea(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SurfaceArea = 5
    mvarSurfaceArea = vData
End Property


Public Property Get SurfaceArea() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SurfaceArea
    SurfaceArea = mvarSurfaceArea
End Property



Public Property Let Ruggedized(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ruggedized = 5
    mvarRuggedized = vData
End Property


Public Property Get Ruggedized() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ruggedized
    Ruggedized = mvarRuggedized
End Property



Public Property Let DR(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR = 5
    mvarDR = vData
End Property


Public Property Set DR(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DR = Form1
    Set mvarDR = vData
End Property


Public Property Get DR() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DR
    If IsObject(mvarDR) Then
        Set DR = mvarDR
    Else
        DR = mvarDR
    End If
End Property



Public Property Let Key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property



Public Property Let Parent(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Parent = 5
    mvarParent = vData
End Property


Public Property Get Parent() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parent
    Parent = mvarParent
End Property



Public Property Let Location(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Location = 5
    mvarLocation = vData
End Property


Public Property Get Location() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Location
    Location = mvarLocation
End Property



Public Property Let SubType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubType = 5
    mvarSubType = vData
End Property


Public Property Get SubType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubType
    SubType = mvarSubType
End Property



Public Property Let PowerReqt(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Power = 5
    mvarPowerReqt = vData
End Property


Public Property Get PowerReqt() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Power
    PowerReqt = mvarPowerReqt
End Property



Public Property Let Cost(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cost = 5
    mvarCost = vData
End Property


Public Property Get Cost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cost
    Cost = mvarCost
End Property



Public Property Let Volume(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Volume = 5
    mvarVolume = vData
End Property


Public Property Get Volume() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Volume
    Volume = mvarVolume
End Property



Public Property Let Weight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Weight = 5
    mvarWeight = vData
End Property


Public Property Get Weight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Weight
    Weight = mvarWeight
End Property



Public Property Let TL(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TL = 5
    mvarTL = vData
End Property


Public Property Get TL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    TL = mvarTL
End Property



Public Property Let ScanRating(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ScanRating = 5
    mvarScanRating = vData
End Property


Public Property Get ScanRating() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ScanRating
    ScanRating = mvarScanRating
End Property



Public Property Let DirectionFacing(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DirectionFacing = 5
    mvarDirectionFacing = vData
End Property


Public Property Get DirectionFacing() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DirectionFacing
    DirectionFacing = mvarDirectionFacing
End Property



Public Property Let Range(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Range = 5
    mvarRange = vData
End Property


Public Property Get Range() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Range
    Range = mvarRange
End Property



Public Property Let FTLRadar(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FTLRadar = 5
    mvarFTLRadar = vData
End Property


Public Property Get FTLRadar() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FTLRadar
    FTLRadar = mvarFTLRadar
End Property



Public Property Let ActiveElectromagneticSensorArray(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ActiveElectromagneticSensorArray = 5
    mvarActiveElectromagneticSensorArray = vData
End Property


Public Property Get ActiveElectromagneticSensorArray() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ActiveElectromagneticSensorArray
    ActiveElectromagneticSensorArray = mvarActiveElectromagneticSensorArray
End Property



Public Property Let Ladar(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ladar = 5
    mvarLadar = vData
End Property


Public Property Get Ladar() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ladar
    Ladar = mvarLadar
End Property



Public Property Let HighResImagingRadar(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HighResImagingRadar = 5
    mvarHighResImagingRadar = vData
End Property


Public Property Get HighResImagingRadar() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HighResImagingRadar
    HighResImagingRadar = mvarHighResImagingRadar
End Property



Public Property Let LowResImagingRadar(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LowResImagingRadar = 5
    mvarLowResImagingRadar = vData
End Property


Public Property Get LowResImagingRadar() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LowResImagingRadar
    LowResImagingRadar = mvarLowResImagingRadar
End Property



Public Property Let AirSearch(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AirSearch = 5
    mvarAirSearch = vData
End Property


Public Property Get AirSearch() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AirSearch
    AirSearch = mvarAirSearch
End Property



Public Property Let SurfaceSearch(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SurfaceSearch = 5
    mvarSurfaceSearch = vData
End Property


Public Property Get SurfaceSearch() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SurfaceSearch
    SurfaceSearch = mvarSurfaceSearch
End Property



Public Property Let NoTargeting(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoTargeting = 5
    mvarNoTargeting = vData
End Property


Public Property Get NoTargeting() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoTargeting
    NoTargeting = mvarNoTargeting
End Property





Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
End Sub


Private Sub Class_Terminate()
    'the class is being destroyed
    #If DebugMode Then
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub



