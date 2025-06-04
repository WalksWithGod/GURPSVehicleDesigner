VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommunicator"
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
Private mvarReceiveOnly As String 'local copy
Private mvarTightBeamRadio As String 'local copy
Private mvarVLFRadio As String 'local copy
Private mvarCellularPhone As String 'local copy
Private mvarRadioJammer As String 'local copy
Private mvarSensitivity As String 'local copy
Private mvarLaser As String 'local copy
Private mvarNeutrino As String 'local copy
Private mvarGravityRipple As String 'local copy
Private mvarFTL As String 'local copy
Private mvarTL As Integer 'local copy
Private mvarWeight As Double 'local copy
Private mvarCost As Double 'local copy
Private mvarRange As String 'local copy
Private mvarPowerReqt As Double 'local copy
Private mvarVolume As Double 'local copy
Private mvarScrambler As String 'local copy
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



Public Property Let Scrambler(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Scrambler = 5
    mvarScrambler = vData
End Property


Public Property Get Scrambler() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Scrambler
    Scrambler = mvarScrambler
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



Public Property Let FTL(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FTL = 5
    mvarFTL = vData
End Property


Public Property Get FTL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FTL
    FTL = mvarFTL
End Property



Public Property Let GravityRipple(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GravityRipple = 5
    mvarGravityRipple = vData
End Property


Public Property Get GravityRipple() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GravityRipple
    GravityRipple = mvarGravityRipple
End Property



Public Property Let Neutrino(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Neutrino = 5
    mvarNeutrino = vData
End Property


Public Property Get Neutrino() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Neutrino
    Neutrino = mvarNeutrino
End Property



Public Property Let Laser(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Laser = 5
    mvarLaser = vData
End Property


Public Property Get Laser() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Laser
    Laser = mvarLaser
End Property



Public Property Let Sensitivity(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sensitivity = 5
    mvarSensitivity = vData
End Property


Public Property Get Sensitivity() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sensitivity
    Sensitivity = mvarSensitivity
End Property



Public Property Let RadioJammer(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RadioJammer = 5
    mvarRadioJammer = vData
End Property


Public Property Get RadioJammer() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RadioJammer
    RadioJammer = mvarRadioJammer
End Property



Public Property Let CellularPhone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CellularPhone = 5
    mvarCellularPhone = vData
End Property


Public Property Get CellularPhone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CellularPhone
    CellularPhone = mvarCellularPhone
End Property



Public Property Let VLFRadio(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VLFRadio = 5
    mvarVLFRadio = vData
End Property


Public Property Get VLFRadio() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VLFRadio
    VLFRadio = mvarVLFRadio
End Property



Public Property Let TightBeamRadio(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TightBeamRadio = 5
    mvarTightBeamRadio = vData
End Property


Public Property Get TightBeamRadio() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TightBeamRadio
    TightBeamRadio = mvarTightBeamRadio
End Property



Public Property Let ReceiveOnly(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReceiveOnly = 5
    mvarReceiveOnly = vData
End Property


Public Property Get ReceiveOnly() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReceiveOnly
    ReceiveOnly = mvarReceiveOnly
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



