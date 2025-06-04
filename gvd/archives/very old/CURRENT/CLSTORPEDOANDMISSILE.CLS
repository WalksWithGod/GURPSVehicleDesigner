VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTorpedoandMissile"
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
Private mvarClassDebugID As String 'local copy
Private mvarImage As Integer 'local copy
Private mvarKey As String 'local copy
Private mvarName As String 'local copy
Private mvarParent As String 'local copy
Private mvarSelectedImage As Integer 'local copy
Private mvarText As String 'local copy
Private mvarMalf As String 'local copy
Private mvarDamage As String 'local copy
Private mvarAcc As String 'local copy
Private mvarhalfD As String 'local copy
Private mvarMax As String 'local copy
Private mvarRoF As String 'local copy
Private mvarVPS As String 'local copy
Private mvarCPS As String 'local copy
Private mvarMin As String 'local copy
Private mvarSpeed As String 'local copy
Private mvarSkill As String 'local copy
Private mvarGuid As String 'local copy
Private mvarWPS As String 'local copy
Private mvarFuelEnd As String 'local copy
Private mvarSubType As String 'local copy
Private mvarLocation As String 'local copy
Private mvarDR As Variant 'local copy
Private mvarRuggedized As Boolean 'local copy
Private mvarSurfaceArea As Double 'local copy
Private mvarHitPoints As Integer 'local copy
Private mvarTL As Integer 'local copy
Private mvarMinTL As Integer 'local copy
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



Public Property Let FuelEnd(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FuelEnd = 5
    mvarFuelEnd = vData
End Property


Public Property Get FuelEnd() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FuelEnd
    FuelEnd = mvarFuelEnd
End Property



Public Property Let WPS(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WPS = 5
    mvarWPS = vData
End Property


Public Property Get WPS() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WPS
    WPS = mvarWPS
End Property



Public Property Let Guid(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Guid = 5
    mvarGuid = vData
End Property


Public Property Get Guid() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Guid
    Guid = mvarGuid
End Property



Public Property Let Skill(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Skill = 5
    mvarSkill = vData
End Property


Public Property Get Skill() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Skill
    Skill = mvarSkill
End Property



Public Property Let Speed(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Speed = 5
    mvarSpeed = vData
End Property


Public Property Get Speed() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Speed
    Speed = mvarSpeed
End Property



Public Property Let Min(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Min = 5
    mvarMin = vData
End Property


Public Property Get Min() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Min
    Min = mvarMin
End Property



Public Property Let CPS(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPS = 5
    mvarCPS = vData
End Property


Public Property Get CPS() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CPS
    CPS = mvarCPS
End Property



Public Property Let VPS(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VPS = 5
    mvarVPS = vData
End Property


Public Property Get VPS() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VPS
    VPS = mvarVPS
End Property



Public Property Let RoF(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RoF = 5
    mvarRoF = vData
End Property


Public Property Get RoF() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RoF
    RoF = mvarRoF
End Property



Public Property Let Max(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Max = 5
    mvarMax = vData
End Property


Public Property Get Max() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Max
    Max = mvarMax
End Property



Public Property Let halfD(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.halfD = 5
    mvarhalfD = vData
End Property


Public Property Get halfD() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.halfD
    halfD = mvarhalfD
End Property



Public Property Let Acc(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Acc = 5
    mvarAcc = vData
End Property


Public Property Get Acc() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Acc
    Acc = mvarAcc
End Property



Public Property Let Damage(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Damage = 5
    mvarDamage = vData
End Property


Public Property Get Damage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Damage
    Damage = mvarDamage
End Property



Public Property Let Malf(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Malf = 5
    mvarMalf = vData
End Property


Public Property Get Malf() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Malf
    Malf = mvarMalf
End Property



Public Function PropQuery(ByVal Row As Integer) As Integer
End Function

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



Public Property Let SelectedImage(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SelectedImage = 5
    mvarSelectedImage = vData
End Property


Public Property Get SelectedImage() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedImage
    SelectedImage = mvarSelectedImage
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



Public Property Let Name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name = 5
    mvarName = vData
End Property


Public Property Get Name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name
    Name = mvarName
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



Public Property Let Image(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Image = 5
    mvarImage = vData
End Property


Public Property Get Image() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Image
    Image = mvarImage
End Property



Public Property Get ClassDebugID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClassDebugID
    ClassDebugID = mvarClassDebugID
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



