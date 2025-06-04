VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntertainmentFacility"
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
Private mvarCost As Double
Private mvarPowerReqt As Double
Private mvarLocation As String
Private mvarParent As String
Private mvarKey As String
Private mvarDR As Long
Private mvarRuggedized As Boolean
Private mvarSurfaceArea As Double
Private mvarHitPoints As Double

Private mvarFloorArea As Single
Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarQuantity As Integer


Private mvarImage As Integer
Private mvarSelectedImage As Integer
Private mvarComment As String
Private mvarCName As String
Private mvarMatrixPos As Long

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
    mvarPrintOutput = vdata
End Property

Public Property Get PrintOutput() As String
   PrintOutput = mvarPrintOutput
End Property

Public Property Let MatrixPos(ByVal vdata As Long)
    mvarMatrixPos = vdata
End Property

Public Property Get MatrixPos() As Long
   MatrixPos = mvarMatrixPos
End Property

Public Property Let CName(ByVal vdata As String)
    mvarCName = vdata
End Property

Public Property Get CName() As String
    CName = mvarCName
End Property

Public Property Let Comment(ByVal vdata As String)
    mvarComment = vdata
End Property

Public Property Get Comment() As String
    Comment = mvarComment
End Property

Public Property Let SelectedImage(ByVal vdata As Integer)
  mvarSelectedImage = vdata
End Property

Public Property Get SelectedImage() As Integer
   SelectedImage = mvarSelectedImage
End Property

Public Property Let Image(ByVal vdata As Integer)
   mvarImage = vdata
End Property

Public Property Get Image() As Integer
    Image = mvarImage
End Property

Public Function LocationCheck() As Boolean
Dim TempCheck As Boolean
Dim InstallPoint As Integer
InstallPoint = Veh.Components(mvarParent).Datatype

        If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Pod) Then
            TempCheck = True
        Else
            InfoPrint 1, "Entertainment Facilities must be placed in Body, Superstructure, Turret, Popturret or Pod."
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
    mvarQuantity = vdata
End Property

Public Property Get Quantity() As Integer
    Quantity = mvarQuantity
End Property

Public Property Let Custom(ByVal vdata As Boolean)
    mvarCustom = vdata
End Property

Public Property Get Custom() As Boolean
    Custom = mvarCustom
End Property

Public Property Let CustomDescription(ByVal vdata As String)
    mvarCustomDescription = vdata
End Property

Public Property Get CustomDescription() As String
    CustomDescription = mvarCustomDescription
End Property

Public Property Let Description(ByVal vdata As String)
  mvarDescription = vdata
End Property

Public Property Get Description() As String
    Description = mvarDescription
End Property

Public Property Let ParentDatatype(ByVal vdata As Integer)
    mvarParentDatatype = vdata
End Property

Public Property Get ParentDatatype() As Integer
    ParentDatatype = mvarParentDatatype
End Property

Public Property Let Datatype(ByVal vdata As Integer)
    mvarDatatype = vdata
End Property

Public Property Get Datatype() As Integer
    Datatype = mvarDatatype
End Property

Public Property Let HitPoints(ByVal vdata As Double)
    mvarHitPoints = vdata
End Property

Public Property Get HitPoints() As Double
    HitPoints = mvarHitPoints
End Property

Public Property Let FloorArea(ByVal vdata As Single)
    mvarFloorArea = vdata
End Property

Public Property Get FloorArea() As Single
    FloorArea = mvarFloorArea
End Property

Public Property Let SurfaceArea(ByVal vdata As Double)
    mvarSurfaceArea = vdata
End Property

Public Property Get SurfaceArea() As Double
    SurfaceArea = mvarSurfaceArea
End Property

Public Property Let Ruggedized(ByVal vdata As Boolean)
    mvarRuggedized = vdata
End Property

Public Property Get Ruggedized() As Boolean
    Ruggedized = mvarRuggedized
End Property

Public Property Let DR(ByVal vdata As Long)
    mvarDR = vdata
End Property

Public Property Get DR() As Long
     DR = mvarDR
End Property

Public Property Let Key(ByVal vdata As String)
    mvarKey = vdata
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Let Parent(ByVal vdata As String)
    mvarParent = vdata
End Property

Public Property Get Parent() As String
    Parent = mvarParent
End Property

Public Property Let Location(ByVal vdata As String)
    mvarLocation = vdata
End Property

Public Property Get Location() As String
    Location = mvarLocation
End Property

Public Property Let PowerReqt(ByVal vdata As Double)
    mvarPowerReqt = vdata
End Property

Public Property Get PowerReqt() As Double
    PowerReqt = mvarPowerReqt
End Property

Public Property Let Cost(ByVal vdata As Double)
    mvarCost = vdata
End Property

Public Property Get Cost() As Double
    Cost = mvarCost
End Property

Public Property Let Volume(ByVal vdata As Double)
    mvarVolume = vdata
End Property

Public Property Get Volume() As Double
    Volume = mvarVolume
End Property

Public Property Let Weight(ByVal vdata As Double)
    mvarWeight = vdata
End Property

Public Property Get Weight() As Double
    Weight = mvarWeight
End Property

Public Property Let TL(ByVal vdata As Integer)
    mvarTL = vdata
    GetMatrixIndex
End Property

Public Property Get TL() As Integer
    TL = mvarTL
End Property


Private Sub Class_Initialize()
 
' set the default properties
mvarCustom = False
TL = gVehicleTL
mvarRuggedized = False
mvarQuantity = 1

End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case Stage
        
        mvarFloorArea = 100
    Case Hall
        
        mvarFloorArea = 600
    Case BarRoom
        
        mvarFloorArea = 500
    Case ConferenceRoom
        
        mvarFloorArea = 400
    Case MovieScreenandProjector
        
    Case MovieScreenandProjectorSmall
       
    Case HoloventureZone
        mvarFloorArea = 100

End Select

End Sub

Public Sub GetMatrixIndex()
Dim i As Integer
If mvarDatatype = 0 Then Exit Sub
mvarMatrixPos = 0 'init the counter
For i = 1 To UBound(EntertainmentMatrix())
    If EntertainmentMatrix(i).ID = mvarDatatype Then
        If EntertainmentMatrix(i).TL >= mvarTL Then
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

Dim AreaMod As Integer
Dim QRugMod As Single 'combined quantity and ruggedized multipliers
Dim RugHitMod As Integer 'ruggedized hit point multiplier
Dim sPrint1 As String
Dim sPrintPlural As String
Dim sPrintPlural2 As String
Dim sPrintPlural3 As String

mvarLocation = GetLocation

'set the ruggedized and quantity multipliers
If mvarRuggedized Then
    QRugMod = 1.5 * mvarQuantity
    RugHitMod = 2
Else
    QRugMod = 1 * mvarQuantity
    RugHitMod = 1
End If

Select Case mvarDatatype
Case Stage, BarRoom, Hall, ConferenceRoom, HoloventureZone
    AreaMod = mvarFloorArea / 100

    mvarWeight = AreaMod * EntertainmentMatrix(mvarMatrixPos).Weight
    mvarCost = AreaMod * EntertainmentMatrix(mvarMatrixPos).Cost
    mvarVolume = AreaMod * EntertainmentMatrix(mvarMatrixPos).Volume
    mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarFloorArea)

    If EntertainmentMatrix(mvarMatrixPos).Power = 0 Then
        mvarPowerReqt = 0
    Else
        mvarPowerReqt = AreaMod * EntertainmentMatrix(mvarMatrixPos).Power
    End If
    'get finals
    mvarWeight = Round(mvarQuantity * mvarWeight, 2)
    mvarCost = Round(mvarQuantity * mvarCost, 2)
    mvarVolume = Round(mvarQuantity * mvarVolume, 2)
    mvarSurfaceArea = Round(mvarQuantity * mvarFloorArea, 2)
    mvarPowerReqt = Round(mvarQuantity * mvarPowerReqt, 2)
Case Else

    mvarWeight = EntertainmentMatrix(mvarMatrixPos).Weight
    mvarCost = EntertainmentMatrix(mvarMatrixPos).Cost
    mvarVolume = EntertainmentMatrix(mvarMatrixPos).Volume
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)

    mvarPowerReqt = EntertainmentMatrix(mvarMatrixPos).Power

    'get finals
    mvarWeight = Round(QRugMod * mvarWeight, 2)
    mvarCost = Round(QRugMod * mvarCost, 2)
    mvarVolume = Round(QRugMod * mvarVolume, 2)
    mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
    mvarPowerReqt = Round(mvarQuantity * mvarPowerReqt, 2)
End Select

'print output
If mvarRuggedized Then
    sPrint1 = "ruggedized "
Else
    sPrint1 = ""
End If

If mvarFloorArea <> 0 Then
        sPrint1 = sPrint1 + Format(mvarFloorArea) + " sq ft "
End If

If mvarQuantity > 1 Then
    sPrintPlural = "s"
    sPrintPlural2 = " each"
    sPrintPlural3 = " total of "
Else
    sPrintPlural = ""
    sPrintPlural2 = ""
    sPrintPlural3 = ""
End If
mvarPrintOutput = NumericToString(mvarQuantity) + " TL" & mvarTL & " " + sPrint1 + mvarCustomDescription + sPrintPlural _
    + " (" + mvarLocation + ", HP " & mvarHitPoints & sPrintPlural2 + ", " + sPrintPlural3 + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ", " + Format(mvarPowerReqt, p_sFormat) + " kW)." + mvarComment

End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub
