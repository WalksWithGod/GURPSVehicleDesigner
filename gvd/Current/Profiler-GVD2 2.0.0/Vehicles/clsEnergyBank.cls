VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnergyBank"
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
Private mvarCost As Double
Private mvarVolume As Double
Private mvarPoweredRewinder As Boolean
Private mvarLocation As String
Private mvarParent As String
Private mvarKey As String
Private mvarDR As Long
Private mvarRuggedized As Boolean
Private mvarSurfaceArea As Double
Private mvarHitPoints As Double
Private mvarPowerConsumed As Single
Private mvarPowerReqt As Double

Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarQuantity As Integer
Private mvarOutput As Single
Private mvarDesiredOutput As Single
Private mvarMotorST As Single

Private mvarCellType As String

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
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrintOutput = 5
    mvarPrintOutput = vdata
End Property


Public Property Get PrintOutput() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrintOutput
    PrintOutput = mvarPrintOutput
End Property

Public Property Let DesiredOutput(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DesiredOutput = 5

    mvarDesiredOutput = vdata
    If mvarZZInit = 0 Then
         Exit Property
    End If
    mvarCellType = "custom"
End Property

Public Property Get DesiredOutput() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DesiredOutput
    DesiredOutput = mvarDesiredOutput
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

        If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = Leg) Or (InstallPoint = equipmentPod) Or (InstallPoint = Module) Then
            TempCheck = True
        Else
            InfoPrint 1, "Energy Banks must be placed in Body, Superstructure, Pod, equipment Pod, Turret, Popturret, Wing, Arm, Leg or Module."
            TempCheck = False
        End If

If TempCheck Then
     SetLogicalParent
End If
LocationCheck = TempCheck
End Function


Private Function GetLocation() As String
    On Error Resume Next
    If mvarLogicalParent = "" Then
         SetLogicalParent
    End If
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

Public Property Let CellType(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CellType = 5
    mvarCellType = vdata
    If mvarZZInit = 0 Then
         Exit Property
    End If
    'user has selected a cell type so we will determine the capacity for them
    If mvarCellType <> "custom" Then

        Select Case mvarCellType
            Case "AA"
                mvarDesiredOutput = 1.8
            Case "A"
                mvarDesiredOutput = 18
            Case "B"
                mvarDesiredOutput = 180
            Case "C"
                mvarDesiredOutput = 1800
            Case "D"
                mvarDesiredOutput = 18000
            Case "E"
                mvarDesiredOutput = 180000
        End Select
        '//if rechargeable, then output is halved
        If mvarDatatype = RechargeablePowerCell Then
            mvarDesiredOutput = mvarDesiredOutput / 2
        End If

        If mvarTL < 7 Then
            'leave it alone
        Else '//increase the output with tech level
            mvarDesiredOutput = mvarDesiredOutput * (mvarTL - 6)
        End If

    End If

End Property


Public Property Get CellType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Location
    CellType = mvarCellType
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


Public Property Let PoweredRewinder(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.poweredrewind = 5
    mvarPoweredRewinder = vdata
    If mvarZZInit = 0 Then
         Exit Property
    End If

    If mvarPoweredRewinder = False Then
         MotorST = 0
    End If

End Property


Public Property Get PoweredRewinder() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Recharging
    PoweredRewinder = mvarPoweredRewinder
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


' POWER REQT IS ONLY FOR RECHARGEABLES, ADVANCED BATTERY,FLYWHEELS and LEAD ACID
Public Property Let PowerReqt(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Power = 5
    mvarPowerReqt = vdata
End Property


Public Property Get PowerReqt() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Power
    PowerReqt = mvarPowerReqt
End Property

Public Property Let Output(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Output = 5
    mvarOutput = vdata
End Property


Public Property Get Output() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Output
    Output = mvarOutput
End Property

Public Property Let PowerConsumed(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PowerConsumed = 5
    mvarPowerConsumed = vdata
End Property


Public Property Get PowerConsumed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PowerConsumed
    PowerConsumed = mvarPowerConsumed
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

Public Property Let MotorST(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MotorST = 5
    mvarMotorST = vdata
End Property


Public Property Get MotorST() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MotorST
    MotorST = mvarMotorST
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
    If mvarZZInit = 0 Then
         Exit Property
    End If
    CellType = mvarCellType
    GetMatrixIndex
End Property


Public Property Get TL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    TL = mvarTL
End Property



Private Sub Class_Initialize()
'set the default dimension of the keychain to 1 element
ReDim mvarPowerConsumptionKeyChain(1)

' set the default properties
mvarCustom = False
TL = gVehicleTL
mvarRuggedized = False
mvarQuantity = 1
mvarDesiredOutput = 2000


End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case ClockWork

    Case LeadAcidBattery

    Case AdvancedBattery

    Case Flywheel

    Case RechargeablePowerCell
        mvarCellType = "custom"
    Case PowerCell
       mvarCellType = "custom"
End Select


End Sub

Public Sub GetMatrixIndex()
Dim i As Integer
If mvarDatatype = 0 Then
     Exit Sub
End If
mvarMatrixPos = 0 'init the counter
For i = 1 To UBound(EnergyBankMatrix())
    If EnergyBankMatrix(i).ID = mvarDatatype Then
        If EnergyBankMatrix(i).TL >= mvarTL Then
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
If mvarMatrixPos = 0 Then
     Exit Sub
End If

Dim TrueWeight As Single
Dim TempCost As Single
Dim TempFuel As Single
Dim TempVolume As Single
Dim i As Long
Dim TempPowerConsumption As Single
Dim QRugMod As Single 'combined quantity and ruggedized multipliers
Dim RugHitMod As Integer 'ruggedized hit point multiplier
Dim sPrint1 As String
Dim sPrint2 As String
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

'determine the weight
TrueWeight = mvarDesiredOutput * EnergyBankMatrix(mvarMatrixPos).Weight

'Find the volume
TempVolume = TrueWeight / EnergyBankMatrix(mvarMatrixPos).Volume

'find cost
TempCost = TrueWeight * EnergyBankMatrix(mvarMatrixPos).Cost

'calc stats for Powered rewindable clockwork
If (mvarDatatype = ClockWork) And (mvarPoweredRewinder) Then
    mvarMotorST = Round(EnergyBankMatrix(mvarMatrixPos).EffectiveST * mvarDesiredOutput, 0)
    TempVolume = EnergyBankMatrix(mvarMatrixPos).PoweredClockCost * TempVolume
    TempCost = EnergyBankMatrix(mvarMatrixPos).PoweredClockCost * TempCost
    TrueWeight = EnergyBankMatrix(mvarMatrixPos).PoweredClockCost * TrueWeight
End If

'get base stats
mvarWeight = TrueWeight
mvarCost = TempCost
mvarVolume = TempVolume
mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)

'get finals
mvarWeight = Round(QRugMod * mvarWeight, 2)
mvarCost = Round(QRugMod * mvarCost, 2)
mvarVolume = Round(QRugMod * mvarVolume, 2)
mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
mvarOutput = mvarDesiredOutput * mvarQuantity


'produce the print output
 If mvarRuggedized Then
     sPrint1 = "ruggedized "
 Else
     sPrint1 = ""
 End If

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

If mvarPoweredRewinder Then
    sPrint2 = ", with ST " + Format(mvarMotorST) + " powered rewinder"
End If

 mvarPrintOutput = NumericToString(mvarQuantity) + " TL" & mvarTL & " " + Format(mvarDesiredOutput, p_sFormat) + " kWs " + sPrint1 + mvarCustomDescription + sPrintPlural + sPrint2 _
     + " (" + mvarLocation + ", HP " & mvarHitPoints & sPrintPlural3 + ", " + sPrintPlural4 + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ")." + mvarComment

End Sub


Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub

