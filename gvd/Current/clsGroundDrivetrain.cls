VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGroundDrivetrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit


Private mvarTL As Integer
Private mvarMotivePower As Single
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
Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean

 'holds the index for DrivetrainMatrix
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
    If mvarZZInit = 0 Then Exit Property
On Error GoTo errorhandler

Dim legarray() As String
Dim UpdateSiblings As Boolean
Dim NumSiblings As Long
Dim i As Long

'get the keys for our sibling leg drivetrains
If mvarDatatype = LegDrivetrain Then
    legarray = Veh.KeyManager.GetCurrentLegDrivetrainKeys
    If legarray(1) <> "" Then
        UpdateSiblings = True
        NumSiblings = UBound(legarray)
    End If
End If

'update the siblings as well
If UpdateSiblings Then
    For i = 1 To NumSiblings
        If Veh.Components(legarray(i)).Ruggedized <> vdata Then
            Veh.Components(legarray(i)).Ruggedized = vdata
            Veh.Components(legarray(i)).StatsUpdate
        End If
    Next
End If
Exit Property
errorhandler:
'when loading a save vehicle, it will try to update the sibling
'which has not yet been created
    Exit Property
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
    If mvarZZInit = 0 Then Exit Property
    
On Error GoTo errorhandler

Dim legarray() As String
Dim UpdateSiblings As Boolean
Dim NumSiblings As Long
Dim i As Long

'get the keys for our sibling leg drivetrains
If mvarDatatype = LegDrivetrain Then
    legarray = Veh.KeyManager.GetCurrentLegDrivetrainKeys
    If legarray(1) <> "" Then
        UpdateSiblings = True
        NumSiblings = UBound(legarray)
    End If
End If


'update the siblings as well
If UpdateSiblings Then
    For i = 1 To NumSiblings
        If Veh.Components(legarray(i)).DR <> vdata Then
            Veh.Components(legarray(i)).DR = vdata
            Veh.Components(legarray(i)).StatsUpdate
        End If
    Next
End If
Exit Property
errorhandler:
'when loading a save vehicle, it will try to update the sibling
'which has not yet been created
    Exit Property
    
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

Public Property Let PowerReqt(ByVal vdata As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PowerReqt = 5
    mvarPowerReqt = vdata
End Property


Public Property Get PowerReqt() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PowerReqt
    PowerReqt = mvarPowerReqt
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



Public Property Let MotivePower(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MotivePower = 5
    mvarMotivePower = vdata
    If mvarZZInit = 0 Then Exit Property
On Error GoTo errorhandler

Dim legarray() As String
Dim UpdateSiblings As Boolean
Dim NumSiblings As Long
Dim i As Long

'get the keys for our sibling leg drivetrains
If mvarDatatype = LegDrivetrain Then
    legarray = Veh.KeyManager.GetCurrentLegDrivetrainKeys
    If legarray(1) <> "" Then
        UpdateSiblings = True
        NumSiblings = UBound(legarray)
    End If
End If

'update the siblings as well
If UpdateSiblings Then
    For i = 1 To NumSiblings
        If Veh.Components(legarray(i)).MotivePower <> vdata Then
            Veh.Components(legarray(i)).MotivePower = vdata
            Veh.Components(legarray(i)).StatsUpdate
        End If
    Next
End If
Exit Property
errorhandler:
'when loading a save vehicle, it will try to update the sibling
'which has not yet been created
    Exit Property
End Property


Public Property Get MotivePower() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MotivePower
    MotivePower = mvarMotivePower
End Property


Public Property Let TL(ByVal vdata As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TL = 5

Dim legarray() As String
Dim UpdateSiblings As Boolean
Dim NumSiblings As Long
Dim i As Long

On Error GoTo errorhandler

    mvarTL = vdata
    If mvarZZInit = 0 Then Exit Property
    
    'get the keys for our sibling leg drivetrains
    If mvarDatatype = LegDrivetrain Then
        legarray = Veh.KeyManager.GetCurrentLegDrivetrainKeys
        If legarray(1) <> "" Then
            UpdateSiblings = True
            NumSiblings = UBound(legarray)
        End If
    End If

    'update the siblings as well
    If UpdateSiblings Then
        For i = 1 To NumSiblings
            If Veh.Components(legarray(i)).TL <> vdata Then
                Veh.Components(legarray(i)).TL = vdata
                Veh.Components(legarray(i)).StatsUpdate
            End If
        Next
    End If
    
    GetMatrixIndex
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



Private Sub Class_Initialize()
 

' set the default properties
mvarCustom = False
TL = gVehicleTL
mvarRuggedized = False
mvarMotivePower = 60
End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Dim legarray() As String
Dim i As Long

Select Case mvarDatatype
    Case WheeledDrivetrain
        
    Case AllWheelDriveWheeledDrivetrain
        
    Case TrackedDrivetrain
        
    Case LegDrivetrain
       
        
        'if this leg drivetrain has been added AFTER other leg drivetrains have been added
        'then this ddrivetrains stats should default to existing stats
        legarray = Veh.KeyManager.GetCurrentLegDrivetrainKeys
        If legarray(1) <> "" Then
            mvarTL = Veh.Components(legarray(1)).TL
            mvarRuggedized = Veh.Components(legarray(1)).Ruggedized
            mvarMotivePower = Veh.Components(legarray(1)).MotivePower
            mvarDR = Veh.Components(legarray(1)).DR
            'must now update all the other legs statsupdate since adding this legdrivetrain and leg
            'results in changed average volume due to increased leg count
            For i = 1 To UBound(legarray)
                Veh.Components(legarray(i)).StatsUpdate
            Next
        End If
        
    Case FlexibodyDrivetrain
       
End Select

End Sub

Public Sub GetMatrixIndex()
Dim i As Integer
If mvarDatatype = 0 Then Exit Sub
mvarMatrixPos = 0 'init the counter
For i = 1 To UBound(DrivetrainMatrix())
    If DrivetrainMatrix(i).ID = mvarDatatype Then
        If DrivetrainMatrix(i).TL >= mvarTL Then
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

Dim TempWeight1 As Single
Dim TempWeight2 As Single
Dim NumLegs As Long
Dim CostMod As Integer
Dim legarray() As String
Dim i As Long
Dim QRugMod As Single 'combined quantity and ruggedized multipliers
Dim RugHitMod As Integer 'ruggedized hit point multiplier
Dim sPrint1 As String

mvarLocation = GetLocation

'set the ruggedized and quantity multipliers
If mvarRuggedized Then
    QRugMod = 1.5
    RugHitMod = 2
Else
    QRugMod = 1
    RugHitMod = 1
End If


' Find how many legs are on the vehicle and apply modifiers
legarray = Veh.KeyManager.GetCurrentLegKeys
NumLegs = UBound(legarray)

' a cost modifier is applied to Legs depending on how
' many are installed.. this is due to the need for
' extra stabilization and control systems
If NumLegs < 2 Then
    NumLegs = 1
    CostMod = 1 'default modifier
End If
If NumLegs = 3 Then
    CostMod = 2
ElseIf NumLegs = 2 Then
    CostMod = 4
Else
    CostMod = 1
End If
    

If mvarDatatype = LegDrivetrain Then
    ' NOTE: Here we must determine the correct >=5 formula.
    ' Vehicles 2dE assumes a single drivetrain.  GVD uses
    ' seperate motors so the >=5kW must take into account total
    ' number of legs
    If mvarMotivePower * NumLegs >= 5 Then
        TempWeight1 = DrivetrainMatrix(mvarMatrixPos).Weight2
        TempWeight2 = DrivetrainMatrix(mvarMatrixPos).Weight3
    Else
        TempWeight1 = DrivetrainMatrix(mvarMatrixPos).Weight1
        TempWeight2 = 0
    End If
    
    ' NOTE: must divide the TempWeight2 by the number of legs since GVD
    ' calcs all stats for each motor as seperate components and not as
    ' one single drivetrain
    TempWeight1 = (mvarMotivePower * TempWeight1) + (TempWeight2 / NumLegs)
Else

    If mvarMotivePower >= 5 Then
        TempWeight1 = DrivetrainMatrix(mvarMatrixPos).Weight2
        TempWeight2 = DrivetrainMatrix(mvarMatrixPos).Weight3
    Else
        TempWeight1 = DrivetrainMatrix(mvarMatrixPos).Weight1
        TempWeight2 = 0
    End If
    
    TempWeight1 = (mvarMotivePower * TempWeight1) + TempWeight2
End If



'get base stats
mvarWeight = TempWeight1
mvarCost = mvarWeight * DrivetrainMatrix(mvarMatrixPos).Cost * CostMod
mvarVolume = mvarWeight / DrivetrainMatrix(mvarMatrixPos).Volume
mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
mvarHitPoints = CalcComponentHitpoints(RugHitMod * mvarSurfaceArea)

'get finals
mvarWeight = Round(QRugMod * mvarWeight, 2)
mvarCost = Round(QRugMod * mvarCost, 2)
mvarVolume = Round(QRugMod * mvarVolume, 2)
mvarSurfaceArea = CalcSurfaceArea(mvarVolume)

mvarPowerReqt = Round(mvarMotivePower, 2)
     
'produce the print output
If mvarRuggedized Then
    sPrint1 = "ruggedized "
Else
    sPrint1 = ""
End If

mvarPrintOutput = "TL" & mvarTL & " " + Format(mvarMotivePower, p_sFormat) + " kW " + sPrint1 + mvarCustomDescription _
    + " (" + mvarLocation + ", HP " & mvarHitPoints & ", " + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + ", " + Format(mvarPowerReqt, p_sFormat) + " kW)." + mvarComment




End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub

Public Function LocationCheck() As Boolean
Dim TempCheck As Boolean
Dim otherdrivetrainarray() As String
Dim component As String
Dim i As Long
    
    'determine if the user is adding more drivetrains than rotors
    otherdrivetrainarray = Veh.KeyManager.GetCurrentOtherGroundDrivetrainKeys
    
    If mvarDatatype <> LegDrivetrain Then
        If otherdrivetrainarray(1) <> "" Then
            For i = 1 To UBound(otherdrivetrainarray)
                component = Veh.Components(otherdrivetrainarray(i)).Datatype
                If component = mvarDatatype Then
                    TempCheck = False
                    InfoPrint 1, "Only one of these types of drivetrains can be installed onto a Vehicle."
                    LocationCheck = TempCheck
                    Exit Function
                'need to check that allwheeldrivewheeledrivetrain checks regular wheeleddrivetrain
                ElseIf component = WheeledDrivetrain And mvarDatatype = AllWheelDriveWheeledDrivetrain Then
                    TempCheck = False
                    InfoPrint 1, "Only one of these types of drivetrains can be installed onto a Vehicle."
                    LocationCheck = TempCheck
                    Exit Function
                ElseIf component = AllWheelDriveWheeledDrivetrain And mvarDatatype = WheeledDrivetrain Then
                    TempCheck = False
                    InfoPrint 1, "Only one of these types of drivetrains can be installed onto a Vehicle."
                    LocationCheck = TempCheck
                    Exit Function
                End If
            Next
        End If
    End If
    
    'determine if the user is placing the component in a valid location
    Select Case mvarDatatype
    
    Case LegDrivetrain
        If Veh.Components(Parent).Datatype <> Leg Then
            InfoPrint 1, "Leg Drivetrains can only be placed in the Vehicle's Legs"
            TempCheck = False
        Else
            TempCheck = True
        End If
    Case Else
        If Veh.Components(Parent).Datatype <> Body Then
            InfoPrint 1, "Wheeled, Flexibody and Tracked Drivetrains must be placed in the Vehicle's Body"
            TempCheck = False
        Else
            TempCheck = True
        End If
    End Select
    
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


