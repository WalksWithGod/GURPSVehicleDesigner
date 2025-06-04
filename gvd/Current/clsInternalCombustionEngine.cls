VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInternalCombustionEngine"
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
Private mvarFuelConsumption As Single
Private mvarVolume As Double
Private mvarLocation As String
Private mvarParent As String
Private mvarKey As String
Private mvarDR As Long
Private mvarRuggedized As Boolean
Private mvarSurfaceArea As Double
Private mvarHitPoints As Double
Private mvarPowerConsumed As Single

Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarQuantity As Integer
Private mvarFuelType As String
Private mvarUnModifiedOutput As Single
Private mvarOutput As Single
Private mvarDesiredOutput As Single
Private mvarEndurance As Single
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


Public Property Let Endurance(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Endurance = 5
    mvarEndurance = vdata
End Property


Public Property Get Endurance() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Endurance
    Endurance = mvarEndurance
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

       If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Pod) Or (InstallPoint = Wing) Or (InstallPoint = equipmentPod) Then
            TempCheck = True
        Else
            InfoPrint 1, "Internal Combustion Engines must go in Body, Pod, equipment Pod or Wing."
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


Public Property Let UnModifiedOutput(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnModifiedOutput = 5
    mvarUnModifiedOutput = vdata
End Property


Public Property Get UnModifiedOutput() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnModifiedOutput
    UnModifiedOutput = mvarUnModifiedOutput
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



Public Property Let FuelType(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FuelType = 5
    mvarFuelType = vdata
End Property


Public Property Get FuelType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FuelType
    FuelType = mvarFuelType
End Property




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



Public Property Let FuelConsumption(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FuelConsumption = 5
    mvarFuelConsumption = vdata
End Property


Public Property Get FuelConsumption() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FuelConsumption
    FuelConsumption = mvarFuelConsumption
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

Public Property Let DesiredOutput(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DesiredOutput = 5
    mvarDesiredOutput = vdata
End Property

Public Property Get DesiredOutput() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DesiredOutput
    DesiredOutput = mvarDesiredOutput
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
ReDim mvarFuelStorageKeyChain(1)

' set the default properties
mvarCustom = False
TL = gVehicleTL
mvarRuggedized = False
mvarQuantity = 1
mvarUnModifiedOutput = 40
mvarOutput = 40
End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case HPCeramicEngine
       
        mvarFuelType = "aviation gas"
    Case TurboHPCeramicEngine
       
        mvarFuelType = "aviation gas"
    Case SuperHPCeramicEngine
       
        mvarFuelType = "aviation gas"
    Case HPGasolineEngine
       
        mvarFuelType = "gasoline"
    Case TurboHPGasolineEngine
        
        mvarFuelType = "gasoline"
    Case SuperHPGasolineEngine
        
        mvarFuelType = "gasoline"
    Case CeramicEngine
        
        mvarFuelType = "gasoline"
    Case TurboCeramicEngine
       
        mvarFuelType = "gasoline"
    Case SuperCeramicEngine
        
        mvarFuelType = "gasoline"
    Case TurboStandardDieselEngine
        
        mvarFuelType = "diesel"
    Case TurboHPDieselEngine
        
        mvarFuelType = "diesel"
    Case MarineDieselEngine
        
        mvarFuelType = "diesel"
    Case StandardDieselEngine
       
        mvarFuelType = "diesel"
    Case HPDieselEngine
        
        mvarFuelType = "diesel"
    Case GasolineEngine
        
        mvarFuelType = "gasoline"
    Case TurboGasolineEngine
        
        mvarFuelType = "gasoline"
    Case SuperGasolineEngine
        
        mvarFuelType = "gasoline"

    Case HydrogenCombustionEngine
        
        mvarFuelType = "hydrogen"

End Select


End Sub

Public Sub GetMatrixIndex()
Dim i As Integer
If mvarDatatype = 0 Then Exit Sub
mvarMatrixPos = 0 'init the counter
For i = 1 To UBound(CombustionEngineMatrix())
    If CombustionEngineMatrix(i).ID = mvarDatatype Then
        If CombustionEngineMatrix(i).TL >= mvarTL Then
            mvarMatrixPos = i
            Exit For
        Else
            mvarMatrixPos = i
        End If
    End If
Next
End Sub

Public Sub StatsUpdate()

Dim WeightMod As Single
Dim TrueWeight As Single
Dim TempCost As Single
Dim TempVolume As Single
Dim element As Object
Dim TempPowerConsumption As Single
Dim TempUnmodifiedOutput As Single
Dim i As Long
Dim QRugMod As Single 'combined quantity and ruggedized multipliers
Dim RugHitMod As Integer 'ruggedized hit point multiplier
Dim sPrint1 As String
Dim sPrint2 As String
Dim sPrintPlural As String
Dim sPrintPlural2 As String
Dim sPrintPlural3 As String
Dim sPrintPlural4 As String

    mvarZZInit = 1
    
    mvarLocation = GetLocation
    
    'set the ruggedized and quantity multipliers
    If mvarRuggedized Then
        QRugMod = 1.5 * mvarQuantity
        RugHitMod = 2
    Else
        QRugMod = 1 * mvarQuantity
        RugHitMod = 1
    End If
    
    
    'Dont run this StatsUpdate procedure after the component
    'has just initialized because the Datatype is still empty
    'which results in a 0 for Counter
    If mvarMatrixPos = 0 Then Exit Sub
    
    'determine if the weight is above or below 5kw and then make adjustments
    If UnModifiedOutput < 5 Then
        WeightMod = 0
        TrueWeight = CombustionEngineMatrix(mvarMatrixPos).Weight1
    Else
        WeightMod = CombustionEngineMatrix(mvarMatrixPos).Weight3
        TrueWeight = CombustionEngineMatrix(mvarMatrixPos).Weight2
    End If
    
    TrueWeight = (mvarUnModifiedOutput * TrueWeight) + WeightMod
    
    'Get Cost
    TempCost = TrueWeight * CombustionEngineMatrix(mvarMatrixPos).Cost
    
    'Get Volume
    TempVolume = TrueWeight / CombustionEngineMatrix(mvarMatrixPos).Volume
    
    ' Set Effective Power before any modifiers
    TempUnmodifiedOutput = mvarUnModifiedOutput
    mvarDesiredOutput = Round(TempUnmodifiedOutput, 2)
    
    'only propane fuel causes change in fuel consumption (i.e. alcohol conversion does not)
    If mvarFuelType = "propane" Then
        mvarFuelConsumption = Round(1.5 * CombustionEngineMatrix(mvarMatrixPos).FuelUsed * TempUnmodifiedOutput, 2)
        TrueWeight = TrueWeight * 1.1
        TempCost = TempCost * 1.1
        TempVolume = TempVolume * 1.1
    ElseIf (mvarFuelType = "ethanol") Or (mvarFuelType = "methanol") Then 'this is for multi fuels
        If (mvarDatatype = CeramicEngine) Or (mvarDatatype = TurboCeramicEngine) Or (mvarDatatype = SuperCeramicEngine) Then
            mvarFuelConsumption = Round(CombustionEngineMatrix(mvarMatrixPos).AlcoholFuelMod * CombustionEngineMatrix(mvarMatrixPos).FuelUsed * TempUnmodifiedOutput, 2)
        Else ' this is for alcohol conversions
            mvarFuelConsumption = Round(CombustionEngineMatrix(mvarMatrixPos).FuelUsed * TempUnmodifiedOutput, 2)
            mvarDesiredOutput = Round(TempUnmodifiedOutput * CombustionEngineMatrix(mvarMatrixPos).AlcoholModifier, 2)
        End If
    Else 'for all other engine types
        mvarFuelConsumption = Round(CombustionEngineMatrix(mvarMatrixPos).FuelUsed * TempUnmodifiedOutput, 2)
    End If
    
    ' check for a Nitrous Oxide booster child component and update effective motive power to 20%
    ' nitrous only boosts internal combustion engines other than Hydrogen!
    If mvarDatatype = HydrogenCombustionEngine Then
    Else
        For Each element In Veh.Components
            If element.Datatype = NitrousOxideBooster Then
               If element.LogicalParent = mvarKey Then
                    mvarDesiredOutput = Round(mvarDesiredOutput * 1.2, 2)
                    Exit For
                End If
            End If
        Next
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
    mvarFuelConsumption = mvarFuelConsumption * mvarQuantity
    
    
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
     
     sPrint2 = ", uses " + sPrintPlural4 + Format(mvarFuelConsumption, p_sFormat) + " gph " + mvarFuelType
    
    
     mvarPrintOutput = NumericToString(mvarQuantity) + " TL" & mvarTL & " " + Format(mvarDesiredOutput, p_sFormat) + " kW " + sPrint1 + mvarCustomDescription + sPrintPlural _
         + " (" + mvarLocation + ", HP " & mvarHitPoints & sPrintPlural3 + ", " + sPrintPlural4 + Format(mvarWeight, p_sFormat) + " lbs., " + Format(mvarVolume, p_sFormat) + " cf., " + "$" + Format(mvarCost, p_sFormat) + sPrint2 + ")." + mvarComment
    

End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub
