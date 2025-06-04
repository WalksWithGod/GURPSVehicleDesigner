VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoftware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarTL As Integer
Private mvarCost As Double
Private mvarComplexity As Long
Private mvarGigabytes As Single
Private mvarSkillPoints As Long
Private mvarLocation As String
Private mvarParent As String
Private mvarKey As String
Private mvarBonusSkillPoints As Long

Private mvarDatatype As Integer
Private mvarParentDatatype As Integer
Private mvarDescription As String
Private mvarCustomDescription As String
Private mvarCustom As Boolean
Private mvarQuantity As Long
'Note:  quantity is disabled for this.  If users request, i can re-enable so leave this property here for now

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

Public Function LocationCheck() As Boolean
Dim TempCheck As Boolean
        If TypeOf Veh.Components(Parent) Is clsComputer Then
            TempCheck = True
        Else
            InfoPrint 1, "Software must be installed in a Computer."
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


Public Property Let Quantity(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quantity = 5
    mvarQuantity = vdata
End Property


Public Property Get Quantity() As Long
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

Public Property Let SkillPoints(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SkillPoints = 5
    mvarSkillPoints = vdata
End Property


Public Property Get SkillPoints() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SkillPoints
    SkillPoints = mvarSkillPoints
End Property

Public Property Let BonusSkillPoints(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BonusSkillPoints = 5
    mvarBonusSkillPoints = vdata
End Property


Public Property Get BonusSkillPoints() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BonusSkillPoints
    BonusSkillPoints = mvarBonusSkillPoints
End Property

Public Property Let Gigabytes(ByVal vdata As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Gigabytes = 5
    mvarGigabytes = vdata
End Property


Public Property Get Gigabytes() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Gigabytes
    Gigabytes = mvarGigabytes
End Property



Public Property Let Complexity(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Complexity = 5
    mvarComplexity = vdata
End Property


Public Property Get Complexity() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Complexity
    Complexity = mvarComplexity
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

mvarQuantity = 1
mvarSkillPoints = 1
mvarGigabytes = 1
End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
End Sub

Public Sub Init()
Select Case mvarDatatype
    Case CartographySoftware
       
    Case ComputerNavigationSoftware
      
    Case DamageControlSoftware
        
    Case DatalinkSoftware
        
    Case FireDirectionSoftware
       
    Case DatabaseSoftware
        
    Case GunnerSoftware
       
    Case PersonalitySimulationSoftwareFull
        
    Case PersonalitySimulationLimited
       
    Case RobotSkillProgramsPhysical
      
    Case RobotSkillProgramsMental
        
    Case RoutineVehicleOperationSoftwarePilot
        
    Case RoutineVehicleOperationSoftwareOther
        
    Case TargetingSoftware
        
    Case TransmissionProfilingSoftware
        
    Case HoloventureProgram
        

End Select

End Sub

Public Sub GetMatrixIndex()
Dim i As Integer
If mvarDatatype = 0 Then Exit Sub
mvarMatrixPos = 0 'init the counter
For i = 1 To UBound(SoftwareMatrix())
    If SoftwareMatrix(i).ID = mvarDatatype Then
        If SoftwareMatrix(i).TL >= mvarTL Then
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
        
    Dim basecomplex As Long
    Dim TempCost As Single
    Dim sPrint1 As String
    
    mvarLocation = GetLocation
    
    'store the base complexity and skillpoint value
    basecomplex = SoftwareMatrix(mvarMatrixPos).Complexity
    mvarSkillPoints = SoftwareMatrix(mvarMatrixPos).BonusSkill
    
    Select Case mvarDatatype
       
       Case DatabaseSoftware
            TempCost = mvarGigabytes * SoftwareMatrix(mvarMatrixPos).Cost
       
       Case CartographySoftware, ComputerNavigationSoftware, _
            DatalinkSoftware, TransmissionProfilingSoftware, HoloventureProgram, _
            PersonalitySimulationSoftwareFull, PersonalitySimulationLimited, _
            RoutineVehicleOperationSoftwarePilot, RoutineVehicleOperationSoftwareOther
            
            TempCost = SoftwareMatrix(mvarMatrixPos).Cost
            
        Case FireDirectionSoftware, TargetingSoftware, DamageControlSoftware, _
            GunnerSoftware
            
            basecomplex = basecomplex + mvarBonusSkillPoints 'increase complexity by 1 point for each user added bonus skill
    
            TempCost = SoftwareMatrix(mvarMatrixPos).Cost
            If mvarBonusSkillPoints > 0 Then TempCost = TempCost * 2 ^ mvarBonusSkillPoints 'each +1 in complexity doubles cost
                  
            
        Case RobotSkillProgramsPhysical, RobotSkillProgramsMental
        
            'get the complexity based on the skill points
            If mvarBonusSkillPoints < 1 Then
                basecomplex = 1
            ElseIf mvarBonusSkillPoints = 1 Then
                basecomplex = 2
            ElseIf mvarBonusSkillPoints = 2 Then
                basecomplex = 3
            ElseIf mvarBonusSkillPoints <= 4 Then
                basecomplex = 4
            ElseIf mvarBonusSkillPoints <= 8 Then
                basecomplex = 5
            Else
                basecomplex = Fix((mvarBonusSkillPoints - 8) / 8)
                If basecomplex < (mvarBonusSkillPoints - 8 / 8) Then basecomplex = basecomplex + 1
                basecomplex = basecomplex + 5
            End If
            
            TempCost = SoftwareMatrix(mvarMatrixPos).Cost * mvarBonusSkillPoints
            If mvarBonusSkillPoints > 8 And mvarBonusSkillPoints < 20 Then
                TempCost = TempCost * 2.5
            ElseIf mvarSkillPoints > 20 Then
                TempCost = TempCost * 5
            End If
        
    End Select
        mvarSkillPoints = mvarSkillPoints + mvarBonusSkillPoints
        mvarComplexity = basecomplex
        mvarCost = Round(TempCost, 2)
    
    'produce the print output
    
    If mvarSkillPoints <> 0 Then
        sPrint1 = ", skill bonus +" + Format(mvarSkillPoints)
    ElseIf mvarDatatype = DatabaseSoftware Then
        sPrint1 = ", " & mvarGigabytes & " gig"
    End If
    
    mvarPrintOutput = " TL" & mvarTL & " " + mvarCustomDescription _
        + " (" + "$" + Format(mvarCost, p_sFormat) + ", complexity " + Format(mvarComplexity) + sPrint1 + ")." + mvarComment

End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    Veh.Components(Parent).StatsUpdate
End If
End Sub

