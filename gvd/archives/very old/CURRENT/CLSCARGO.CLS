VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarCost As Double 'local copy
Private mvarVolume As Double 'local copy
Private mvarSubType As String 'local copy
Private mvarImage As Integer 'local copy
Private mvarKey As String 'local copy
Private mvarLocation As String 'local copy
Private mvarSelectedImage As Integer 'local copy
Private mvarText As String 'local copy
Private mvarParent As String 'local copy
Private mvarCargoSpace As Double

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

Public Property Let Volume(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Volume = 5
    mvarVolume = vData
End Property


Public Property Get CargoSpace() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CargoSpace
    CargoSpace = mvarCargoSpace
End Property

Public Property Let CargoSpace(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CargoSpace = 5
    mvarCargoSpace = vData
' Calculate the cost if Hidden Cargo
If mvarSubType = "Hidden" Then
   mvarCost = 20 * mvarVolume
   mvarVolume = mvarCargoSpace
ElseIf mvarSubType = "Open" Then
   mvarVolume = mvarCargoSpace / 2
   mvarCost = 0
Else
    mvarCost = 0
    mvarVolume = mvarCargoSpace
End If
End Property


Public Property Get Volume() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Volume
    Volume = mvarVolume
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

Private Sub Class_Initialize()
' set the default properties

mvarCost = 0
mvarSubType = "Standard"
mvarVolume = 0
mvarCargoSpace = 0
End Sub


Private Sub Class_Terminate()
    'the class is being destroyed
    
End Sub


Public Function PropQuery(ByVal Row As Integer) As Integer
Const Combobox = 1
Const Textbox = 2
Const CommandButton = 3
' clear the Properties form's combobox
frmProperties.Combo1.Clear
    Select Case Row
    Case 0
        frmProperties.Combo1.AddItem "Standard"
        frmProperties.Combo1.AddItem "Hidden"
        frmProperties.Combo1.AddItem "Open"
        PropQuery = Combobox
    Case 1
        PropQuery = Textbox
    End Select
End Function

Public Sub PropUpdate(iRow As Integer, vNewValue As Variant)
Select Case iRow
    Case 0
        SubType = vNewValue
    Case 1
        CargoSpace = vNewValue

End Select
End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    colVehicle.Item(Parent).StatsUpdate
End If
End Sub


