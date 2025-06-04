VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNuclearAntimatterandMassConversionReactors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarTL As Variant 'local copy
Private mvarComponentType As Variant 'local copy
Private mvarWeight As Variant 'local copy
Private mvarCost As Variant 'local copy
Private mvaryears As Variant 'local copy
Private mvarVolume As Variant 'local copy
Public Property Let ComponentType(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ComponentType = 5
    mvarComponentType = vData
End Property


Public Property Set ComponentType(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ComponentType = Form1
    Set mvarComponentType = vData
End Property


Public Property Get ComponentType() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ComponentType
    If IsObject(mvarComponentType) Then
        Set ComponentType = mvarComponentType
    Else
        ComponentType = mvarComponentType
    End If
End Property

Public Property Let Volume(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Volume = 5
    mvarVolume = vData
End Property


Public Property Set Volume(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Volume = Form1
    Set mvarVolume = vData
End Property


Public Property Get Volume() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Volume
    If IsObject(mvarVolume) Then
        Set Volume = mvarVolume
    Else
        Volume = mvarVolume
    End If
End Property



Public Property Let years(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.years = 5
    mvaryears = vData
End Property


Public Property Set years(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.years = Form1
    Set mvaryears = vData
End Property


Public Property Get years() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.years
    If IsObject(mvaryears) Then
        Set years = mvaryears
    Else
        years = mvaryears
    End If
End Property



Public Property Let Cost(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cost = 5
    mvarCost = vData
End Property


Public Property Set Cost(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.cost = Form1
    Set mvarCost = vData
End Property


Public Property Get Cost() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cost
    If IsObject(mvarCost) Then
        Set Cost = mvarCost
    Else
        Cost = mvarCost
    End If
End Property



Public Property Let Weight(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Weight = 5
    mvarWeight = vData
End Property


Public Property Set Weight(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Weight = Form1
    Set mvarWeight = vData
End Property


Public Property Get Weight() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Weight
    If IsObject(mvarWeight) Then
        Set Weight = mvarWeight
    Else
        Weight = mvarWeight
    End If
End Property






Public Property Let TL(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TL = 5
    mvarTL = vData
End Property


Public Property Set TL(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.TL = Form1
    Set mvarTL = vData
End Property


Public Property Get TL() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    If IsObject(mvarTL) Then
        Set TL = mvarTL
    Else
        TL = mvarTL
    End If
End Property



