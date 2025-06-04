VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWeaponLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarCost As Double
Private mvarParent As String
Private mvarKey As String
Private mvarDatatype As Integer
Private mvarDescription As String
Private mvarKeyChain As Variant

Public Property Let Description(ByVal vdata As String)
    mvarDescription = vdata
End Property

Public Property Get Description() As String
    Description = mvarDescription
End Property



Public Property Let KeyChain(ByVal vdata As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.KeyChain = 5
    mvarKeyChain = vdata
End Property

Public Property Get KeyChain() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.KeyChain
    KeyChain = mvarKeyChain

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


Public Function GetCurrentKeys() As String()
GetCurrentKeys = VariantArrayToStringArray(mvarKeyChain)
End Function

Public Sub AddKey(WeaponKey As String)
mvarKeyChain = mAddKey(KeyChain, WeaponKey)

'update the cost of the link depending on how many weapons are in it
If mvarKeyChain(1) <> "" Then mvarCost = 50 * UBound(mvarKeyChain) Else mvarCost = 0

End Sub

Public Sub RemoveKey(WeaponKey As String)
mvarKeyChain = mRemoveKey(mvarKeyChain, WeaponKey)
'update the cost of the link depending on how many weapons are in it
'todo: shouldn't be hardcoding cost
If mvarKeyChain(1) <> "" Then mvarCost = 50 * UBound(mvarKeyChain) Else mvarCost = 0
End Sub

Private Sub Class_Initialize()
ReDim mvarKeyChain(1)
End Sub


