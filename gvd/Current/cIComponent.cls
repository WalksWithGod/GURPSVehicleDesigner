VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'//Exposed Interface Properties and Methods


' logical parent is used for 2 things.
' 1) creating the print output with the correct location abbreviation
' 2) if a part of the hull is hit, this is needed to determine which
'    children are inthat particular subassembly and therefore suceptible to
'    have been hit or destroyed when the subassembly was hit
'  The only question is, is this the right place to have it?
'  I'm thinking INode might be better....

' Lets consider something like armor which implements cINode but not cIComponent.
' obviously armor can be graphed properly since its a cINode, but logical parent
' is not relevant to it... its parent is whatever its attached too.  (note that
' armor cant be attached to virtual space like "group" nodes.  I think im leaning
' towards having armor as a member object of cIComponent's.  It also makes it
' easier to find the armor objects protecting a particular component.  E.g.
' target itemX and you can access its armor methods like itemX.armor.hit(x damage, x face)

' So in this respect, logical parent is only relevant to a "component" and not
' a node.  Or rather, a component's parent is not necessarily the same as its
' node self parent.  So by this logic, leavin logical parent here is ok and preferred.
Public Property Let LogicalParent(ByVal l As Long)
End Property
Public Property Get LogicalParent() As Long
End Property
Public Property Let SurfaceArea(ByVal dbl As Double)
End Property
Public Property Get SurfaceArea() As Double
End Property
Public Property Let Cost(ByVal dbl As Double)
End Property
Public Property Get Cost() As Double
End Property
Public Property Let Volume(ByVal dbl As Double)
End Property
Public Property Get Volume() As Double
End Property
Public Property Let Weight(ByVal dbl As Double)
End Property
Public Property Get Weight() As Double
End Property
Public Property Get HitPoints() As Double
End Property
Public Property Let HitPoints(ByVal l As Double)
End Property
Public Property Let TL(ByVal sng As Single)
End Property
Public Property Get TL() As Single
End Property








