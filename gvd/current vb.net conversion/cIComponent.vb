Option Strict Off
Option Explicit On
Public Interface _cIComponent
	 Property LogicalParent As Integer
	 Property SurfaceArea As Double
	 Property Cost As Double
	 Property Volume As Double
	 Property Weight As Double
	 Property HitPoints As Double
	 Property TL As Single
End Interface
<System.Runtime.InteropServices.ProgId("cIComponent_NET.cIComponent")> Public Class cIComponent
	Implements _cIComponent
	
	
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
	Public Property LogicalParent() As Integer Implements _cIComponent.LogicalParent
		Get
		End Get
		Set(ByVal Value As Integer)
		End Set
	End Property
	Public Property SurfaceArea() As Double Implements _cIComponent.SurfaceArea
		Get
		End Get
		Set(ByVal Value As Double)
		End Set
	End Property
	Public Property Cost() As Double Implements _cIComponent.Cost
		Get
		End Get
		Set(ByVal Value As Double)
		End Set
	End Property
	Public Property Volume() As Double Implements _cIComponent.Volume
		Get
		End Get
		Set(ByVal Value As Double)
		End Set
	End Property
	Public Property Weight() As Double Implements _cIComponent.Weight
		Get
		End Get
		Set(ByVal Value As Double)
		End Set
	End Property
	Public Property HitPoints() As Double Implements _cIComponent.HitPoints
		Get
		End Get
		Set(ByVal Value As Double)
		End Set
	End Property
	Public Property TL() As Single Implements _cIComponent.TL
		Get
		End Get
		Set(ByVal Value As Single)
		End Set
	End Property
End Class