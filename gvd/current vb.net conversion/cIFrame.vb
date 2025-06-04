Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cIFrame_NET.cIFrame")> Public Class cIFrame
	
	
	Public Property Robotic() As Boolean
		Get
		End Get
		Set(ByVal Value As Boolean)
		End Set
	End Property
	Public Property Biomechanical() As Boolean
		Get
		End Get
		Set(ByVal Value As Boolean)
		End Set
	End Property
	Public Property Responsive() As Boolean
		Get
		End Get
		Set(ByVal Value As Boolean)
		End Set
	End Property
	Public Property LivingMetal() As Boolean
		Get
		End Get
		Set(ByVal Value As Boolean)
		End Set
	End Property
	
	Public Property Compartmentalization() As Integer
		Get
		End Get
		Set(ByVal Value As Integer)
			'note: compartmentalization uses constants to represent none, basic and heavy
		End Set
	End Property
	Public Property SlopeLeft() As Integer
		Get
		End Get
		Set(ByVal Value As Integer)
			'note: slope uses constants to represent 0, 30 or 60
		End Set
	End Property
	Public Property SlopeRight() As Integer
		Get
		End Get
		Set(ByVal Value As Integer)
		End Set
	End Property
	Public Property SlopeFront() As Integer
		Get
		End Get
		Set(ByVal Value As Integer)
			'note: slope uses constants to represent 0, 30 or 60
		End Set
	End Property
	Public Property SlopeBack() As Integer
		Get
		End Get
		Set(ByVal Value As Integer)
			'note: slope uses constants to represent 0, 30 or 60
		End Set
	End Property
	Public Property Materials() As String
		Get
		End Get
		Set(ByVal Value As String)
			'todo: maybe this uses constants which the Def Editor will also be aware of
		End Set
	End Property
	Public Property FrameStrength() As String
		Get
		End Get
		Set(ByVal Value As String)
		End Set
	End Property
	
	'todo: is this correct place to have streamlining?  It would have to be set individually for each subassembly?
	Public Property StreamLining() As String
		Get
		End Get
		Set(ByVal Value As String)
		End Set
	End Property
End Class