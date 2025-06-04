Option Strict Off
Option Explicit On
Public Interface _cIContainer
	 Property VolumeExtra As Double
	 Property VolumeAccessSpace As Double
	 Property VolumeBattleSuit As Double
	ReadOnly Property ContainerAbbrev As String
	Function AcceptObject(ByRef oComponent As _cIComponent) As Integer
End Interface
<System.Runtime.InteropServices.ProgId("cIContainer_NET.cIContainer")> Public Class cIContainer
	Implements _cIContainer
	
	
	'//Unique to Containers
	' VolumeExtra covers both empty space and added volume
	Public Property VolumeExtra() As Double Implements _cIContainer.VolumeExtra
		Get
		End Get
		Set(ByVal Value As Double)
		End Set
	End Property
	Public Property VolumeAccessSpace() As Double Implements _cIContainer.VolumeAccessSpace
		Get
		End Get
		Set(ByVal Value As Double)
		End Set
	End Property
	Public Property VolumeBattleSuit() As Double Implements _cIContainer.VolumeBattleSuit
		Get
		End Get
		Set(ByVal Value As Double)
		End Set
	End Property
	Public ReadOnly Property ContainerAbbrev() As String Implements _cIContainer.ContainerAbbrev
		Get
		End Get
	End Property
	
	' todo: this is still a point of confusion.
	Public Function AcceptObject(ByRef oComponent As _cIComponent) As Integer Implements _cIContainer.AcceptObject
	End Function
End Class