VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'//Unique to Containers
' VolumeExtra covers both empty space and added volume
Public Property Let VolumeExtra(ByVal vdata As Double)
End Property
Public Property Get VolumeExtra() As Double
End Property
Public Property Let VolumeAccessSpace(ByVal vdata As Double)
End Property
Public Property Get VolumeAccessSpace() As Double
End Property
Public Property Let VolumeBattleSuit(ByVal vdata As Double)
End Property
Public Property Get VolumeBattleSuit() As Double
End Property
Public Property Get ContainerAbbrev() As String
End Property

' todo: this is still a point of confusion.
Public Function AcceptObject(ByRef oComponent As cIComponent) As Long
End Function
