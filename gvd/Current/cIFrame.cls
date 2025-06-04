VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
       

Public Property Let Robotic(ByVal b As Boolean)
End Property
Public Property Get Robotic() As Boolean
End Property
Public Property Let Biomechanical(ByVal b As Boolean)
End Property
Public Property Get Biomechanical() As Boolean
End Property
Public Property Let Responsive(ByVal b As Boolean)
End Property
Public Property Get Responsive() As Boolean
End Property
Public Property Let LivingMetal(ByVal b As Boolean)
End Property
Public Property Get LivingMetal() As Boolean
End Property

Public Property Let Compartmentalization(ByVal l As Long)
'note: compartmentalization uses constants to represent none, basic and heavy
End Property
Public Property Get Compartmentalization() As Long
End Property
Public Property Let SlopeLeft(ByVal l As Long)
'note: slope uses constants to represent 0, 30 or 60
End Property
Public Property Get SlopeLeft() As Long
End Property
Public Property Let SlopeRight(ByVal l As Long)
End Property
Public Property Get SlopeRight() As Long
End Property
Public Property Let SlopeFront(ByVal l As Long)
'note: slope uses constants to represent 0, 30 or 60
End Property
Public Property Get SlopeFront() As Long
End Property
Public Property Let SlopeBack(ByVal l As Long)
'note: slope uses constants to represent 0, 30 or 60
End Property
Public Property Get SlopeBack() As Long
End Property
Public Property Let Materials(ByRef s As String)
'todo: maybe this uses constants which the Def Editor will also be aware of
End Property
Public Property Get Materials() As String
End Property
Public Property Let FrameStrength(ByRef s As String)
End Property
Public Property Get FrameStrength() As String
End Property

'todo: is this correct place to have streamlining?  It would have to be set individually for each subassembly?
Public Property Let StreamLining(ByRef s As String)
End Property
Public Property Get StreamLining() As String
End Property
