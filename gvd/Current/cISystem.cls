VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cISystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function isFunctional() As Boolean
'todo: what about components like "beds" which aren't systems per se, but can be used by crew to "gain rest poitns" back?
    ' otherwise this just mimics damage amount.  Like a bed, you would
    ' only wonder if it was "functional" if it was heavily damaged.
    ' as for a system that uses power, isfunctional would look at both
    ' damage and power availability.  Need to consider this more
End Function
