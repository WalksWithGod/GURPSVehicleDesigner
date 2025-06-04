VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Const MAGIC_NUMBER = &H948585  ' used to identify this is a template file




'also add things like
'Description, Refernce, PgNum, Comments
' icon can be moved here too
'this prog will also have an IDL component browser which can extract the header info during browsing

' will need to figure how to best store the matrix data and how to tell the program
'how to use it for stats calcs

Private m_Header As THeader

Private Sub Class_Initialize()
    m_Header.lngMagic = MAGIC_NUMBER
 
End Sub


Public Function setHeader(uHeader As CT_Version) As Long

    setHeader = uHeader
    
End Function

