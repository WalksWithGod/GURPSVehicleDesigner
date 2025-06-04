VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements cIPersist

Private m_sGUID As String
Private m_iMajor As Integer
Private m_iMinor As Integer
Private m_iRevision As Integer
Private m_dateCreated As Date
Private m_dateModified As Date
Private m_boolAutoIncrement As Boolean


Public Property Get GUID() As String
    GUID = m_sGUID
End Property
Public Property Let GUID(s As String)
    m_sGUID = s
End Property

Public Property Get Major() As Integer
    Major = m_iMajor
End Property
Public Property Let Major(i As Integer)
    m_iMajor = i
End Property
Public Property Get Minor() As Integer
    Minor = m_iMinor
End Property
Public Property Let Minor(i As Integer)
    m_iMinor = i
End Property
Public Property Get Revision() As Integer
    Revision = m_iRevision
End Property
Public Property Let Revision(i As Integer)
    m_iRevision = i
End Property
Public Property Get dateCreated() As Date
    dateCreated = m_dateCreated
End Property
Public Property Let dateCreated(D As Date)
    m_dateCreated = D
End Property
Public Property Get dateModified() As Date
    dateModified = m_dateModified
End Property
Public Property Let dateModified(D As Date)
    m_dateModified = D
End Property
Public Property Let AutoIncrement(ByVal f As Boolean)
    'todo: this should be more sophisticated.  We can check for changes to properties and that will increment
    ' revision, but adding/removing of components results in minor update.  *shrug*
    m_boolAutoIncrement = f
End Property
Public Property Get AutoIncrement() As Boolean
    AutoIncrement = m_boolAutoIncrement
End Property
'------------------
Private Property Get cIPersist_Classname() As String

End Property

Private Property Get cIPersist_GUID() As String

End Property

Private Sub cIPersist_LoadProperties(ByVal op As PersistenceManager.clsObjProperties, ByVal iMode As Long)
    m_sGUID = op.Load("guid")
    m_iMajor = op.Load("major")
    m_iMinor = op.Load("minor")
    m_iRevision = op.Load("revision")
    m_dateCreated = op.Load("created")
    m_dateModified = op.Load("modified")
    m_boolAutoIncrement = op.Load("autoincrement")
End Sub

Private Sub cIPersist_StoreProperties(ByVal op As PersistenceManager.clsObjProperties)

End Sub
