VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' NOTE:  This class is only to be used with saved vehicles.  .VEH (cmp) files will use this instead'
'        to denote the author of entire vehicles.  For user created xml def's, we'll just use a
'        XML comment <!-- >  to display any copyright/author info.  .VEH's need them though becuase on the
'        vehicle factory site, its good to be able to sort by author
' author specific

Implements cIPersist

Private m_sFirst As String
Private m_sLast As String
Private m_sMiddle As String
Private m_sNick As String
Private m_sEmail As String
Private m_sURL As String


Public Property Let firstName(s As String)
    m_sFirst = s
End Property
Public Property Get firstName() As String
    firstName = m_sFirst
End Property

Public Property Let lastName(s As String)
    m_sLast = s
End Property
Public Property Get lastName() As String
    lastName = m_sLast
End Property

Public Property Let middleName(s As String)
    m_sMiddle = s
End Property
Public Property Get middleName() As String
    middleName = m_sMiddle
End Property

Public Property Let nickName(s As String)
    m_sNick = s
End Property
Public Property Get nickName() As String
    nickName = m_sNick
End Property

Public Property Let email(s As String)
    m_sEmail = s
End Property
Public Property Get email() As String
    email = m_sEmail
End Property

Public Property Let url(s As String)
    m_sURL = s
End Property
Public Property Get url() As String
    url = m_sURL
End Property
'--------

Private Property Get cIPersist_Classname() As String
    cIPersist_Classname = TypeName(Me)
End Property

Private Property Get cIPersist_GUID() As String

End Property

Private Sub cIPersist_LoadProperties(ByVal op As PersistenceManager.clsObjProperties, ByVal iMode As Long)
    Dim i As Long
    ' todo: i dont believe this object needs to implement any interface cept cIPersist... need to think about this jsut a tiny bit more
    '
    m_sFirst = op.Load("first")
    m_sMiddle = op.Load("middle")
    m_sLast = op.Load("last")
    m_sNick = op.Load("nick")
    m_sEmail = op.Load("email")
    m_sURL = op.Load("url")
End Sub

Private Sub cIPersist_StoreProperties(ByVal op As PersistenceManager.clsObjProperties)

End Sub

'---------
