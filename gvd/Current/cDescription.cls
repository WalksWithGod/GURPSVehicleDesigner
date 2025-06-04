VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements cIPersist

' description
Private mvarCategory As String
Private mvarSubCategory As String '<-- i think these categories can be deduced based on the primary performance system used and the subassemblies used (e.g. wings  or rotars or flexibody, etc)

Private m_sName As String     '<-- the name of the vehicle e.g. General Dynamics F4 Falcon Block 16E
Private m_sClassification As String ' this is sort of hte "purpose" or "role" of the vehicle.  e.g. Heavy lift utility helicopter, multi-role jet fighter, etc
Private m_sDescription As String
Private mvarHeader As String  '<-- would be used for copyright stuff?
Private mvarFooter As String  '<-- would be used for bibliography or footnotes

' TO HELL WITH THE IMAGE FILE.  Todo: maybe in another version, but not this one
'Private Const MAX_JPEG_SIZE_IN_KILOBYTES = 50
'Private m_sJPEGFilename As String


'Private mvarDetails As String  'pg 27.  These really should be properties of existing cIContainers or individual components.. in some cases, free without volume/weight/cost.
'Private mvarVision As String ' todo: vision is a property of cIContainer that is applicable whenever there are crew/passenger seats/stations/accomodations installed in them.  see page 25!
'                               further, its not for the user to just write anything, this needs to be a drop
'                               down list with the options "good, fair, poor, no view"

'Public Property Let Details(ByVal vdata As String)
'    mvarDetails = vdata
'End Property
'Public Property Get Details() As String
'    Details = mvarDetails
'End Property
' todo:delete this property
'Public Property Let Vision(ByVal vdata As String)
'    mvarVision = vdata
'End Property
'Public Property Get Vision() As String
'    Vision = mvarVision
'End Property

Private Sub Class_Initialize()
    'mvarCategory = ""
    'mvarSubCategory = ""
   
End Sub


Public Property Let Name(ByVal vdata As String)
    m_sName = vdata
End Property
Public Property Get Name() As String
    Name = m_sName
End Property
Public Property Let Classification(ByVal s As String)
    m_sClassification = s
End Property
Public Property Get Classification() As String
    Classification = m_sClassification
End Property
Public Property Let Description(ByVal s As String)
    m_sDescription = s
End Property
Public Property Get Description() As String
    Description = m_sDescription
End Property
Public Property Let Footer(ByVal vdata As String)
    mvarFooter = vdata
End Property
Public Property Get Footer() As String
    Footer = mvarFooter
End Property
Public Property Let Header(ByVal vdata As String)
    mvarHeader = vdata
End Property
Public Property Get Header() As String
    Header = mvarHeader
End Property
Public Property Let Category(ByVal vdata As String)
    mvarCategory = vdata
End Property
Public Property Get Category() As String
    Category = mvarCategory
End Property
Public Property Let SubCategory(ByVal vdata As String)
    mvarSubCategory = vdata
End Property
Public Property Get SubCategory() As String
    SubCategory = mvarSubCategory
End Property


'//cIPersist Interface
Private Property Get cIPersist_Classname() As String
End Property
Private Property Get cIPersist_GUID() As String
End Property

Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Long)
       m_sName = op.Load(XML_NODE_NAME)
       m_sClassification = op.Load("classification")
       m_sDescription = op.Load(XML_NODE_DESCRIPTION)
       mvarHeader = op.Load("header")
       mvarFooter = op.Load("footer")
       
End Sub
Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
End Sub
