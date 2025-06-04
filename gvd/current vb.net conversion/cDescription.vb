Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cDescription_NET.cDescription")> Public Class cDescription
	Implements _cIPersist
	
	' description
	Private mvarCategory As String
	Private mvarSubCategory As String '<-- i think these categories can be deduced based on the primary performance system used and the subassemblies used (e.g. wings  or rotars or flexibody, etc)
	
	Private m_sName As String '<-- the name of the vehicle e.g. General Dynamics F4 Falcon Block 16E
	Private m_sClassification As String ' this is sort of hte "purpose" or "role" of the vehicle.  e.g. Heavy lift utility helicopter, multi-role jet fighter, etc
	Private m_sDescription As String
	Private mvarHeader As String '<-- would be used for copyright stuff?
	Private mvarFooter As String '<-- would be used for bibliography or footnotes
	
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
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'mvarCategory = ""
		'mvarSubCategory = ""
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	Public Property Name() As String
		Get
			Name = m_sName
		End Get
		Set(ByVal Value As String)
			m_sName = Value
		End Set
	End Property
	Public Property Classification() As String
		Get
			Classification = m_sClassification
		End Get
		Set(ByVal Value As String)
			m_sClassification = Value
		End Set
	End Property
	Public Property Description() As String
		Get
			Description = m_sDescription
		End Get
		Set(ByVal Value As String)
			m_sDescription = Value
		End Set
	End Property
	Public Property Footer() As String
		Get
			Footer = mvarFooter
		End Get
		Set(ByVal Value As String)
			mvarFooter = Value
		End Set
	End Property
	Public Property Header() As String
		Get
			Header = mvarHeader
		End Get
		Set(ByVal Value As String)
			mvarHeader = Value
		End Set
	End Property
	Public Property Category() As String
		Get
			Category = mvarCategory
		End Get
		Set(ByVal Value As String)
			mvarCategory = Value
		End Set
	End Property
	Public Property SubCategory() As String
		Get
			SubCategory = mvarSubCategory
		End Get
		Set(ByVal Value As String)
			mvarSubCategory = Value
		End Set
	End Property
	
	
	'//cIPersist Interface
	Private ReadOnly Property cIPersist_Classname() As String
		Get
		End Get
	End Property
	Private ReadOnly Property cIPersist_GUID() As String
		Get
		End Get
	End Property
	
	Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Integer)
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sName = op.Load(XML_NODE_NAME)
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sClassification = op.Load("classification")
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sDescription = op.Load(XML_NODE_DESCRIPTION)
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarHeader = op.Load("header")
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarFooter = op.Load("footer")
		
	End Sub
	Private Sub cIPersist_StoreProperties(ByVal op As clsObjProperties)
	End Sub
End Class