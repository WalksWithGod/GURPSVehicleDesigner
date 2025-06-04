Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cAuthor_NET.cAuthor")> Public Class cAuthor
	Implements _cIPersist
	' NOTE:  This class is only to be used with saved vehicles.  .VEH (cmp) files will use this instead'
	'        to denote the author of entire vehicles.  For user created xml def's, we'll just use a
	'        XML comment <!-- >  to display any copyright/author info.  .VEH's need them though becuase on the
	'        vehicle factory site, its good to be able to sort by author
	' author specific
	
	
	Private m_sFirst As String
	Private m_sLast As String
	Private m_sMiddle As String
	Private m_sNick As String
	Private m_sEmail As String
	Private m_sURL As String
	
	
	Public Property firstName() As String
		Get
			firstName = m_sFirst
		End Get
		Set(ByVal Value As String)
			m_sFirst = Value
		End Set
	End Property
	
	Public Property lastName() As String
		Get
			lastName = m_sLast
		End Get
		Set(ByVal Value As String)
			m_sLast = Value
		End Set
	End Property
	
	Public Property middleName() As String
		Get
			middleName = m_sMiddle
		End Get
		Set(ByVal Value As String)
			m_sMiddle = Value
		End Set
	End Property
	
	Public Property nickName() As String
		Get
			nickName = m_sNick
		End Get
		Set(ByVal Value As String)
			m_sNick = Value
		End Set
	End Property
	
	Public Property email() As String
		Get
			email = m_sEmail
		End Get
		Set(ByVal Value As String)
			m_sEmail = Value
		End Set
	End Property
	
	Public Property url() As String
		Get
			url = m_sURL
		End Get
		Set(ByVal Value As String)
			m_sURL = Value
		End Set
	End Property
	'--------
	
	Private ReadOnly Property cIPersist_Classname() As String
		Get
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			cIPersist_Classname = TypeName(Me)
		End Get
	End Property
	
	Private ReadOnly Property cIPersist_GUID() As String
		Get
			
		End Get
	End Property
	
	Private Sub cIPersist_LoadProperties(ByVal op As PersistenceManager.clsObjProperties, ByVal iMode As Integer)
		Dim PersistenceManager As Object
		Dim i As Integer
		' todo: i dont believe this object needs to implement any interface cept cIPersist... need to think about this jsut a tiny bit more
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sFirst = op.Load("first")
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sMiddle = op.Load("middle")
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sLast = op.Load("last")
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sNick = op.Load("nick")
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sEmail = op.Load("email")
		'UPGRADE_WARNING: Couldn't resolve default property of object op.Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sURL = op.Load("url")
	End Sub
	
	Private Sub cIPersist_StoreProperties(ByVal op As PersistenceManager.clsObjProperties)
		Dim PersistenceManager As Object
		
	End Sub
	
	'---------
End Class