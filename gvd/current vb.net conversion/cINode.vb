Option Strict Off
Option Explicit On
Public Interface _cINode
	Function getChildrenByClassName(ByRef Classname As String, ByRef hChilds() As Integer) As Boolean
	Function addChild(ByRef oChild As _cINode) As Boolean
	Function removeChild(ByVal lngIndex As Integer) As Boolean
	Function getChildIndexByHandle(ByVal h As Integer) As Integer
	Function getChild(ByVal lngIndex As Integer) As cINode
	ReadOnly Property childCount As Integer
	ReadOnly Property Classname As String
	ReadOnly Property Attributes As Integer
	 Property Handle As Integer
	 Property Parent As Integer
	 Property Name As String
	 Property Description As String
	 Property Image As String
End Interface
'UPGRADE_WARNING: Class instancing was changed to public. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ED41034B-3890-49FC-8076-BD6FC2F42A85"'
<System.Runtime.InteropServices.ProgId("cINode_NET.cINode")> Public Class cINode
	Implements _cINode
	
	
	Public Function getChildrenByClassName(ByRef Classname As String, ByRef hChilds() As Integer) As Boolean Implements _cINode.getChildrenByClassName
		' accepts a classname AND a long array to store the results
	End Function
	Public Function addChild(ByRef oChild As _cINode) As Boolean Implements _cINode.addChild
		' TODO: if this turns out to be either a memory hog or bottle neck
		' we can store LONG pointers to the objects using the IShellFolderEx_TLB
		' to
		
		'        Dim lptr as long
		'        Dim o As IShellFolderEx_TLB.IUnknown
		'        lptr = objptr(oChild)
		'        Set o = oChild
		'        o.AddRef
		'        Set o = Nothing
		'        Then add the lPtr to our array of m_ptrChildren() as long
		'        then in removeChild()  we MUST of course call o.Release to remove the reference count
		
		'        As a final optimization, we can add a m_lngAllocationSize variable to indicate the amount of
		'        free space for adding new elements to the array.  This way we dont have to redim every time
		'        but only when our m_lngchildcount > the m_lngArraySize size then we can just increase array size by
		'        m_lngArraySize = m_lngArraySize +  m_lngAllocationSize then redim preserve it m_lngArraySize
		'        whch could = 4 for for instance, so we are allocating 4 extra units each time rather than every time.
		'        This could definetly speed up load times when we are quickly adding lots of elements at a time.
		'        However, memory wise we dont gain anything since these objects are still kept in memory and storing
		'        a reference isnt that much more.  As for "load times"  the bottleneck there will most like
		'        be in reading the XML.
		
	End Function
	Public Function removeChild(ByVal lngIndex As Integer) As Boolean Implements _cINode.removeChild
		' todo: following up with the IUnknown code (see above) used in addChild()
		' as a further optimization, with childs stored as long ptrs, we can use
		' copymemory to shift the items.
	End Function
	Public Function getChildIndexByHandle(ByVal h As Integer) As Integer Implements _cINode.getChildIndexByHandle
	End Function
	Public Function getChild(ByVal lngIndex As Integer) As _cINode Implements _cINode.getChild
	End Function
	Public ReadOnly Property childCount() As Integer Implements _cINode.childCount
		Get
		End Get
	End Property
	Public ReadOnly Property Classname() As String Implements _cINode.Classname
		Get
		End Get
	End Property
	Public ReadOnly Property Attributes() As Integer Implements _cINode.Attributes
		Get
		End Get
	End Property
	Public Property Handle() As Integer Implements _cINode.Handle
		Get
		End Get
		Set(ByVal Value As Integer)
		End Set
	End Property
	Public Property Parent() As Integer Implements _cINode.Parent
		Get
		End Get
		Set(ByVal Value As Integer)
		End Set
	End Property
	Public Property Name() As String Implements _cINode.Name
		Get
		End Get
		Set(ByVal Value As String)
		End Set
	End Property
	Public Property Description() As String Implements _cINode.Description
		Get
		End Get
		Set(ByVal Value As String)
		End Set
	End Property
	Public Property Image() As String Implements _cINode.Image
		Get
		End Get
		Set(ByVal Value As String) ' this is actually the filepath of our image.  Note we use the filepath as the key value in the image control too.
		End Set
	End Property
End Class