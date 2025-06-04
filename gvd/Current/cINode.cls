VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cINode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function getChildrenByClassName(ByRef Classname As String, ByRef hChilds() As Long) As Boolean
    ' accepts a classname AND a long array to store the results
End Function
Public Function addChild(ByRef oChild As cINode) As Boolean
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
Public Function removeChild(ByVal lngIndex As Long) As Boolean
    ' todo: following up with the IUnknown code (see above) used in addChild()
    ' as a further optimization, with childs stored as long ptrs, we can use
    ' copymemory to shift the items.
End Function
Public Function getChildIndexByHandle(ByVal h As Long) As Long
End Function
Public Function getChild(ByVal lngIndex As Long) As cINode
End Function
Public Property Get childCount() As Long
End Property
Public Property Get Classname() As String
End Property
Public Property Get Attributes() As Long
End Property
Public Property Get Handle() As Long
End Property
Public Property Let Handle(ByVal l As Long)
End Property
Public Property Let Parent(ByVal l As Long)
End Property
Public Property Get Parent() As Long
End Property
Public Property Let Name(ByVal s As String)
End Property
Public Property Get Name() As String
End Property
Public Property Let Description(ByVal s As String)
End Property
Public Property Get Description() As String
End Property
Public Property Let Image(ByVal s As String) ' this is actually the filepath of our image.  Note we use the filepath as the key value in the image control too.
End Property
Public Property Get Image() As String
End Property

