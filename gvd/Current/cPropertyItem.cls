VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPropertyItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements cIPersist


'//IPersist Interface
Dim m_sClassname As String
Dim m_lngDefID As Long

'//class data
'todo: NOTE, eventually, we'll be caching these properties when components are loaded.  No sense in loading a "surface area" property
      'every friggen time.  HOWEVER, that also means that if ICOmponent has surfaceArea and for instance IArmor has its own (since it doesnt implement ICOmponent)
      ' then the properties need DIFFERENT ID'S!!!  On the bright side, the errors should be evident when callbyname fails to find an interface
      
' todo: after we're finished 2.0, try setting some of these longs to integers and the strings to fixed length to see if we can speed up loading
' and conserve memory when items are loaded
Private m_lngInterfaceID As Long '<-- used to tell the Properties_Show which interface to use when calling "Callbyname" to access the variable
Private m_sCallByName As String  '<-- used to CallByName and get the "data" or "let" the user modify them
Private m_lngCallByType As Byte  ' 0 = property, 1 = function
Private m_sCaption As String
Private m_bReadOnly As Boolean '<-- a hacker could change a stat value by calling settext using the handle. Wouldn't matter if there wasnt a "Let" procedure to even accept it would it?
Private m_lngDataType As Long  ' this is the datatype for use by the proplist control
Private m_lngUnitType As Long  ' this is the unit type required by the conversion routines

Private m_sList() As String     ' if this was a list, it would contain list of selection options
Private m_iSubscript As Integer  ' in our m_Options() and m_UserInput() its the subscript for this particular item
Private m_sNotes As String     ' e.g. "Set this to the tech level you want this component to be..."
Private m_sPageInfo As String    ' "Veh2ed3rd-Pg1,2,4" ??? 'todo: used fixed length string here?


Private Sub Class_Initialize()
    m_lngInterfaceID = -1  '<-- we know it hasnt been set unless its loaded from XML and that will mean its at least set to 0 if successful
End Sub
Private Sub Class_Terminate()
    Erase m_sList
End Sub
'//Local Interface
Public Property Get InterfaceID() As Long
    InterfaceID = m_lngInterfaceID
End Property
Public Property Get CallByName() As String
    CallByName = m_sCallByName
End Property
Public Property Get CallByType() As Long
    CallByType = m_lngCallByType
End Property
Public Property Get Caption() As String
    Caption = m_sCaption
End Property
Public Property Get ReadOnly() As Boolean
    ReadOnly = m_bReadOnly
End Property
Public Property Get Datatype() As Long
    Datatype = m_lngDataType
End Property
Public Property Get UnitType() As Long
    UnitType = m_lngUnitType
End Property
Public Property Get List() As String()
    List = m_sList()
End Property
Public Function ListItem(ByVal index As Long) As String
    Debug.Assert index >= 0
    Debug.Assert index <= UBound(m_sList)
    ListItem = m_sList(index)
End Function
Public Property Get Subscript() As Long
    ' varname + subscript = value
    ' e.g.  Options + 3 = m_Options(3)
    ' this property returns the subscript so that we know the element position in the array
    Subscript = m_iSubscript
End Property

Public Function getSelectionIndexFromValue(ByRef s As String) As Long
    Dim I As Long
    For I = LBound(m_sList) To UBound(m_sList)
        If m_sList(I) = s Then
            getSelectionIndexFromValue = I
            Exit Function
        End If
    Next
    getSelectionIndexFromValue = -1
End Function
Public Property Get Notes() As String
    'todo: add code to truncate length of note?  Dont want users to enter an essay do we?
    Notes = m_sNotes
End Property
Public Property Get PageInfo() As String
    PageInfo = m_sPageInfo
End Property

Private Property Get cIPersist_Classname() As String
End Property
Private Property Get cIPersist_GUID() As String
End Property
Private Sub cIPersist_LoadProperties(ByVal op As clsObjProperties, ByVal iMode As Long)
    On Error GoTo err
    m_lngInterfaceID = op.Load("interfaceid")
    m_sCallByName = op.Load("callby")
    m_lngCallByType = op.Load("callbytype") '0 = property 1 = function
    m_sCaption = op.Load("caption")
    m_bReadOnly = op.Load("readonly")
    m_lngDataType = CLng(op.Load("datatype"))
    m_lngUnitType = CLng(op.Load("unittype"))
    m_sList = op.Load("Selections") ' note: not all properties have selections.  Todo: should check if lngDatatype <> wdList then skip trying to load selections
    m_iSubscript = op.Load("subscript")
    
    m_sNotes = op.Load("notes")
    m_sPageInfo = op.Load("pageinfo") 'todo: probably just include page info in the notes and delete this var?
       
    #If DEBUG_MODE Then
        InfoPrint 1, "BEGIN -- Property Attributes for '" & m_sCallByName & "'"
        InfoPrint 1, "m_lngInterfaceID = " & m_lngInterfaceID
        InfoPrint 1, "m_sCaption = " & m_sCaption
        InfoPrint 1, "m_bReadOnly = " & m_bReadOnly
        InfoPrint 1, "m_lngDataType = " & m_lngDataType
        InfoPrint 1, "m_lngUnitType = " & m_lngUnitType
        InfoPrint 1, "m_sNotes = " & m_sNotes
        InfoPrint 1, "m_sPageInfo = " & m_sPageInfo
    #End If
    Exit Sub
err:
    Debug.Print "cPropertyItem:cIPersist_LoadProperties() --  ERROR #" & err.Number & " " & err.Description
    Resume Next
End Sub

Private Sub cIPersist_StoreProperties(ByVal op As PersistenceManager.clsObjProperties)
End Sub

