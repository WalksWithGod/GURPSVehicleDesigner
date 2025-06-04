VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComponentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


 Private m_lngCount As Long
 
 Public Function LoadComponent(sPath As String) As Long
 

    Set m_oXML = New cXML
    lngRet = m_oXML.Initialize(pavAuto)
    If lngRet = 0 Then
        InfoPrint 1, "Error:  Cannot load components.  Reason:  Could not initialize XML Parser.  Solution?:  Install Microsoft XML parser."
        Exit Function
    End If
    
    If m_oXML.OpenFromFile(sPath & sFileName, True) Then
    Dim o As Object
    'TODO: Best way to do this, is to also load the DEF file info first, then overwrite it with .CMP if
    ' applicable.  That way we are guaranteed to load all data since the DEF file is required to have
    ' EVERYTHING except things such as attributes output and print output stuff?
    Set o = m_oXML.FindNode("/GVD_SAVED_COMP/Component/Definition/Name")
    'sComponentName = oXML.ReadNode("/GVD_SAVED_COMP/Component/Definition/Name")
    'sComponentName = oxml.NodeCount
    
    
    'resortXML.XMLDocument.documentElement.childNodes.item(1).
    sIconPath = App.Path & m_oXML.ReadNode("/GVD_SAVED_COMP/Component/Definition/IconFile")
    
    hImgSmall = SHGetFileInfo(sIconPath, _
                0&, m_shinfo, Len(m_shinfo), _
                BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
                

End Function


Public Property Get ItemCount() As Long
    ItemCount = m_lngCount
End Property
